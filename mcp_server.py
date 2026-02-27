"""
MCP (Model Context Protocol) server for PPTX-to-SCORM conversion.
Exposes converter.py functions as MCP tools for AI assistants.

Usage:
    python mcp_server.py                  # stdio transport (Claude Desktop)
    python mcp_server.py --transport sse  # SSE transport (remote)
"""
import base64
import json
import os
import secrets
import tempfile
import time

from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings
from mcp.server.auth.settings import AuthSettings, ClientRegistrationOptions, RevocationOptions
from mcp.server.auth.provider import AccessToken
from mcp.shared.auth import OAuthClientInformationFull, OAuthToken

import converter

# ─────────────────────────────────────────────
# OAuth Provider (auto-approve, in-memory)
# ─────────────────────────────────────────────

SERVER_URL = os.environ.get(
    "SERVER_URL",
    "https://pptconverter.calmbay-775ec59e.eastus.azurecontainerapps.io",
)


class SimpleOAuthProvider:
    """Minimal OAuth 2.1 provider for MCP auth.

    Auto-approves all authorization requests (no login page).
    Stores clients, codes, and tokens in memory.
    """

    def __init__(self):
        self.clients: dict[str, OAuthClientInformationFull] = {}
        self.auth_codes: dict[str, dict] = {}
        self.access_tokens: dict[str, AccessToken] = {}
        self.refresh_tokens: dict[str, dict] = {}

    async def get_client(self, client_id: str) -> OAuthClientInformationFull | None:
        return self.clients.get(client_id)

    async def register_client(
        self, client_info: OAuthClientInformationFull
    ) -> None:
        self.clients[client_info.client_id] = client_info

    async def authorize(
        self, client: OAuthClientInformationFull, params
    ) -> str:
        """Auto-approve: generate an auth code and redirect immediately."""
        code = secrets.token_urlsafe(48)
        self.auth_codes[code] = {
            "client_id": client.client_id,
            "code_challenge": params.code_challenge,
            "redirect_uri": str(params.redirect_uri),
            "scopes": params.scopes or [],
            "expires": time.time() + 300,
        }
        redirect = f"{params.redirect_uri}?code={code}"
        if params.state:
            redirect += f"&state={params.state}"
        return redirect

    async def load_authorization_code(
        self, client: OAuthClientInformationFull, authorization_code: str
    ) -> str | None:
        data = self.auth_codes.get(authorization_code)
        if not data or data["client_id"] != client.client_id:
            return None
        if time.time() > data["expires"]:
            self.auth_codes.pop(authorization_code, None)
            return None
        return authorization_code

    async def exchange_authorization_code(
        self, client: OAuthClientInformationFull, authorization_code: str
    ) -> OAuthToken:
        data = self.auth_codes.pop(authorization_code, None)
        if not data:
            raise ValueError("Invalid authorization code")

        access_tok = secrets.token_urlsafe(48)
        refresh_tok = secrets.token_urlsafe(48)
        expires_in = 3600

        self.access_tokens[access_tok] = AccessToken(
            token=access_tok,
            client_id=client.client_id,
            scopes=data.get("scopes", []),
            expires_at=int(time.time()) + expires_in,
        )
        self.refresh_tokens[refresh_tok] = {
            "client_id": client.client_id,
            "scopes": data.get("scopes", []),
        }
        return OAuthToken(
            access_token=access_tok,
            token_type="Bearer",
            expires_in=expires_in,
            refresh_token=refresh_tok,
        )

    async def load_refresh_token(
        self, client: OAuthClientInformationFull, refresh_token: str
    ) -> str | None:
        data = self.refresh_tokens.get(refresh_token)
        if not data or data["client_id"] != client.client_id:
            return None
        return refresh_token

    async def exchange_refresh_token(
        self,
        client: OAuthClientInformationFull,
        refresh_token: str,
        scopes: list[str],
    ) -> OAuthToken:
        data = self.refresh_tokens.pop(refresh_token, None)
        if not data:
            raise ValueError("Invalid refresh token")

        new_access = secrets.token_urlsafe(48)
        new_refresh = secrets.token_urlsafe(48)
        expires_in = 3600
        effective_scopes = scopes or data.get("scopes", [])

        self.access_tokens[new_access] = AccessToken(
            token=new_access,
            client_id=client.client_id,
            scopes=effective_scopes,
            expires_at=int(time.time()) + expires_in,
        )
        self.refresh_tokens[new_refresh] = {
            "client_id": client.client_id,
            "scopes": effective_scopes,
        }
        return OAuthToken(
            access_token=new_access,
            token_type="Bearer",
            expires_in=expires_in,
            refresh_token=new_refresh,
        )

    async def load_access_token(self, token: str) -> AccessToken | None:
        at = self.access_tokens.get(token)
        if not at:
            return None
        if at.expires_at and time.time() > at.expires_at:
            self.access_tokens.pop(token, None)
            return None
        return at

    async def revoke_token(self, token) -> None:
        tok_str = token.token if hasattr(token, "token") else str(token)
        self.access_tokens.pop(tok_str, None)
        self.refresh_tokens.pop(tok_str, None)


# ─────────────────────────────────────────────
# MCP Server
# ─────────────────────────────────────────────

mcp = FastMCP(
    "pptconverter",
    transport_security=TransportSecuritySettings(
        enable_dns_rebinding_protection=False,
    ),
    auth_server_provider=SimpleOAuthProvider(),
    auth=AuthSettings(
        issuer_url=SERVER_URL,
        resource_server_url=SERVER_URL,
        client_registration_options=ClientRegistrationOptions(
            enabled=True,
            valid_scopes=[],
            default_scopes=[],
        ),
        revocation_options=RevocationOptions(enabled=True),
    ),
)


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def _save_upload(file_content_base64: str, filename: str) -> str:
    """Decode base64 file content, write to a temp file, return the path."""
    ext = os.path.splitext(filename)[1].lower() if filename else ".pptx"
    if ext not in (".pptx", ".ppt"):
        ext = ".pptx"
    tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False, prefix="mcp_upload_")
    tmp.write(base64.b64decode(file_content_base64))
    tmp.close()
    return tmp.name


# ─────────────────────────────────────────────
# Tools
# ─────────────────────────────────────────────

@mcp.tool()
def convert_presentation(file_content_base64: str, filename: str = "presentation.pptx", course_title: str = "") -> str:
    """Convert a PPTX/PPT file to a SCORM 2004 3rd Edition ZIP package.

    The file must be provided as base64-encoded content. Read the file first,
    then pass its base64-encoded bytes here.

    Args:
        file_content_base64: Base64-encoded content of the .pptx or .ppt file.
        filename: Original filename (used for title fallback). Default: presentation.pptx
        course_title: Optional override for the course title.

    Returns:
        Base64-encoded SCORM ZIP file content that can be saved to disk.
    """
    try:
        pptx_path = _save_upload(file_content_base64, filename)
    except Exception as e:
        return f"Error: Invalid base64 content: {e}"

    output_dir = tempfile.mkdtemp(prefix="scorm_")
    title = course_title if course_title else None

    try:
        zip_path = converter.convert_pptx_to_scorm(pptx_path, output_dir, title)

        with open(zip_path, "rb") as f:
            zip_b64 = base64.b64encode(f.read()).decode("ascii")

        # Suggest a filename for saving
        zip_name = os.path.basename(zip_path)
        return json.dumps({
            "filename": zip_name,
            "content_base64": zip_b64,
            "message": f"SCORM package '{zip_name}' created. Save the content_base64 field as a .zip file.",
        })
    except Exception as e:
        return f"Error during conversion: {e}"
    finally:
        os.unlink(pptx_path)


@mcp.tool()
def get_presentation_info(file_content_base64: str, filename: str = "presentation.pptx") -> str:
    """Get metadata from a PPTX file without converting it.

    The file must be provided as base64-encoded content.

    Args:
        file_content_base64: Base64-encoded content of the .pptx or .ppt file.
        filename: Original filename. Default: presentation.pptx

    Returns:
        JSON string with title, slide_count, and slide_titles.
    """
    try:
        pptx_path = _save_upload(file_content_base64, filename)
    except Exception as e:
        return json.dumps({"error": f"Invalid base64 content: {e}"})

    try:
        title = converter._get_presentation_title(pptx_path)
        titles = converter._get_slide_titles(pptx_path)
        return json.dumps({
            "title": title,
            "slide_count": len(titles),
            "slide_titles": titles,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
    finally:
        os.unlink(pptx_path)


if __name__ == "__main__":
    mcp.run(transport="stdio")
