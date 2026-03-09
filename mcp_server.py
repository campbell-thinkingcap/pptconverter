"""
MCP (Model Context Protocol) server for PPTX-to-SCORM conversion.
Exposes converter.py functions as MCP tools for AI assistants.

Usage:
    python mcp_server.py                  # stdio transport (Claude Desktop)
    python mcp_server.py --transport sse  # SSE transport (remote)
"""
import base64
import hashlib
import json
import os
import secrets
import tempfile
import time
from dataclasses import dataclass

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


@dataclass
class AuthCode:
    """Auth code object with attributes the MCP SDK token handler expects."""
    client_id: str
    code_challenge: str
    redirect_uri: str
    redirect_uri_provided_explicitly: bool
    scopes: list[str]
    expires_at: float


class SimpleOAuthProvider:
    """Minimal OAuth 2.1 provider for MCP auth.

    Auto-approves all authorization requests (no login page).
    Stores clients, codes, and tokens in memory.
    """

    def __init__(self):
        self.clients: dict[str, OAuthClientInformationFull] = {}
        self.auth_codes: dict[str, AuthCode] = {}
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
        self.auth_codes[code] = AuthCode(
            client_id=client.client_id,
            code_challenge=params.code_challenge,
            redirect_uri=str(params.redirect_uri),
            redirect_uri_provided_explicitly=params.redirect_uri_provided_explicitly,
            scopes=params.scopes or [],
            expires_at=time.time() + 300,
        )
        redirect = f"{params.redirect_uri}?code={code}"
        if params.state:
            redirect += f"&state={params.state}"
        return redirect

    async def load_authorization_code(
        self, client: OAuthClientInformationFull, authorization_code: str
    ) -> AuthCode | None:
        auth_code = self.auth_codes.get(authorization_code)
        if not auth_code or auth_code.client_id != client.client_id:
            return None
        if time.time() > auth_code.expires_at:
            self.auth_codes.pop(authorization_code, None)
            return None
        return auth_code

    async def exchange_authorization_code(
        self, client: OAuthClientInformationFull, authorization_code: AuthCode
    ) -> OAuthToken:
        # Remove the code (one-time use) — find it by matching the object
        code_str = None
        for k, v in self.auth_codes.items():
            if v is authorization_code:
                code_str = k
                break
        if code_str:
            self.auth_codes.pop(code_str, None)

        access_tok = secrets.token_urlsafe(48)
        refresh_tok = secrets.token_urlsafe(48)
        expires_in = 3600

        self.access_tokens[access_tok] = AccessToken(
            token=access_tok,
            client_id=client.client_id,
            scopes=authorization_code.scopes,
            expires_at=int(time.time()) + expires_in,
        )
        self.refresh_tokens[refresh_tok] = {
            "client_id": client.client_id,
            "scopes": authorization_code.scopes,
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
# Shared file staging registry
# ─────────────────────────────────────────────

# In-memory registry: id → {"path": str, "filename": str, "created": float}
# Shared between MCP tools and FastAPI download endpoint.
staged_files: dict[str, dict] = {}

UPLOAD_DIR = tempfile.mkdtemp(prefix="mcp_uploads_")


def _save_base64_file(file_content_base64: str, filename: str) -> str:
    """Decode base64 content, save to temp file, return path."""
    try:
        raw = base64.b64decode(file_content_base64)
    except Exception:
        raise ValueError("Invalid base64 content. Ensure the file is base64-encoded.")

    ext = os.path.splitext(filename)[1].lower() if filename else ".pptx"
    if ext not in (".ppt", ".pptx"):
        raise ValueError(f"Unsupported file type '{ext}'. Only .ppt and .pptx are accepted.")

    file_id = hashlib.sha256(os.urandom(32)).hexdigest()[:16]
    dest = os.path.join(UPLOAD_DIR, f"{file_id}{ext}")
    with open(dest, "wb") as f:
        f.write(raw)

    staged_files[file_id] = {
        "path": dest,
        "filename": filename or f"presentation{ext}",
        "created": time.time(),
    }
    return dest


# ─────────────────────────────────────────────
# Tools
# ─────────────────────────────────────────────

@mcp.tool()
def convert_presentation(
    file_content_base64: str,
    filename: str = "presentation.pptx",
    course_title: str = "",
    display_mode: str = "paginated",
    theme: str = "light",
) -> str:
    """Convert a PPTX/PPT file to a SCORM 2004 3rd Edition ZIP package.

    Send the file content as base64 and get back a download URL for the SCORM ZIP.

    Args:
        file_content_base64: The .pptx or .ppt file content, base64-encoded.
        filename: Original filename (e.g. "My Slides.pptx"). Used for the output name.
        course_title: Optional override for the course title.
        display_mode: 'paginated' (one slide at a time) or 'continuous' (scrollable view). Default: 'paginated'.
        theme: 'light' or 'dark' color scheme for the player. Default: 'light'.

    Returns:
        JSON with download_url for the SCORM ZIP package.
    """
    try:
        pptx_path = _save_base64_file(file_content_base64, filename)
    except ValueError as e:
        return json.dumps({"error": str(e)})

    output_dir = tempfile.mkdtemp(prefix="scorm_")
    title = course_title if course_title else None

    try:
        mode = display_mode if display_mode in ("paginated", "continuous") else "paginated"
        selected_theme = theme if theme in ("light", "dark") else "light"
        zip_path = converter.convert_pptx_to_scorm(pptx_path, output_dir, title, mode, selected_theme)
        zip_name = os.path.basename(zip_path)
        zip_size = os.path.getsize(zip_path)

        # Stage the result for download
        result_id = hashlib.sha256(os.urandom(32)).hexdigest()[:16]
        staged_files[result_id] = {
            "path": zip_path,
            "filename": zip_name,
            "created": time.time(),
        }

        download_url = f"{SERVER_URL}/api/download/{result_id}"
        return json.dumps({
            "filename": zip_name,
            "download_url": download_url,
            "file_size_bytes": zip_size,
            "message": f"SCORM package ready. Download: {download_url}",
        })
    except Exception as e:
        return json.dumps({"error": f"Conversion failed: {e}"})


@mcp.tool()
def get_presentation_info(
    file_content_base64: str,
    filename: str = "presentation.pptx",
) -> str:
    """Get metadata from a PPTX file without converting it.

    Send the file content as base64 to preview slide count and titles.

    Args:
        file_content_base64: The .pptx or .ppt file content, base64-encoded.
        filename: Original filename (e.g. "My Slides.pptx").

    Returns:
        JSON with title, slide_count, and slide_titles.
    """
    try:
        pptx_path = _save_base64_file(file_content_base64, filename)
    except ValueError as e:
        return json.dumps({"error": str(e)})

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


if __name__ == "__main__":
    mcp.run(transport="stdio")
