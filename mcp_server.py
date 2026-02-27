"""
MCP (Model Context Protocol) server for PPTX-to-SCORM conversion.
Exposes converter.py functions as MCP tools for AI assistants.

Usage:
    python mcp_server.py                  # stdio transport (Claude Desktop)
    python mcp_server.py --transport sse  # SSE transport (remote)
"""
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
# Shared file staging registry (populated by main.py /api/upload)
# ─────────────────────────────────────────────

# Imported by main.py to register uploaded files; read by MCP tools.
staged_files: dict[str, dict] = {}


def _get_staged_path(file_id: str) -> tuple[str, str]:
    """Look up a staged file by ID. Returns (path, filename) or raises."""
    info = staged_files.get(file_id)
    if not info or not os.path.exists(info["path"]):
        raise FileNotFoundError(
            f"file_id '{file_id}' not found. Upload the file first via POST /api/upload."
        )
    return info["path"], info["filename"]


# ─────────────────────────────────────────────
# Tools
# ─────────────────────────────────────────────

@mcp.tool()
def convert_presentation(file_id: str, course_title: str = "") -> str:
    """Convert an uploaded PPTX/PPT file to a SCORM 2004 3rd Edition ZIP package.

    WORKFLOW:
    1. First upload the file via POST {server_url}/api/upload → returns a file_id
    2. Then call this tool with the file_id
    3. The result includes a download_url for the SCORM ZIP

    Args:
        file_id: The file_id returned by POST /api/upload.
        course_title: Optional override for the course title.

    Returns:
        JSON with download_url, filename, and slide_count.
    """
    try:
        pptx_path, original_name = _get_staged_path(file_id)
    except FileNotFoundError as e:
        return json.dumps({"error": str(e)})

    output_dir = tempfile.mkdtemp(prefix="scorm_")
    title = course_title if course_title else None

    try:
        zip_path = converter.convert_pptx_to_scorm(pptx_path, output_dir, title)
        zip_name = os.path.basename(zip_path)
        zip_size = os.path.getsize(zip_path)

        # Stage the result for download
        import hashlib
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
def get_presentation_info(file_id: str) -> str:
    """Get metadata from an uploaded PPTX file without converting it.

    Upload the file first via POST {server_url}/api/upload, then call this with the file_id.

    Args:
        file_id: The file_id returned by POST /api/upload.

    Returns:
        JSON with title, slide_count, and slide_titles.
    """
    try:
        pptx_path, _ = _get_staged_path(file_id)
    except FileNotFoundError as e:
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
