"""
PPTX → SCORM Conversion API
FastAPI application with OpenAPI 3.0 documentation.
"""
import os
import time
import tempfile
import hashlib
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Annotated, Optional

from fastapi import FastAPI, File, Form, UploadFile, HTTPException, BackgroundTasks, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.security import APIKeyHeader
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from google.oauth2 import id_token
from google.auth.transport import requests as google_requests

from converter import convert_pptx_to_scorm
from mcp_server import mcp as mcp_server, staged_files

# ─────────────────────────────────────────────
# Auth Configuration
# ─────────────────────────────────────────────

API_KEY = os.environ.get("API_KEY", "")
GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID", "")
ALLOWED_DOMAINS = [d.strip().lower() for d in os.environ.get("ALLOWED_DOMAINS", "").split(",") if d.strip()]

_api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)
_google_token_header = APIKeyHeader(name="X-Google-Token", auto_error=False)


def _verify_google_token(token: str) -> dict:
    """Validate a Google ID token. Returns claims dict or raises HTTPException."""
    try:
        claims = id_token.verify_oauth2_token(
            token,
            google_requests.Request(),
            GOOGLE_CLIENT_ID,
        )
    except ValueError as exc:
        raise HTTPException(status_code=401, detail=f"Invalid Google token: {exc}")

    email = claims.get("email", "")
    if not claims.get("email_verified"):
        raise HTTPException(status_code=401, detail="Google account email not verified.")

    if ALLOWED_DOMAINS:
        domain = email.split("@")[-1].lower()
        if domain not in ALLOWED_DOMAINS:
            raise HTTPException(status_code=403, detail=f"Email domain '{domain}' is not allowed.")

    return claims


async def verify_auth(
    api_key: str = Depends(_api_key_header),
    google_token: str = Depends(_google_token_header),
):
    """
    Accept either X-API-Key (programmatic) or X-Google-Token (browser UI).
    If neither auth method is configured, all requests pass through.
    """
    has_oauth = bool(GOOGLE_CLIENT_ID)
    has_api_key = bool(API_KEY)

    if not has_oauth and not has_api_key:
        return

    if api_key and has_api_key:
        if api_key == API_KEY:
            return
        raise HTTPException(status_code=401, detail="Invalid API key.")

    if google_token and has_oauth:
        _verify_google_token(google_token)
        return

    # No credentials provided — require auth only when API_KEY is configured.
    # When only Google OAuth is set, API calls without a token pass through
    # (Google sign-in is enforced by the browser UI, not the API).
    if has_api_key:
        raise HTTPException(status_code=401, detail="Invalid or missing API key.")

# ─────────────────────────────────────────────
# MCP Session Manager + App Lifespan
# ─────────────────────────────────────────────

# Build the MCP Starlette app (includes OAuth + MCP routes)
_mcp_app = mcp_server.streamable_http_app()

# Paths handled by the MCP Starlette app (OAuth + MCP endpoint)
_MCP_PATHS = frozenset({
    "/mcp",
    "/authorize",
    "/token",
    "/register",
    "/revoke",
    "/.well-known/oauth-authorization-server",
    "/.well-known/oauth-protected-resource",
})

@asynccontextmanager
async def _lifespan(app):
    """Start the MCP session manager alongside the FastAPI app."""
    async with mcp_server.session_manager.run():
        yield

# ─────────────────────────────────────────────
# App Setup
# ─────────────────────────────────────────────

app = FastAPI(
    title="PPTX → SCORM Conversion API",
    description="""
## Convert PowerPoint presentations to SCORM 2004 e-learning packages.

### Features
- Supports **.ppt** and **.pptx** formats
- High-fidelity slide rendering via LibreOffice
- SCORM 2004 3rd Edition compliant (`cmi.progress_measure`)
- Each slide becomes a page; completion tracked as % of slides viewed
- Built-in **Vue 3** player — no additional software needed
- Returns a ready-to-upload **.zip** file for any SCORM LMS

### Workflow
1. `POST /api/convert` with your .pptx file
2. Download the returned `.zip`
3. Upload to your LMS (Moodle, Canvas, TalentLMS, etc.)

### SCORM Data Model
| CMI Element | Behavior |
|---|---|
| `cmi.progress_measure` | 0.0 → 1.0 as slides are viewed |
| `cmi.completion_status` | `incomplete` → `completed` when all slides seen |
    | `cmi.success_status` | `passed` when completed |
    """,
    version="1.0.0",
    contact={"name": "PPTX SCORM API", "url": "https://github.com/"},
    license_info={"name": "MIT"},
    openapi_tags=[
        {"name": "conversion", "description": "Convert presentations to SCORM packages"},
        {"name": "health", "description": "Service health and metadata"},
    ],
    lifespan=_lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ─────────────────────────────────────────────
# ASGI dispatcher: MCP paths → _mcp_app, everything else → FastAPI
# ─────────────────────────────────────────────


# Temp storage (in production, use S3/blob storage)
OUTPUT_DIR = tempfile.mkdtemp(prefix="scorm_output_")
UPLOAD_DIR = tempfile.mkdtemp(prefix="scorm_uploads_")

# staged_files is imported from mcp_server (shared between FastAPI and MCP tools)


# ─────────────────────────────────────────────
# Response Models
# ─────────────────────────────────────────────

class ConversionInfo(BaseModel):
    filename: str
    slide_count: int
    file_size_bytes: int
    file_size_mb: float
    download_url: str
    generated_at: float

class HealthResponse(BaseModel):
    status: str
    version: str
    python_pptx_available: bool


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

ALLOWED_EXTENSIONS = {".ppt", ".pptx", ".pdf", ".doc", ".docx"}

def validate_upload(file: UploadFile) -> None:
    ext = Path(file.filename or "").suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=422,
            detail=f"Unsupported file type '{ext}'. Only .ppt, .pptx, .pdf, .doc, and .docx files are accepted."
        )

def cleanup_file(path: str) -> None:
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


# ─────────────────────────────────────────────
# React Frontend (static build from frontend/dist)
# ─────────────────────────────────────────────

_FRONTEND_DIST = Path(__file__).parent / "frontend" / "dist"
_FRONTEND_INDEX = _FRONTEND_DIST / "index.html"

if _FRONTEND_DIST.is_dir():
    app.mount("/assets", StaticFiles(directory=str(_FRONTEND_DIST / "assets")), name="static-assets")


# ─────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse, include_in_schema=False)
async def root():
    """Serve the React SPA."""
    if _FRONTEND_INDEX.is_file():
        return HTMLResponse(_FRONTEND_INDEX.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>Frontend not built</h1><p>Run <code>npm run build</code> in the <code>frontend/</code> directory.</p>", status_code=503)


@app.get("/health", response_model=HealthResponse, tags=["health"])
async def health_check():
    """
    Check the health and capabilities of the conversion service.
    """
    try:
        import pptx
        pptx_available = True
    except ImportError:
        pptx_available = False
    return HealthResponse(
        status="ok",
        version="1.0.0",
        python_pptx_available=pptx_available,
    )


@app.get("/api/config", include_in_schema=False)
async def get_config():
    """Public config for the browser (Google client ID only)."""
    return JSONResponse({"googleClientId": GOOGLE_CLIENT_ID})


@app.post(
    "/api/convert",
    tags=["conversion"],
    summary="Convert PPTX/PPT to SCORM 2004 package",
    response_class=FileResponse,
    responses={
        200: {
            "description": "SCORM ZIP package ready for LMS upload",
            "content": {"application/zip": {}},
        },
        422: {"description": "Invalid file type or malformed presentation"},
        500: {"description": "Conversion failed"},
    },
)
async def convert_presentation(
    background_tasks: BackgroundTasks,
    file: Annotated[UploadFile, File(description="The .pptx, .ppt, or .pdf file to convert")],
    title: Annotated[Optional[str], Form(description="Override the course title (optional)")] = None,
    mode: Annotated[Optional[str], Form(description="Display mode: 'paginated' or 'continuous' (default: paginated)")] = "paginated",
    theme: Annotated[Optional[str], Form(description="Theme: 'light' or 'dark' (default: light)")] = "light",
    _auth: None = Depends(verify_auth),
):
    """
    ## Convert a PowerPoint presentation to a SCORM 2004 e-learning package.

    ### Input
    - **file**: A `.pptx` or `.ppt` file (multipart/form-data)
    - **title** *(optional)*: Override the course title extracted from the file

    ### Output
    A `.zip` file containing:
    - `imsmanifest.xml` - SCORM 2004 3rd Edition manifest
    - `index.html` - Vue 3 slide player (SCO entry point)
    - `scorm.js` - SCORM 2004 API wrapper
    - `slides/slide-01.png … slide-NN.png` — High-res slide images

    ### Progress Tracking
    | Slides Viewed | `cmi.progress_measure` | `cmi.completion_status` |
    |---|---|---|
    | 0% | 0.0000 | incomplete |
    | 50% | 0.5000 | incomplete |
    | 100% | 1.0000 | **completed** |

    ### LMS Compatibility
    Tested with Moodle, Canvas, TalentLMS, SCORM Cloud, and any
    SCORM 2004-compliant LMS.
    """
    validate_upload(file)

    # Save upload to temp file
    with tempfile.NamedTemporaryFile(
        suffix=Path(file.filename or "upload.pptx").suffix.lower(),
        delete=False
    ) as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name

    try:
        display_mode = mode if mode in ("paginated", "continuous") else "paginated"
        selected_theme = theme if theme in ("light", "dark") else "light"
        zip_path = convert_pptx_to_scorm(
            pptx_path=tmp_path,
            output_dir=OUTPUT_DIR,
            course_title=title or None,
            display_mode=display_mode,
            theme=selected_theme,
        )
    except Exception as e:
        cleanup_file(tmp_path)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    finally:
        cleanup_file(tmp_path)

    zip_filename = Path(zip_path).name

    # Schedule cleanup after response is sent
    background_tasks.add_task(cleanup_file, zip_path)

    return FileResponse(
        path=zip_path,
        filename=zip_filename,
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="{zip_filename}"',
            "X-Slide-Count": "unknown",  # Could parse this from converter
        },
    )


# ─────────────────────────────────────────────
# MCP file staging (upload → file_id, download by file_id)
# ─────────────────────────────────────────────

@app.post("/api/upload", tags=["conversion"], summary="Upload a file for MCP conversion")
async def upload_for_mcp(
    file: Annotated[UploadFile, File(description="The .pptx, .ppt, or .pdf file to stage")],
):
    """Stage a presentation file and return a file_id.

    Use this before calling the MCP convert_presentation tool.
    The file_id is valid for 1 hour.
    """
    validate_upload(file)
    file_id = hashlib.sha256(os.urandom(32)).hexdigest()[:16]
    ext = Path(file.filename or "upload.pptx").suffix.lower()
    dest = os.path.join(UPLOAD_DIR, f"{file_id}{ext}")
    content = await file.read()
    with open(dest, "wb") as f:
        f.write(content)
    staged_files[file_id] = {
        "path": dest,
        "filename": file.filename or "presentation.pptx",
        "created": time.time(),
    }
    return {"file_id": file_id, "filename": file.filename, "size_bytes": len(content)}


@app.get("/api/download/{file_id}", tags=["conversion"], summary="Download a converted file")
async def download_converted(file_id: str):
    """Download a converted SCORM package by file_id."""
    info = staged_files.get(file_id)
    if not info or not os.path.exists(info["path"]):
        raise HTTPException(status_code=404, detail="File not found or expired.")
    return FileResponse(
        path=info["path"],
        filename=info["filename"],
        media_type="application/zip" if info["filename"].endswith(".zip") else "application/octet-stream",
        headers={"Content-Disposition": f'attachment; filename="{info["filename"]}"'},
    )


@app.get(
    "/api/info",
    tags=["health"],
    summary="API capabilities and format support",
)
async def api_info():
    """
    Returns metadata about the API: supported formats, SCORM version,
    and conversion capabilities.
    """
    return {
        "api_version": "1.0.0",
        "scorm_version": "2004 3rd Edition",
        "supported_input_formats": [".pptx", ".ppt", ".pdf", ".doc", ".docx"],
        "output_format": "zip",
        "slide_renderer": "LibreOffice SVG (via pdftocairo)",
        "scorm_player": "Vue 3 (CDN, no build step)",
        "tracked_cmi_elements": [
            "cmi.progress_measure",
            "cmi.completion_status",
            "cmi.success_status",
        ],
        "completion_trigger": "All slides viewed",
        "lms_compatibility": [
            "Moodle",
            "Canvas",
            "TalentLMS",
            "SCORM Cloud",
            "Blackboard",
            "Cornerstone OnDemand",
        ],
        "endpoints": {
            "POST /api/convert": "Upload .pptx/.ppt/.doc/.pdf → download SCORM ZIP",
            "GET /api/info": "This endpoint",
            "GET /health": "Service health check",
            "GET /docs": "Interactive OpenAPI documentation (Swagger UI)",
            "GET /redoc": "Alternative API docs (ReDoc)",
        },
    }


# ─────────────────────────────────────────────
# ASGI dispatcher: MCP paths → _mcp_app, everything else → FastAPI
# ─────────────────────────────────────────────

_fastapi_app = app  # keep a ref to the raw FastAPI app


async def combined_app(scope, receive, send):
    """Route requests to MCP Starlette app or FastAPI based on path."""
    if scope["type"] == "http":
        path = scope.get("path", "")
        if path in _MCP_PATHS:
            await _mcp_app(scope, receive, send)
            return
        # RFC 9728: /.well-known/oauth-protected-resource/<resource-path>
        # Claude requests /.well-known/oauth-protected-resource/mcp
        # Rewrite to the root metadata endpoint the MCP SDK provides
        if path.startswith("/.well-known/oauth-protected-resource/"):
            scope = dict(scope, path="/.well-known/oauth-protected-resource")
            await _mcp_app(scope, receive, send)
            return
    await _fastapi_app(scope, receive, send)


app = combined_app  # uvicorn uses main:app


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
