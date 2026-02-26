"""
PPTX → SCORM Conversion API
FastAPI application with OpenAPI 3.0 documentation.
"""
import os
import time
import tempfile
import hashlib
from pathlib import Path
from typing import Annotated, Optional

from fastapi import FastAPI, File, Form, UploadFile, HTTPException, BackgroundTasks, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.security import APIKeyHeader
from pydantic import BaseModel

from google.oauth2 import id_token
from google.auth.transport import requests as google_requests

from converter import convert_pptx_to_scorm

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

    if has_api_key and not has_oauth:
        raise HTTPException(status_code=401, detail="Invalid or missing API key.")
    if has_oauth and not has_api_key:
        raise HTTPException(status_code=401, detail="Google sign-in required.")
    raise HTTPException(status_code=401, detail="Authentication required.")

# ─────────────────────────────────────────────
# App Setup
# ─────────────────────────────────────────────

app = FastAPI(
    title="PPTX → SCORM Conversion API",
    description="""
## Convert PowerPoint presentations to SCORM 2004 e-learning packages.

### Features
- 📊 Supports **.ppt** and **.pptx** formats
- 🖼️ High-fidelity slide rendering via LibreOffice
- 📈 SCORM 2004 3rd Edition compliant (`cmi.progress_measure`)
- 🎯 Each slide becomes a page; completion tracked as % of slides viewed
- 🖥️ Built-in **Vue 3** player — no additional software needed
- 📦 Returns a ready-to-upload **.zip** file for any SCORM LMS

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
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Temp output storage (in production, use S3/blob storage)
OUTPUT_DIR = tempfile.mkdtemp(prefix="scorm_output_")


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

ALLOWED_EXTENSIONS = {".ppt", ".pptx"}

def validate_upload(file: UploadFile) -> None:
    ext = Path(file.filename or "").suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=422,
            detail=f"Unsupported file type '{ext}'. Only .ppt and .pptx are accepted."
        )

def cleanup_file(path: str) -> None:
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


# ─────────────────────────────────────────────
# Upload Page HTML
# ─────────────────────────────────────────────

UPLOAD_PAGE_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>PPTX to SCORM Converter</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: 'Segoe UI', system-ui, sans-serif;
      background: #0f172a;
      color: #e2e8f0;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }

    .container {
      width: 100%;
      max-width: 580px;
    }

    /* ── Header ── */
    .page-header {
      text-align: center;
      margin-bottom: 32px;
    }
    .page-header h1 {
      font-size: 24px;
      font-weight: 700;
      color: #f1f5f9;
      margin-bottom: 8px;
    }
    .page-header p {
      font-size: 14px;
      color: #94a3b8;
    }

    /* ── Upload Zone ── */
    .upload-zone {
      background: #1e293b;
      border: 2px dashed #334155;
      border-radius: 12px;
      padding: 48px 24px;
      text-align: center;
      cursor: pointer;
      transition: border-color 0.2s, background 0.2s;
      position: relative;
    }
    .upload-zone:hover {
      border-color: #475569;
    }
    .upload-zone.drag-over {
      border-color: #6366f1;
      background: rgba(99, 102, 241, 0.08);
    }
    .upload-zone.has-file {
      cursor: default;
      border-style: solid;
      border-color: #334155;
      padding: 32px 24px;
    }

    .upload-icon {
      margin-bottom: 16px;
    }
    .upload-icon svg {
      width: 48px;
      height: 48px;
      color: #64748b;
    }
    .upload-zone.drag-over .upload-icon svg {
      color: #6366f1;
    }

    .upload-text {
      font-size: 15px;
      color: #94a3b8;
      margin-bottom: 4px;
    }
    .upload-text strong {
      color: #e2e8f0;
    }
    .upload-hint {
      font-size: 12px;
      color: #64748b;
      margin-bottom: 16px;
    }

    .browse-btn {
      display: inline-block;
      padding: 8px 20px;
      background: #1e293b;
      color: #94a3b8;
      border: 1px solid #334155;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      transition: background 0.15s, color 0.15s;
    }
    .browse-btn:hover {
      background: #263548;
      color: #e2e8f0;
    }

    /* ── File Selected State ── */
    .file-info {
      display: flex;
      align-items: center;
      gap: 12px;
      background: #0f172a;
      border-radius: 8px;
      padding: 12px 16px;
      margin-bottom: 16px;
      text-align: left;
    }
    .file-icon svg {
      width: 32px;
      height: 32px;
      color: #6366f1;
      flex-shrink: 0;
    }
    .file-details {
      flex: 1;
      min-width: 0;
    }
    .file-name {
      font-size: 14px;
      font-weight: 600;
      color: #f1f5f9;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .file-size {
      font-size: 12px;
      color: #64748b;
    }
    .remove-btn {
      background: none;
      border: none;
      color: #64748b;
      cursor: pointer;
      padding: 4px;
      border-radius: 4px;
      transition: color 0.15s;
      flex-shrink: 0;
    }
    .remove-btn:hover {
      color: #ef4444;
    }

    .title-input {
      width: 100%;
      padding: 10px 14px;
      background: #0f172a;
      border: 1px solid #334155;
      border-radius: 8px;
      color: #e2e8f0;
      font-size: 13px;
      font-family: inherit;
      margin-bottom: 16px;
      outline: none;
      transition: border-color 0.15s;
    }
    .title-input::placeholder { color: #475569; }
    .title-input:focus { border-color: #6366f1; }

    .convert-btn {
      width: 100%;
      padding: 12px;
      background: #6366f1;
      color: #fff;
      border: none;
      border-radius: 8px;
      font-size: 15px;
      font-weight: 700;
      cursor: pointer;
      transition: background 0.15s;
    }
    .convert-btn:hover { background: #4f46e5; }

    /* ── Progress ── */
    .progress-section {
      background: #1e293b;
      border: 1px solid #334155;
      border-radius: 12px;
      padding: 32px 24px;
      text-align: center;
    }
    .progress-status {
      font-size: 15px;
      font-weight: 600;
      color: #f1f5f9;
      margin-bottom: 16px;
    }
    .progress-bar {
      width: 100%;
      height: 8px;
      background: #334155;
      border-radius: 4px;
      overflow: hidden;
      margin-bottom: 12px;
    }
    .progress-fill {
      height: 100%;
      background: linear-gradient(90deg, #6366f1, #8b5cf6);
      border-radius: 4px;
      transition: width 0.3s ease;
      width: 0%;
    }
    .progress-fill.indeterminate {
      width: 100%;
      animation: pulse 1.8s ease-in-out infinite;
    }
    @keyframes pulse {
      0%, 100% { opacity: 0.35; }
      50% { opacity: 1; }
    }
    .progress-hint {
      font-size: 12px;
      color: #64748b;
    }

    /* ── Result ── */
    .result-section {
      background: #1e293b;
      border: 1px solid #334155;
      border-radius: 12px;
      padding: 32px 24px;
      text-align: center;
    }
    .result-icon {
      width: 56px;
      height: 56px;
      background: rgba(34, 197, 94, 0.15);
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0 auto 16px;
    }
    .result-icon svg {
      width: 28px;
      height: 28px;
      color: #22c55e;
    }
    .result-title {
      font-size: 18px;
      font-weight: 700;
      color: #22c55e;
      margin-bottom: 6px;
    }
    .result-filename {
      font-size: 13px;
      color: #94a3b8;
      margin-bottom: 20px;
      word-break: break-all;
    }
    .result-actions {
      display: flex;
      gap: 10px;
      justify-content: center;
    }
    .btn-download {
      padding: 10px 24px;
      background: #22c55e;
      color: #fff;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      transition: background 0.15s;
    }
    .btn-download:hover { background: #16a34a; }
    .btn-another {
      padding: 10px 24px;
      background: #1e293b;
      color: #94a3b8;
      border: 1px solid #334155;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: background 0.15s, color 0.15s;
    }
    .btn-another:hover { background: #263548; color: #e2e8f0; }

    /* ── Error ── */
    .error-section {
      background: #1e293b;
      border: 1px solid #7f1d1d;
      border-radius: 12px;
      padding: 32px 24px;
      text-align: center;
    }
    .error-icon {
      width: 56px;
      height: 56px;
      background: rgba(239, 68, 68, 0.15);
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0 auto 16px;
    }
    .error-icon svg {
      width: 28px;
      height: 28px;
      color: #ef4444;
    }
    .error-title {
      font-size: 18px;
      font-weight: 700;
      color: #ef4444;
      margin-bottom: 6px;
    }
    .error-detail {
      font-size: 13px;
      color: #94a3b8;
      margin-bottom: 20px;
    }
    .btn-retry {
      padding: 10px 24px;
      background: #6366f1;
      color: #fff;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      transition: background 0.15s;
    }
    .btn-retry:hover { background: #4f46e5; }

    /* ── Validation error (inline) ── */
    .validation-error {
      color: #ef4444;
      font-size: 13px;
      margin-top: 12px;
    }

    /* ── Footer ── */
    .page-footer {
      text-align: center;
      margin-top: 24px;
      font-size: 12px;
      color: #475569;
    }
    .page-footer a {
      color: #6366f1;
      text-decoration: none;
    }
    .page-footer a:hover { text-decoration: underline; }

    .hidden { display: none !important; }

    /* ── Auth overlay ── */
    .auth-overlay {
      position: fixed;
      inset: 0;
      background: #0f172a;
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 100;
      flex-direction: column;
      gap: 24px;
    }
    .auth-card {
      background: #1e293b;
      border: 1px solid #334155;
      border-radius: 16px;
      padding: 48px 40px;
      text-align: center;
      max-width: 400px;
      width: 100%;
    }
    .auth-card h2 {
      font-size: 20px;
      font-weight: 700;
      color: #f1f5f9;
      margin-bottom: 8px;
    }
    .auth-card p {
      font-size: 14px;
      color: #64748b;
      margin-bottom: 32px;
    }
    #g_id_signin {
      display: flex;
      justify-content: center;
    }
    .auth-error {
      color: #ef4444;
      font-size: 13px;
      margin-top: 16px;
    }

    /* ── User bar ── */
    .user-bar {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 12px;
      margin-bottom: 16px;
      font-size: 13px;
      color: #94a3b8;
    }
    .user-bar img {
      width: 28px;
      height: 28px;
      border-radius: 50%;
      border: 1px solid #334155;
    }
    .sign-out-btn {
      background: none;
      border: 1px solid #334155;
      border-radius: 6px;
      color: #64748b;
      padding: 3px 10px;
      font-size: 12px;
      cursor: pointer;
      transition: color 0.15s, border-color 0.15s;
    }
    .sign-out-btn:hover { color: #ef4444; border-color: #ef4444; }

    @media (max-width: 480px) {
      .page-header h1 { font-size: 20px; }
      .upload-zone { padding: 32px 16px; }
      .result-actions { flex-direction: column; }
    }
  </style>
</head>
<body>
<!-- Auth overlay — shown until signed in (only when Google OAuth is configured) -->
<div id="auth-overlay" class="auth-overlay hidden">
  <div class="auth-card">
    <h2>Sign in to continue</h2>
    <p>Use your organisation Google account to access the converter.</p>
    <div id="g_id_signin"></div>
    <div id="auth-error" class="auth-error hidden"></div>
  </div>
</div>

<div class="container">
  <!-- User bar — shown when signed in via Google -->
  <div id="user-bar" class="user-bar hidden">
    <img id="user-avatar" src="" alt=""/>
    <span id="user-email"></span>
    <button class="sign-out-btn" id="sign-out-btn">Sign out</button>
  </div>

  <div class="page-header">
    <h1>PPTX to SCORM Converter</h1>
    <p>Convert PowerPoint presentations into SCORM 2004 e-learning packages</p>
  </div>

  <!-- Upload Zone -->
  <div id="upload-zone" class="upload-zone">
    <div id="idle-content">
      <div class="upload-icon">
        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor">
          <path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5m-13.5-9L12 3m0 0 4.5 4.5M12 3v13.5"/>
        </svg>
      </div>
      <div class="upload-text"><strong>Drag &amp; drop</strong> your presentation here</div>
      <div class="upload-hint">.pptx or .ppt files only</div>
      <button type="button" class="browse-btn" id="browse-btn">Browse Files</button>
      <input type="file" id="file-input" accept=".ppt,.pptx" hidden/>
      <div id="validation-error" class="validation-error hidden"></div>
    </div>

    <div id="file-content" class="hidden">
      <div class="file-info">
        <div class="file-icon">
          <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor">
            <path stroke-linecap="round" stroke-linejoin="round" d="M3.75 3v11.25A2.25 2.25 0 0 0 6 16.5h2.25M3.75 3h-1.5m1.5 0h16.5m0 0h1.5m-1.5 0v11.25A2.25 2.25 0 0 1 18 16.5h-2.25m-7.5 0h7.5m-7.5 0-1 3m8.5-3 1 3m0 0 .5 1.5m-.5-1.5h-9.5m0 0-.5 1.5M9 11.25v1.5M12 9v3.75m3-6v6"/>
          </svg>
        </div>
        <div class="file-details">
          <div class="file-name" id="file-name"></div>
          <div class="file-size" id="file-size"></div>
        </div>
        <button class="remove-btn" id="remove-btn" title="Remove file">
          <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor">
            <path stroke-linecap="round" stroke-linejoin="round" d="M6 18 18 6M6 6l12 12"/>
          </svg>
        </button>
      </div>
      <input type="text" class="title-input" id="title-input" placeholder="Course title (optional — uses file metadata if blank)"/>
      <input type="password" class="title-input" id="api-key-input" placeholder="API key (required if server has API_KEY set)"/>
      <button class="convert-btn" id="convert-btn">Convert to SCORM</button>
    </div>
  </div>

  <!-- Progress -->
  <div id="progress-section" class="progress-section hidden">
    <div class="progress-status" id="progress-status">Uploading...</div>
    <div class="progress-bar">
      <div class="progress-fill" id="progress-fill"></div>
    </div>
    <div class="progress-hint" id="progress-hint">Sending your file to the server</div>
  </div>

  <!-- Result -->
  <div id="result-section" class="result-section hidden">
    <div class="result-icon">
      <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor">
        <path stroke-linecap="round" stroke-linejoin="round" d="m4.5 12.75 6 6 9-13.5"/>
      </svg>
    </div>
    <div class="result-title">SCORM Package Ready!</div>
    <div class="result-filename" id="result-filename"></div>
    <div class="result-actions">
      <button class="btn-download" id="download-btn">Download Again</button>
      <button class="btn-another" id="another-btn">Convert Another</button>
    </div>
  </div>

  <!-- Error -->
  <div id="error-section" class="error-section hidden">
    <div class="error-icon">
      <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor">
        <path stroke-linecap="round" stroke-linejoin="round" d="M12 9v3.75m9-.75a9 9 0 1 1-18 0 9 9 0 0 1 18 0Zm-9 3.75h.008v.008H12v-.008Z"/>
      </svg>
    </div>
    <div class="error-title">Conversion Failed</div>
    <div class="error-detail" id="error-detail"></div>
    <button class="btn-retry" id="retry-btn">Try Again</button>
  </div>

  <div class="page-footer">
    <a href="/docs">API Documentation</a> &middot; Powered by FastAPI + LibreOffice
  </div>
</div>

<script src="https://accounts.google.com/gsi/client" async defer></script>
<script>
(function() {
  // ── Elements ──
  const zone = document.getElementById('upload-zone');
  const idleContent = document.getElementById('idle-content');
  const fileContent = document.getElementById('file-content');
  const fileInput = document.getElementById('file-input');
  const fileNameEl = document.getElementById('file-name');
  const fileSizeEl = document.getElementById('file-size');
  const removeBtn = document.getElementById('remove-btn');
  const titleInput = document.getElementById('title-input');
  const apiKeyInput = document.getElementById('api-key-input');
  const convertBtn = document.getElementById('convert-btn');
  const progressSection = document.getElementById('progress-section');
  const progressStatus = document.getElementById('progress-status');
  const progressFill = document.getElementById('progress-fill');
  const progressHint = document.getElementById('progress-hint');
  const resultSection = document.getElementById('result-section');
  const resultFilename = document.getElementById('result-filename');
  const downloadBtn = document.getElementById('download-btn');
  const anotherBtn = document.getElementById('another-btn');
  const errorSection = document.getElementById('error-section');
  const errorDetail = document.getElementById('error-detail');
  const retryBtn = document.getElementById('retry-btn');
  const validationError = document.getElementById('validation-error');

  // ── Auth elements ──
  const authOverlay = document.getElementById('auth-overlay');
  const userBar = document.getElementById('user-bar');
  const userAvatar = document.getElementById('user-avatar');
  const userEmailEl = document.getElementById('user-email');
  const signOutBtn = document.getElementById('sign-out-btn');
  const authError = document.getElementById('auth-error');

  var googleClientId = '';
  var googleIdToken = '';
  var authMode = 'none';  // 'none' | 'google' | 'apikey'

  let selectedFile = null;
  let lastBlobUrl = null;
  let lastFilename = null;

  const ALLOWED = ['.ppt', '.pptx'];

  // ── Helpers ──
  function show(el) { el.classList.remove('hidden'); }
  function hide(el) { el.classList.add('hidden'); }

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  function getExtension(name) {
    const i = name.lastIndexOf('.');
    return i >= 0 ? name.substring(i).toLowerCase() : '';
  }

  function parseFilename(header) {
    if (!header) return 'scorm_package.zip';
    const match = header.match(/filename="?([^";]+)"?/);
    return match ? match[1] : 'scorm_package.zip';
  }

  // ── State transitions ──
  function resetToIdle() {
    selectedFile = null;
    fileInput.value = '';
    titleInput.value = '';
    if (lastBlobUrl) { URL.revokeObjectURL(lastBlobUrl); lastBlobUrl = null; }

    show(zone);
    show(idleContent);
    hide(fileContent);
    hide(progressSection);
    hide(resultSection);
    hide(errorSection);
    hide(validationError);

    zone.classList.remove('has-file');
    progressFill.style.width = '0%';
    progressFill.classList.remove('indeterminate');
  }

  function showFileSelected(file) {
    selectedFile = file;
    fileNameEl.textContent = file.name;
    fileSizeEl.textContent = formatSize(file.size);

    hide(idleContent);
    show(fileContent);
    zone.classList.add('has-file');
    hide(validationError);
  }

  function showValidationError(msg) {
    validationError.textContent = msg;
    show(validationError);
  }

  function showProgress() {
    hide(zone);
    hide(resultSection);
    hide(errorSection);
    show(progressSection);
    progressFill.style.width = '0%';
    progressFill.classList.remove('indeterminate');
    progressStatus.textContent = 'Uploading...';
    progressHint.textContent = 'Sending your file to the server';
  }

  function updateProgress(pct) {
    progressFill.style.width = pct + '%';
    progressHint.textContent = pct + '% uploaded';
  }

  function showConverting() {
    progressStatus.textContent = 'Converting...';
    progressHint.textContent = 'Rendering slides — this may take a moment';
    progressFill.style.width = '';
    progressFill.classList.add('indeterminate');
  }

  function showResult(filename) {
    hide(progressSection);
    show(resultSection);
    resultFilename.textContent = filename;
    lastFilename = filename;
  }

  function showError(msg) {
    hide(progressSection);
    hide(zone);
    show(errorSection);
    errorDetail.textContent = msg;
  }

  // ── File selection ──
  function handleFile(file) {
    if (!file) return;
    const ext = getExtension(file.name);
    if (!ALLOWED.includes(ext)) {
      showValidationError('Only .ppt and .pptx files are supported.');
      return;
    }
    showFileSelected(file);
  }

  // ── Drag & drop ──
  ['dragenter', 'dragover'].forEach(function(evt) {
    zone.addEventListener(evt, function(e) {
      e.preventDefault();
      e.stopPropagation();
      if (!zone.classList.contains('has-file')) {
        zone.classList.add('drag-over');
      }
    });
  });

  ['dragleave', 'drop'].forEach(function(evt) {
    zone.addEventListener(evt, function(e) {
      e.preventDefault();
      e.stopPropagation();
      zone.classList.remove('drag-over');
    });
  });

  zone.addEventListener('drop', function(e) {
    if (zone.classList.contains('has-file')) return;
    var file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });

  zone.addEventListener('click', function(e) {
    if (zone.classList.contains('has-file')) return;
    if (e.target.closest('#browse-btn') || e.target === zone || e.target.closest('.upload-icon') || e.target.closest('.upload-text') || e.target.closest('.upload-hint')) {
      fileInput.click();
    }
  });

  fileInput.addEventListener('change', function() {
    if (fileInput.files[0]) handleFile(fileInput.files[0]);
  });

  removeBtn.addEventListener('click', function(e) {
    e.stopPropagation();
    resetToIdle();
  });

  // ── Conversion ──
  convertBtn.addEventListener('click', function() {
    if (!selectedFile) return;
    uploadFile(selectedFile, titleInput.value.trim(), apiKeyInput.value.trim());
  });

  function uploadFile(file, title, apiKey) {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/api/convert');
    if (authMode === 'google' && googleIdToken) {
      xhr.setRequestHeader('X-Google-Token', googleIdToken);
    } else if (apiKey) {
      xhr.setRequestHeader('X-API-Key', apiKey);
    }
    xhr.responseType = 'blob';

    xhr.upload.addEventListener('progress', function(e) {
      if (e.lengthComputable) {
        var pct = Math.round((e.loaded / e.total) * 100);
        updateProgress(pct);
        if (pct >= 100) showConverting();
      }
    });

    xhr.addEventListener('load', function() {
      if (xhr.status === 200) {
        var filename = parseFilename(xhr.getResponseHeader('Content-Disposition'));
        lastBlobUrl = URL.createObjectURL(xhr.response);
        triggerDownload(lastBlobUrl, filename);
        showResult(filename);
      } else if ((xhr.status === 401 || xhr.status === 403) && authMode === 'google') {
        // Token expired or domain rejected — re-prompt sign-in
        googleIdToken = '';
        hide(userBar);
        authError.textContent = xhr.status === 403
          ? 'Your email domain is not allowed.'
          : 'Your session expired. Please sign in again.';
        show(authError);
        show(authOverlay);
        renderGoogleButton();
        resetToIdle();
      } else {
        // Parse error from JSON blob
        var reader = new FileReader();
        reader.onload = function() {
          try {
            var json = JSON.parse(reader.result);
            showError(json.detail || 'Conversion failed.');
          } catch(e) {
            showError('Conversion failed (HTTP ' + xhr.status + ').');
          }
        };
        reader.onerror = function() {
          showError('Conversion failed (HTTP ' + xhr.status + ').');
        };
        reader.readAsText(xhr.response);
      }
    });

    xhr.addEventListener('error', function() {
      showError('Network error. Is the server running?');
    });

    var form = new FormData();
    form.append('file', file);
    if (title) form.append('title', title);

    showProgress();
    xhr.send(form);
  }

  function triggerDownload(url, filename) {
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  // ── Result actions ──
  downloadBtn.addEventListener('click', function() {
    if (lastBlobUrl && lastFilename) triggerDownload(lastBlobUrl, lastFilename);
  });

  anotherBtn.addEventListener('click', resetToIdle);
  retryBtn.addEventListener('click', resetToIdle);

  // ── Google OAuth ──
  function renderGoogleButton() {
    google.accounts.id.renderButton(
      document.getElementById('g_id_signin'),
      { theme: 'filled_black', size: 'large', shape: 'rectangular', width: 280 }
    );
  }

  function handleCredential(response) {
    googleIdToken = response.credential;
    // Decode JWT payload for display (server does real verification)
    try {
      var parts = googleIdToken.split('.');
      var payload = JSON.parse(atob(parts[1].replace(/-/g, '+').replace(/_/g, '/')));
      userEmailEl.textContent = payload.email || '';
      if (payload.picture) { userAvatar.src = payload.picture; show(userAvatar); }
    } catch(e) {}
    show(userBar);
    hide(authOverlay);
    hide(authError);
  }

  function waitForGIS(fn) {
    if (typeof google !== 'undefined' && google.accounts && google.accounts.id) {
      fn();
    } else {
      setTimeout(function() { waitForGIS(fn); }, 100);
    }
  }

  function initAuth() {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', '/api/config');
    xhr.onload = function() {
      if (xhr.status === 200) {
        try {
          var cfg = JSON.parse(xhr.responseText);
          googleClientId = cfg.googleClientId || '';
        } catch(e) {}
      }
      if (googleClientId) {
        authMode = 'google';
        // Hide API key input — not needed with Google auth
        apiKeyInput.style.display = 'none';
        show(authOverlay);
        waitForGIS(function() {
          google.accounts.id.initialize({
            client_id: googleClientId,
            callback: handleCredential,
            auto_select: false,
            cancel_on_tap_outside: false
          });
          renderGoogleButton();
        });
      } else {
        authMode = 'apikey';
        hide(authOverlay);
      }
    };
    xhr.onerror = function() {
      authMode = 'apikey';
      hide(authOverlay);
    };
    xhr.send();
  }

  signOutBtn.addEventListener('click', function() {
    googleIdToken = '';
    userEmailEl.textContent = '';
    userAvatar.src = '';
    hide(userBar);
    if (authMode === 'google') {
      google.accounts.id.disableAutoSelect();
      show(authOverlay);
      hide(authError);
      renderGoogleButton();
    }
    resetToIdle();
  });

  // ── Init ──
  resetToIdle();
  initAuth();
})();
</script>
</body>
</html>"""


# ─────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse, include_in_schema=False)
async def root():
    """Upload page — drag and drop PPTX files for conversion."""
    return HTMLResponse(UPLOAD_PAGE_HTML)


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
    file: Annotated[UploadFile, File(description="The .pptx or .ppt file to convert")],
    title: Annotated[Optional[str], Form(description="Override the course title (optional)")] = None,
    _auth: None = Depends(verify_auth),
):
    """
    ## Convert a PowerPoint presentation to a SCORM 2004 e-learning package.

    ### Input
    - **file**: A `.pptx` or `.ppt` file (multipart/form-data)
    - **title** *(optional)*: Override the course title extracted from the file

    ### Output
    A `.zip` file containing:
    - `imsmanifest.xml` — SCORM 2004 3rd Edition manifest
    - `index.html` — Vue 3 slide player (SCO entry point)
    - `scorm.js` — SCORM 2004 API wrapper
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
        zip_path = convert_pptx_to_scorm(
            pptx_path=tmp_path,
            output_dir=OUTPUT_DIR,
            course_title=title or None,
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
        "supported_input_formats": [".pptx", ".ppt"],
        "output_format": "zip",
        "slide_renderer": "python-pptx HTML extraction",
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
            "POST /api/convert": "Upload .pptx/.ppt → download SCORM ZIP",
            "GET /api/info": "This endpoint",
            "GET /health": "Service health check",
            "GET /docs": "Interactive OpenAPI documentation (Swagger UI)",
            "GET /redoc": "Alternative API docs (ReDoc)",
        },
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)
