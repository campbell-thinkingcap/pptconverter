# CLAUDE.md

## Project Overview

PPTX-to-SCORM conversion API. Takes PowerPoint files (.pptx/.ppt) and produces SCORM 2004 3rd Edition ZIP packages with a Vue 3 slide player.

## Tech Stack

- **Python 3.11** with **FastAPI** (web framework) and **Uvicorn** (ASGI server)
- **python-pptx** for reading slide titles and presentation metadata
- **LibreOffice** (headless) for PPTX-to-PDF conversion
- **pdftoppm** (Poppler) for PDF-to-PNG rendering at 150 DPI
- **Vue 3** (CDN, no build step) for the embedded SCORM slide player
- **Docker** for deployment (includes LibreOffice + Poppler)

## Project Structure

```
main.py           - FastAPI app, routes, response models
converter.py      - Core conversion logic (slide extraction, SCORM manifest, HTML player, ZIP packaging)
requirements.txt  - Python dependencies
Dockerfile        - Python 3.11-slim + LibreOffice + Poppler
docker-compose.yml - Single-service compose config
```

## Key Architecture

1. `POST /api/convert` receives a .pptx/.ppt upload
2. `converter.py:convert_pptx_to_scorm()` orchestrates the pipeline:
   - Extract slide titles via python-pptx
   - LibreOffice converts PPTX to PDF
   - pdftoppm renders PDF pages to PNG
   - Builds `imsmanifest.xml` (SCORM 2004 3rd Edition)
   - Builds `index.html` (Vue 3 player with inline SCORM tracking)
   - Bundles everything into a ZIP
3. FastAPI returns the ZIP as a file download, then cleans up via background task

## Running Locally

```bash
# Docker (recommended)
docker-compose up

# Without Docker (requires libreoffice + poppler-utils installed)
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

- Upload UI at `http://localhost:8000`
- Swagger docs at `http://localhost:8000/docs`

## SCORM Details

- Tracks `cmi.progress_measure` (0.0-1.0) based on slides viewed
- Sets `cmi.completion_status` to `completed` and `cmi.success_status` to `passed` when all slides are viewed
- The SCORM JS wrapper (`SCORM_JS` constant in converter.py) finds `API_1484_11` in parent frames

## Web Frontend

- `GET /` serves a drag-and-drop upload page (`UPLOAD_PAGE_HTML` constant in `main.py`)
- Self-contained inline HTML/CSS/JS — no build tools or external assets
- Uses XHR (not fetch) for upload progress tracking
- Dark theme matching the SCORM player palette (#0f172a, #1e293b, #6366f1)
- States: idle → file selected → uploading → converting → complete/error

## Important Notes

- The HTML player and SCORM JS are generated as Python f-strings in `converter.py` (note double-brace `{{` escaping for JS/CSS)
- The upload page HTML is a regular string (not f-string) in `main.py` — no brace escaping needed
- Slide images use zero-padded naming: `slide-01.png`, `slide-02.png`, etc.
- No persistent storage — output ZIPs are written to a temp directory and cleaned up after response
- CORS is wide open (`allow_origins=["*"]`) — restrict for production
