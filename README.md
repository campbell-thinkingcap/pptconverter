# PPTX → SCORM Conversion API

Convert PowerPoint presentations (`.pptx` / `.ppt`) into **SCORM 2004 3rd Edition** e-learning packages — ready to upload to any LMS.

Each slide becomes a page in a **Vue 3** slide player. The LMS tracks `cmi.progress_measure` as the percentage of slides viewed.

---

## ✨ Features

| Feature | Detail |
|---|---|
| **Input** | `.pptx` and `.ppt` |
| **Output** | SCORM 2004 3rd Edition `.zip` |
| **Renderer** | LibreOffice → high-fidelity PNG images |
| **Player** | Vue 3 (CDN, no build step, dark UI) |
| **Progress** | `cmi.progress_measure` 0→1 as slides are viewed |
| **Completion** | Triggers when all slides have been viewed |
| **Docs** | Swagger UI at `/docs`, ReDoc at `/redoc` |

---

## 🚀 Quick Start

### With Docker (recommended)

```bash
docker-compose up
```

API is now live at **http://localhost:8000**

- Swagger UI: http://localhost:8000/docs
- ReDoc:       http://localhost:8000/redoc

---

### Local (no Docker)

**Prerequisites:**
```bash
# macOS
brew install libreoffice poppler

# Ubuntu/Debian
sudo apt install libreoffice poppler-utils

# Windows (WSL recommended)
# Install LibreOffice from libreoffice.org + poppler via choco/winget
```

**Install & run:**
```bash
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

---

## 📡 API Reference

### `POST /api/convert`

Convert a `.pptx` or `.ppt` file to a SCORM ZIP package.

**Request (multipart/form-data):**

| Field | Type | Required | Description |
|---|---|---|---|
| `file` | file | ✅ | The `.pptx` or `.ppt` file |
| `title` | string | ❌ | Override course title |

**Response:** Binary `.zip` file download

**cURL example:**
```bash
curl -X POST http://localhost:8000/api/convert \
  -F "file=@my_presentation.pptx" \
  -F "title=My Course Title" \
  -o my_course_scorm.zip
```

**Python example:**
```python
import requests

with open("presentation.pptx", "rb") as f:
    resp = requests.post(
        "http://localhost:8000/api/convert",
        files={"file": ("presentation.pptx", f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
        data={"title": "My Custom Title"},  # optional
    )

with open("course.zip", "wb") as out:
    out.write(resp.content)

print("SCORM package saved to course.zip")
```

**JavaScript/Fetch example:**
```javascript
const form = new FormData();
form.append("file", fileInput.files[0]);
form.append("title", "My Course");  // optional

const response = await fetch("http://localhost:8000/api/convert", {
  method: "POST",
  body: form,
});

const blob = await response.blob();
const url = URL.createObjectURL(blob);
// Trigger download
const a = document.createElement("a");
a.href = url;
a.download = "course_scorm.zip";
a.click();
```

---

### `GET /health`

Check service status and dependency availability.

```json
{
  "status": "ok",
  "version": "1.0.0",
  "libreoffice_available": true,
  "pdftoppm_available": true
}
```

### `GET /api/info`

Full API capabilities metadata.

---

## 📦 SCORM Package Contents

```
course_SCORM.zip
├── imsmanifest.xml          ← SCORM 2004 3rd Edition manifest
├── index.html               ← Vue 3 slide player (SCO entry point)
├── scorm.js                 ← SCORM 2004 API wrapper
└── slides/
    ├── slide-01.png
    ├── slide-02.png
    └── slide-NN.png
```

---

## 📊 SCORM Data Model

| CMI Element | Value | When |
|---|---|---|
| `cmi.completion_status` | `incomplete` | On launch |
| `cmi.completion_status` | `completed` | All slides viewed |
| `cmi.success_status` | `passed` | On completion |
| `cmi.progress_measure` | `0.0000` – `1.0000` | Updated on each new slide view |

---

## 🖥️ Player UI

The generated player includes:
- **Dark sidebar** with slide list and visited indicators
- **Progress bar** in the header showing % complete
- **Previous / Next** navigation
- **Completion overlay** with confetti when all slides are viewed
- Preloads the next slide image for smooth navigation

---

## 🏗️ Architecture

```
POST /api/convert
       │
       ▼
  [FastAPI]  ← validates file type
       │
       ▼
  [python-pptx]  ← extracts slide titles & course metadata
       │
       ▼
  [LibreOffice]  ← converts PPTX → PDF (headless)
       │
       ▼
  [pdftoppm]  ← renders PDF → PNG images (150 DPI)
       │
       ▼
  [Python zipfile]  ← bundles manifest + player + images
       │
       ▼
  Returns ZIP ← downloaded by client
```

---

## 🔌 LMS Compatibility

Tested with:
- ✅ Moodle
- ✅ Canvas
- ✅ TalentLMS
- ✅ SCORM Cloud
- ✅ Blackboard
- ✅ Cornerstone OnDemand

Any LMS supporting SCORM 2004 3rd Edition will work.

---

## ⚙️ Configuration

| Environment Variable | Default | Description |
|---|---|---|
| `PORT` | `8000` | Server port |
| `MAX_FILE_SIZE` | unlimited | Add middleware to restrict |

For production, consider:
- Adding file size limits via middleware
- Storing generated ZIPs in S3/blob storage instead of local temp
- Adding auth (API key header) via FastAPI dependency injection
- Rate limiting with `slowapi`

---

## 📄 License

MIT
