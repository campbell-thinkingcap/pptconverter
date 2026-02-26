"""
PPTX/PPT → SCORM 2004 3rd Edition Converter
Extracts slide content (text, images, tables) as HTML, wraps in a Vue 3 SCORM player package.
"""
import html
import os
import re
import json
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.util import Inches, Emu, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE


# ─────────────────────────────────────────────
# HTML Slide Extraction
# ─────────────────────────────────────────────

def _run_to_html(run) -> str:
    """Convert a single text Run to an HTML span with inline styles."""
    text = html.escape(run.text)
    if not text:
        return ""

    styles = []
    font = run.font

    if font.size is not None:
        styles.append(f"font-size:{font.size.pt:.1f}pt")

    try:
        if font.color and font.color.rgb is not None:
            rgb = font.color.rgb
            r, g, b = int(str(rgb)[0:2], 16), int(str(rgb)[2:4], 16), int(str(rgb)[4:6], 16)
            luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
            if luminance > 0.85:
                styles.append("color:#111111")
            else:
                styles.append(f"color:#{rgb}")
    except (AttributeError, ValueError):
        pass  # theme color without explicit RGB

    if font.bold:
        styles.append("font-weight:bold")
    if font.italic:
        styles.append("font-style:italic")
    if font.underline:
        styles.append("text-decoration:underline")
    if font.name:
        styles.append(f"font-family:'{html.escape(font.name)}',sans-serif")

    if styles:
        return f'<span style="{";".join(styles)}">{text}</span>'
    return text


def _text_frame_to_html(text_frame, is_title: bool = False) -> str:
    """Convert a TextFrame's paragraphs and runs to formatted HTML."""
    parts = []
    first_nonempty = True

    for paragraph in text_frame.paragraphs:
        # Determine alignment
        align_map = {
            PP_ALIGN.LEFT: "left",
            PP_ALIGN.CENTER: "center",
            PP_ALIGN.RIGHT: "right",
            PP_ALIGN.JUSTIFY: "justify",
        }
        text_align = align_map.get(paragraph.alignment, "inherit")
        p_style = f"text-align:{text_align};margin:0;padding:2px 0;line-height:1.4"

        # Build runs HTML
        runs_html = "".join(_run_to_html(run) for run in paragraph.runs)

        if not runs_html.strip():
            parts.append(f'<p style="{p_style}">&nbsp;</p>')
            continue

        if is_title and first_nonempty:
            parts.append(f'<h1 style="{p_style}">{runs_html}</h1>')
            first_nonempty = False
        elif is_title:
            parts.append(f'<h2 style="{p_style}">{runs_html}</h2>')
        else:
            parts.append(f'<p style="{p_style}">{runs_html}</p>')

    return "\n".join(parts)


def _get_slide_background(slide) -> str:
    """Extract slide background color, defaulting to white."""
    try:
        fill = slide.background.fill
        if fill.type is not None:
            rgb = fill.fore_color.rgb
            return f"#{rgb}"
    except Exception:
        pass
    return "#ffffff"


def _picture_to_html(shape, pos_style: str, images_dir: str, image_counter: list) -> str:
    """Extract image from a Picture shape, save to disk, return an <img> tag."""
    image_counter[0] += 1
    try:
        blob = shape.image.blob
        ext = shape.image.ext
    except Exception:
        return ""

    filename = f"img-{str(image_counter[0]).zfill(3)}.{ext}"
    filepath = os.path.join(images_dir, filename)
    with open(filepath, "wb") as f:
        f.write(blob)

    rel_path = f"slides/images/{filename}"
    return (
        f'<div style="{pos_style}">'
        f'<img src="{rel_path}" alt="" style="width:100%;height:100%;object-fit:contain;" />'
        f'</div>'
    )


def _table_to_html(shape, pos_style: str) -> str:
    """Convert a Table shape to an HTML <table> element."""
    table = shape.table
    rows_html = []

    # Track cells claimed by merges
    claimed = set()
    for row_idx in range(len(table.rows)):
        for col_idx in range(len(table.columns)):
            cell = table.cell(row_idx, col_idx)
            if cell.is_merge_origin:
                for dr in range(cell.span_height):
                    for dc in range(cell.span_width):
                        if dr == 0 and dc == 0:
                            continue
                        claimed.add((row_idx + dr, col_idx + dc))

    for row_idx in range(len(table.rows)):
        cells_html = []
        for col_idx in range(len(table.columns)):
            if (row_idx, col_idx) in claimed:
                continue
            cell = table.cell(row_idx, col_idx)
            cell_content = _text_frame_to_html(cell.text_frame)
            span_attrs = ""
            if cell.is_merge_origin:
                if cell.span_height > 1:
                    span_attrs += f' rowspan="{cell.span_height}"'
                if cell.span_width > 1:
                    span_attrs += f' colspan="{cell.span_width}"'
            cells_html.append(
                f'<td{span_attrs} style="border:1px solid #ccc;padding:4px 8px;vertical-align:top;">'
                f'{cell_content}</td>'
            )
        rows_html.append(f'<tr>{"".join(cells_html)}</tr>')

    return (
        f'<div style="{pos_style}overflow:auto;">'
        f'<table style="width:100%;height:100%;border-collapse:collapse;font-size:11pt;">'
        f'{"".join(rows_html)}'
        f'</table></div>'
    )


def _text_shape_to_html(shape, pos_style: str) -> str:
    """Convert a text-containing shape to a positioned HTML div."""
    is_title = False
    if shape.is_placeholder:
        try:
            ph_idx = shape.placeholder_format.idx
            is_title = ph_idx in (0, 1, 13)
        except Exception:
            pass

    # Vertical alignment
    valign_css = ""
    try:
        anchor = shape.text_frame.vertical_anchor
        if anchor is not None:
            anchor_val = int(anchor)
            if anchor_val == 1:  # MIDDLE
                valign_css = "display:flex;flex-direction:column;justify-content:center;"
            elif anchor_val == 2:  # BOTTOM
                valign_css = "display:flex;flex-direction:column;justify-content:flex-end;"
    except Exception:
        pass

    # Shape background
    bg_css = ""
    try:
        fill = shape.fill
        if fill.type is not None and str(fill.type) != "BACKGROUND (6)":
            rgb = fill.fore_color.rgb
            bg_css = f"background-color:#{rgb};"
    except Exception:
        pass

    inner = _text_frame_to_html(shape.text_frame, is_title)
    if not inner.strip():
        return ""

    return f'<div style="{pos_style}{valign_css}{bg_css}overflow:hidden;">{inner}</div>'


def _shape_to_html(
    shape,
    slide_width: int,
    slide_height: int,
    images_dir: str,
    image_counter: list,
) -> str:
    """Convert a shape to an absolutely-positioned HTML element."""
    # Calculate position as percentages
    left_pct = (shape.left / slide_width) * 100
    top_pct = (shape.top / slide_height) * 100
    width_pct = (shape.width / slide_width) * 100
    height_pct = (shape.height / slide_height) * 100

    pos_style = (
        f"position:absolute;"
        f"left:{left_pct:.4f}%;"
        f"top:{top_pct:.4f}%;"
        f"width:{width_pct:.4f}%;"
        f"height:{height_pct:.4f}%;"
    )

    # Add rotation if present
    if shape.rotation:
        pos_style += f"transform:rotate({shape.rotation}deg);transform-origin:center center;"

    shape_type = shape.shape_type

    # Picture shapes
    if shape_type == MSO_SHAPE_TYPE.PICTURE:
        return _picture_to_html(shape, pos_style, images_dir, image_counter)

    # Table shapes
    if shape.has_table:
        return _table_to_html(shape, pos_style)

    # Group shapes
    if shape_type == MSO_SHAPE_TYPE.GROUP:
        return _group_to_html(shape, pos_style, images_dir, image_counter)

    # Text shapes (auto shapes, text boxes, placeholders)
    if shape.has_text_frame:
        return _text_shape_to_html(shape, pos_style)

    return ""  # Skip unsupported shapes (charts, SmartArt, freeform, etc.)


def _group_to_html(
    shape, pos_style: str, images_dir: str, image_counter: list,
) -> str:
    """Recursively process a group shape."""
    children_html = []
    for child in shape.shapes:
        child_html = _shape_to_html(
            child, shape.width, shape.height, images_dir, image_counter
        )
        if child_html:
            children_html.append(child_html)

    if not children_html:
        return ""

    return f'<div style="{pos_style}">{"".join(children_html)}</div>'


def _extract_slide_html(pptx_path: str, work_dir: str) -> list[str]:
    """
    Extract each slide from the PPTX as an HTML string.
    Images are saved to work_dir/slides/images/ and referenced via relative paths.
    Returns list of HTML strings, one per slide.
    """
    prs = Presentation(pptx_path)
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    aspect_pct = (slide_height / slide_width) * 100

    images_dir = os.path.join(work_dir, "slides", "images")
    os.makedirs(images_dir, exist_ok=True)

    image_counter = [0]
    slide_html_list = []

    for slide in prs.slides:
        bg_color = _get_slide_background(slide)

        shapes_html = []
        for shape in slide.shapes:
            shape_html = _shape_to_html(
                shape, slide_width, slide_height, images_dir, image_counter
            )
            if shape_html:
                shapes_html.append(shape_html)

        slide_div = (
            f'<div class="slide-content" style="position:relative;width:100%;'
            f'padding-bottom:{aspect_pct:.4f}%;background:{bg_color};">'
            f'{"".join(shapes_html)}'
            f'</div>'
        )
        slide_html_list.append(slide_div)

    return slide_html_list


def _get_slide_titles(pptx_path: str) -> list[str]:
    """Extract slide titles from PPTX for accessibility & navigation."""
    prs = Presentation(pptx_path)
    titles = []
    for i, slide in enumerate(prs.slides, 1):
        title = f"Slide {i}"
        if slide.shapes.title and slide.shapes.title.text.strip():
            title = slide.shapes.title.text.strip()
        titles.append(title)
    return titles


def _get_presentation_title(pptx_path: str) -> str:
    """Get the presentation title from core properties."""
    prs = Presentation(pptx_path)
    core = prs.core_properties
    if core.title and core.title.strip():
        return core.title.strip()
    # Fallback: first slide title
    if prs.slides:
        s = prs.slides[0]
        if s.shapes.title and s.shapes.title.text.strip():
            return s.shapes.title.text.strip()
    return Path(pptx_path).stem.replace("_", " ").replace("-", " ").title()


# ─────────────────────────────────────────────
# imsmanifest.xml (SCORM 2004 3rd Edition)
# ─────────────────────────────────────────────

def _build_manifest(course_id: str, title: str, image_files: list[str]) -> str:
    image_refs = "".join(
        f'<file href="slides/images/{f}"/>\n      ' for f in image_files
    )
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<manifest identifier="{course_id}"
  xmlns="http://www.imsglobal.org/xsd/imscp_v1p1"
  xmlns:adlcp="http://www.adlnet.org/xsd/adlcp_v1p3"
  xmlns:adlseq="http://www.adlnet.org/xsd/adlseq_v1p3"
  xmlns:adlnav="http://www.adlnet.org/xsd/adlnav_v1p3"
  xmlns:imsss="http://www.imsglobal.org/xsd/imsss"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:schemaLocation="http://www.imsglobal.org/xsd/imscp_v1p1 imscp_v1p1.xsd
    http://www.adlnet.org/xsd/adlcp_v1p3 adlcp_v1p3.xsd
    http://www.adlnet.org/xsd/adlseq_v1p3 adlseq_v1p3.xsd
    http://www.adlnet.org/xsd/adlnav_v1p3 adlnav_v1p3.xsd
    http://www.imsglobal.org/xsd/imsss imsss_v1p0.xsd">

  <metadata>
    <schema>ADL SCORM</schema>
    <schemaversion>2004 3rd Edition</schemaversion>
  </metadata>

  <organizations default="ORG-{course_id}">
    <organization identifier="ORG-{course_id}">
      <title>{title}</title>
      <item identifier="ITEM-{course_id}" identifierref="RES-{course_id}">
        <title>{title}</title>
        <adlcp:completionThreshold minProgressMeasure="1.0" completedByMeasure="true"/>
        <imsss:sequencing>
          <imsss:deliveryControls completionSetByContent="true" objectiveSetByContent="true"/>
        </imsss:sequencing>
      </item>
    </organization>
  </organizations>

  <resources>
    <resource identifier="RES-{course_id}" type="webcontent"
      adlcp:scormType="sco" href="index.html">
      <file href="index.html"/>
      <file href="scorm.js"/>
      {image_refs}
    </resource>
  </resources>
</manifest>"""


# ─────────────────────────────────────────────
# Vue 3 SCORM Player (index.html)
# ─────────────────────────────────────────────

def _build_player_html(title: str, slide_titles: list[str], slide_html_list: list[str]) -> str:
    slides_json = json.dumps([
        {"content": h, "title": t}
        for h, t in zip(slide_html_list, slide_titles)
    ])

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>{title}</title>
  <script src="https://unpkg.com/vue@3/dist/vue.global.prod.js"></script>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

    body {{
      font-family: 'Segoe UI', system-ui, sans-serif;
      background: #0f172a;
      color: #e2e8f0;
      height: 100vh;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }}

    /* ── Header ── */
    .header {{
      background: #1e293b;
      border-bottom: 1px solid #334155;
      padding: 10px 20px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-shrink: 0;
      gap: 12px;
    }}
    .course-title {{
      font-size: 15px;
      font-weight: 600;
      color: #f1f5f9;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      max-width: 400px;
    }}
    .progress-wrap {{
      display: flex;
      align-items: center;
      gap: 10px;
      flex: 1;
      max-width: 400px;
    }}
    .progress-bar {{
      flex: 1;
      height: 6px;
      background: #334155;
      border-radius: 3px;
      overflow: hidden;
    }}
    .progress-fill {{
      height: 100%;
      background: linear-gradient(90deg, #6366f1, #8b5cf6);
      border-radius: 3px;
      transition: width 0.4s ease;
    }}
    .progress-label {{
      font-size: 12px;
      color: #94a3b8;
      white-space: nowrap;
    }}
    .completion-badge {{
      background: #22c55e;
      color: #fff;
      font-size: 11px;
      font-weight: 700;
      padding: 3px 10px;
      border-radius: 20px;
      white-space: nowrap;
    }}

    /* ── Main layout ── */
    .main {{
      display: flex;
      flex: 1;
      min-height: 0;
    }}

    /* ── Sidebar ── */
    .sidebar {{
      width: 220px;
      background: #1e293b;
      border-right: 1px solid #334155;
      overflow-y: auto;
      flex-shrink: 0;
    }}
    .sidebar-label {{
      font-size: 10px;
      font-weight: 700;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: #64748b;
      padding: 12px 14px 6px;
    }}
    .slide-thumb {{
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 8px 14px;
      cursor: pointer;
      border-left: 3px solid transparent;
      transition: background 0.15s, border-color 0.15s;
    }}
    .slide-thumb:hover {{ background: #263548; }}
    .slide-thumb.active {{
      background: #1d3a5e;
      border-left-color: #6366f1;
    }}
    .slide-thumb.visited .thumb-num {{
      background: #6366f1;
      color: #fff;
    }}
    .thumb-num {{
      width: 26px;
      height: 26px;
      border-radius: 50%;
      background: #334155;
      font-size: 11px;
      font-weight: 700;
      display: flex;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
      transition: background 0.2s;
    }}
    .thumb-title {{
      font-size: 12px;
      color: #94a3b8;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .slide-thumb.active .thumb-title {{ color: #e2e8f0; }}

    /* ── Slide viewer ── */
    .viewer {{
      flex: 1;
      display: flex;
      flex-direction: column;
      align-items: center;
      padding: 16px 20px;
      overflow-y: auto;
      min-height: 0;
    }}
    .slide-container {{
      position: relative;
      width: 100%;
      max-width: 960px;
      background: #fff;
      border-radius: 8px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.5);
      overflow: auto;
      flex-shrink: 0;
    }}

    /* ── Slide HTML content ── */
    .slide-html-content {{
      width: 100%;
    }}
    .slide-html-content .slide-content {{
      position: relative;
      width: 100%;
      color: #111111;
      font-family: 'Segoe UI', Calibri, Arial, sans-serif;
    }}
    .slide-html-content .slide-content h1 {{
      margin: 0;
      line-height: 1.2;
    }}
    .slide-html-content .slide-content h2 {{
      margin: 0;
      line-height: 1.2;
    }}
    .slide-html-content .slide-content p {{
      margin: 0;
      line-height: 1.4;
    }}
    .slide-html-content .slide-content img {{
      max-width: 100%;
      display: block;
    }}
    .slide-html-content .slide-content table {{
      border-collapse: collapse;
    }}

    /* ── Controls ── */
    .controls {{
      display: flex;
      align-items: center;
      gap: 12px;
      margin-top: 14px;
      flex-shrink: 0;
    }}
    .btn {{
      padding: 9px 20px;
      border-radius: 8px;
      border: none;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.15s;
      display: flex;
      align-items: center;
      gap: 6px;
    }}
    .btn:disabled {{
      opacity: 0.35;
      cursor: not-allowed;
    }}
    .btn-primary {{
      background: #6366f1;
      color: #fff;
    }}
    .btn-primary:not(:disabled):hover {{ background: #4f46e5; }}
    .btn-ghost {{
      background: #1e293b;
      color: #94a3b8;
      border: 1px solid #334155;
    }}
    .btn-ghost:not(:disabled):hover {{ background: #263548; color: #e2e8f0; }}

    .slide-counter {{
      font-size: 13px;
      color: #64748b;
      min-width: 90px;
      text-align: center;
    }}

    /* ── Completed overlay ── */
    .complete-overlay {{
      position: absolute;
      inset: 0;
      background: rgba(15,23,42,0.85);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 16px;
      border-radius: 8px;
      backdrop-filter: blur(4px);
      z-index: 10;
    }}
    .complete-icon {{
      font-size: 56px;
    }}
    .complete-title {{
      font-size: 22px;
      font-weight: 700;
      color: #22c55e;
    }}
    .complete-sub {{
      font-size: 14px;
      color: #94a3b8;
    }}
    .btn-restart {{
      margin-top: 8px;
      padding: 10px 28px;
      background: #22c55e;
      color: #fff;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
    }}
    .btn-restart:hover {{ background: #16a34a; }}

    /* Scrollbar */
    .sidebar::-webkit-scrollbar {{ width: 4px; }}
    .sidebar::-webkit-scrollbar-track {{ background: transparent; }}
    .sidebar::-webkit-scrollbar-thumb {{ background: #334155; border-radius: 2px; }}
  </style>
</head>
<body>
<div id="app">
  <!-- Header -->
  <div class="header">
    <span class="course-title">{title}</span>
    <div class="progress-wrap">
      <div class="progress-bar">
        <div class="progress-fill" :style="{{width: progressPct + '%'}}"></div>
      </div>
      <span class="progress-label">{{{{progressPct}}}}%</span>
    </div>
    <span v-if="completed" class="completion-badge">✓ Completed</span>
  </div>

  <!-- Body -->
  <div class="main">
    <!-- Sidebar -->
    <div class="sidebar">
      <div class="sidebar-label">Slides</div>
      <div v-for="(slide, i) in slides" :key="i"
           class="slide-thumb"
           :class="{{active: current === i, visited: visited.has(i)}}"
           @click="goTo(i)">
        <div class="thumb-num">{{{{i + 1}}}}</div>
        <div class="thumb-title">{{{{slide.title}}}}</div>
      </div>
    </div>

    <!-- Viewer -->
    <div class="viewer">
      <div class="slide-container">
        <div class="slide-html-content" v-html="slides[current].content"></div>

        <!-- Completion overlay -->
        <div v-if="showComplete" class="complete-overlay">
          <div class="complete-icon">🎉</div>
          <div class="complete-title">Course Complete!</div>
          <div class="complete-sub">You've viewed all {{{{slides.length}}}} slides.</div>
          <button class="btn-restart" @click="restart">Review Again</button>
        </div>
      </div>

      <!-- Controls -->
      <div class="controls">
        <button class="btn btn-ghost" @click="prev" :disabled="current === 0">
          ← Previous
        </button>
        <span class="slide-counter">{{{{current + 1}}}} / {{{{slides.length}}}}</span>
        <button class="btn btn-primary" @click="next"
                :disabled="current === slides.length - 1 && completed">
          {{{{current === slides.length - 1 ? 'Finish' : 'Next →'}}}}
        </button>
      </div>
    </div>
  </div>
</div>

<script src="scorm.js"></script>
<script>
const {{ createApp, ref, computed, onMounted, onBeforeUnmount }} = Vue;

const SLIDES = {slides_json};

createApp({{
  setup() {{
    const slides = ref(SLIDES);
    const current = ref(0);
    const visited = ref(new Set([0]));
    const showComplete = ref(false);
    const completed = ref(false);

    const progressPct = computed(() => {{
      const pct = Math.round((visited.value.size / slides.value.length) * 100);
      return pct;
    }});

    function saveBookmark() {{
      SCORM.setLocation(String(current.value));
      const data = JSON.stringify(Array.from(visited.value));
      SCORM.setSuspendData(data);
    }}

    function markVisited(idx) {{
      visited.value.add(idx);
      const measure = visited.value.size / slides.value.length;
      SCORM.setProgress(measure);
      saveBookmark();

      if (visited.value.size === slides.value.length && !completed.value) {{
        completed.value = true;
        SCORM.complete();
        showComplete.value = true;
      }}
    }}

    function goTo(idx) {{
      current.value = idx;
      markVisited(idx);
      showComplete.value = false;
    }}

    function next() {{
      if (current.value < slides.value.length - 1) {{
        goTo(current.value + 1);
      }} else {{
        markVisited(current.value);
        if (completed.value) showComplete.value = true;
      }}
    }}

    function prev() {{
      if (current.value > 0) goTo(current.value - 1);
    }}

    function restart() {{
      showComplete.value = false;
      goTo(0);
    }}

    onMounted(() => {{
      SCORM.init();

      // Restore bookmark from previous session
      const savedLoc = SCORM.getLocation();
      const savedData = SCORM.getSuspendData();

      if (savedData) {{
        try {{
          const arr = JSON.parse(savedData);
          if (Array.isArray(arr)) {{
            arr.forEach(i => visited.value.add(i));
          }}
        }} catch(e) {{ /* ignore corrupt data */ }}
      }}

      const resumeSlide = savedLoc ? parseInt(savedLoc, 10) : 0;
      if (resumeSlide >= 0 && resumeSlide < slides.value.length) {{
        current.value = resumeSlide;
      }}

      // Update progress from restored visited set
      markVisited(current.value);
    }});

    onBeforeUnmount(() => {{
      SCORM.terminate();
    }});

    window.addEventListener('beforeunload', () => SCORM.terminate());

    return {{ slides, current, visited, progressPct, showComplete, completed,
              goTo, next, prev, restart }};
  }}
}}).mount('#app');
</script>
</body>
</html>"""


# ─────────────────────────────────────────────
# SCORM 2004 API Wrapper (scorm.js)
# ─────────────────────────────────────────────

SCORM_JS = """
/**
 * SCORM 2004 3rd Edition API Wrapper
 * Finds the LMS API in parent frames and exposes a clean interface.
 */
const SCORM = (() => {
  let _api = null;
  let _initialized = false;
  let _terminated = false;

  function findAPI(win) {
    let attempts = 0;
    while (win && !win.API_1484_11 && attempts < 10) {
      if (win.parent && win.parent !== win) {
        win = win.parent;
      } else {
        break;
      }
      attempts++;
    }
    return (win && win.API_1484_11) ? win.API_1484_11 : null;
  }

  function getAPI() {
    if (_api) return _api;
    _api = findAPI(window);
    if (!_api && window.opener) _api = findAPI(window.opener);
    return _api;
  }

  function init() {
    if (_initialized || _terminated) return;
    const api = getAPI();
    if (!api) {
      console.warn('[SCORM] No LMS API found – running in standalone mode.');
      return;
    }
    const result = api.Initialize('');
    if (result === 'true' || result === true) {
      _initialized = true;
      // Only set defaults if not resuming a previous session
      const existingStatus = api.GetValue('cmi.completion_status');
      if (!existingStatus || existingStatus === 'not attempted' || existingStatus === 'unknown') {
        api.SetValue('cmi.completion_status', 'incomplete');
        api.SetValue('cmi.success_status', 'unknown');
        api.SetValue('cmi.progress_measure', '0');
      }
      api.Commit('');
      console.log('[SCORM] Initialized.');
    } else {
      console.error('[SCORM] Initialize failed:', api.GetLastError(), api.GetErrorString(api.GetLastError()));
    }
  }

  function setProgress(measure) {
    if (!_initialized || _terminated) return;
    const api = getAPI();
    if (!api) return;
    const val = Math.min(1, Math.max(0, measure));
    api.SetValue('cmi.progress_measure', String(val.toFixed(4)));
    api.Commit('');
    console.log('[SCORM] Progress:', (val * 100).toFixed(1) + '%');
  }

  function setLocation(loc) {
    if (!_initialized || _terminated) return;
    const api = getAPI();
    if (!api) return;
    api.SetValue('cmi.location', String(loc));
    api.Commit('');
  }

  function getLocation() {
    if (!_initialized) return '';
    const api = getAPI();
    if (!api) return '';
    return api.GetValue('cmi.location') || '';
  }

  function setSuspendData(data) {
    if (!_initialized || _terminated) return;
    const api = getAPI();
    if (!api) return;
    api.SetValue('cmi.suspend_data', data);
    api.Commit('');
  }

  function getSuspendData() {
    if (!_initialized) return '';
    const api = getAPI();
    if (!api) return '';
    return api.GetValue('cmi.suspend_data') || '';
  }

  function complete() {
    if (!_initialized || _terminated) return;
    const api = getAPI();
    if (!api) return;
    api.SetValue('cmi.progress_measure', '1');
    api.SetValue('cmi.completion_status', 'completed');
    api.SetValue('cmi.success_status', 'passed');
    api.Commit('');
    console.log('[SCORM] Completed!');
  }

  function terminate() {
    if (!_initialized || _terminated) return;
    const api = getAPI();
    if (!api) return;
    _terminated = true;
    api.Terminate('');
    console.log('[SCORM] Terminated.');
  }

  return { init, setProgress, setLocation, getLocation, setSuspendData, getSuspendData, complete, terminate };
})();
"""


# ─────────────────────────────────────────────
# Main Converter Entry Point
# ─────────────────────────────────────────────

def convert_pptx_to_scorm(
    pptx_path: str,
    output_dir: str,
    course_title: Optional[str] = None,
) -> str:
    """
    Convert a PPTX/PPT file to a SCORM 2004 ZIP package.

    Args:
        pptx_path: Absolute path to the .pptx/.ppt file
        output_dir: Directory where the ZIP will be written
        course_title: Override the course title (uses file metadata if None)

    Returns:
        Path to the generated .zip file
    """
    with tempfile.TemporaryDirectory() as work_dir:
        # 1. Extract metadata
        title = course_title or _get_presentation_title(pptx_path)
        slide_titles = _get_slide_titles(pptx_path)

        # 2. Extract slides as HTML
        slide_html_list = _extract_slide_html(pptx_path, work_dir)
        if not slide_html_list:
            raise RuntimeError("No slides were found in the presentation.")

        # Pad slide titles if count differs
        while len(slide_titles) < len(slide_html_list):
            slide_titles.append(f"Slide {len(slide_titles) + 1}")

        # 3. Build SCORM package directory
        pkg_dir = os.path.join(work_dir, "package")
        os.makedirs(pkg_dir)

        # 3a. Copy extracted images to package
        src_images = os.path.join(work_dir, "slides", "images")
        dst_images = os.path.join(pkg_dir, "slides", "images")
        image_files = []
        if os.path.exists(src_images) and os.listdir(src_images):
            shutil.copytree(src_images, dst_images)
            image_files = sorted(os.listdir(dst_images))

        # 3b. Write imsmanifest.xml
        course_id = re.sub(r"[^A-Za-z0-9]", "_", title)[:40].upper()
        manifest = _build_manifest(course_id, title, image_files)
        with open(os.path.join(pkg_dir, "imsmanifest.xml"), "w", encoding="utf-8") as f:
            f.write(manifest)

        # 3c. Write Vue player (index.html)
        player_html = _build_player_html(title, slide_titles, slide_html_list)
        with open(os.path.join(pkg_dir, "index.html"), "w", encoding="utf-8") as f:
            f.write(player_html)

        # 3d. Write scorm.js
        with open(os.path.join(pkg_dir, "scorm.js"), "w", encoding="utf-8") as f:
            f.write(SCORM_JS)

        # 4. Zip it up
        safe_name = re.sub(r"[^A-Za-z0-9_\-]", "_", title)[:60]
        zip_name = f"{safe_name}_SCORM.zip"
        zip_path = os.path.join(output_dir, zip_name)

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(pkg_dir):
                for file in files:
                    abs_path = os.path.join(root, file)
                    arcname = os.path.relpath(abs_path, pkg_dir)
                    zf.write(abs_path, arcname)

        return zip_path
