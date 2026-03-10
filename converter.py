"""
PPTX/PPT → SCORM 2004 3rd Edition Converter
Renders slides as SVG via LibreOffice + pdftocairo, wraps in a Vue 3 SCORM player package.
"""
import os
import re
import json
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Optional

from pptx import Presentation


# ─────────────────────────────────────────────
# .ppt → .pptx conversion (LibreOffice)
# ─────────────────────────────────────────────

def _convert_ppt_to_pptx(ppt_path: str) -> str:
    """Convert a .ppt file to .pptx using LibreOffice headless. Returns path to .pptx."""
    out_dir = tempfile.mkdtemp(prefix="ppt_convert_")
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pptx", "--outdir", out_dir, ppt_path],
        capture_output=True, text=True, timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr.strip()}")

    pptx_name = Path(ppt_path).stem + ".pptx"
    pptx_path = os.path.join(out_dir, pptx_name)
    if not os.path.exists(pptx_path):
        raise RuntimeError(f"LibreOffice did not produce expected output: {pptx_name}")
    return pptx_path


# ─────────────────────────────────────────────
# SVG Slide Extraction (LibreOffice + pdftocairo)
# ─────────────────────────────────────────────

def _extract_slide_svgs(pptx_path: str, work_dir: str) -> list[str]:
    """Convert each slide to SVG via LibreOffice + pdftocairo.

    Pipeline: PPTX → PDF (LibreOffice) → SVG per page (pdftocairo)
    SVGs are saved to work_dir/slides/ and returned as <img> tag HTML strings.
    """
    slides_dir = os.path.join(work_dir, "slides")
    os.makedirs(slides_dir, exist_ok=True)

    # 1. PPTX → PDF via LibreOffice
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", work_dir, pptx_path],
        capture_output=True, text=True, timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice PDF conversion failed: {result.stderr.strip()}")

    pdf_name = Path(pptx_path).stem + ".pdf"
    pdf_path = os.path.join(work_dir, pdf_name)
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"LibreOffice did not produce expected PDF: {pdf_name}")

    # 2. Get slide count from python-pptx
    prs = Presentation(pptx_path)
    num_slides = len(prs.slides)

    # 3. PDF → SVG per page via pdftocairo
    slide_content_list = []
    for page in range(1, num_slides + 1):
        svg_name = f"slide-{page:02d}"
        svg_path = os.path.join(slides_dir, svg_name)
        result = subprocess.run(
            ["pdftocairo", "-svg", "-f", str(page), "-l", str(page), pdf_path, svg_path],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            raise RuntimeError(f"pdftocairo failed on page {page}: {result.stderr.strip()}")

        # pdftocairo appends .svg to the output name
        rel_path = f"slides/{svg_name}.svg"
        slide_content_list.append(
            f'<img src="{rel_path}" class="slide-svg-img" alt="Slide {page}" />'
        )

    return slide_content_list


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


def _slides_have_titles(pptx_path: str) -> list[bool]:
    """Return a boolean per slide indicating whether it has a real title shape."""
    prs = Presentation(pptx_path)
    return [
        bool(slide.shapes.title and slide.shapes.title.text.strip())
        for slide in prs.slides
    ]


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

def _build_manifest(course_id: str, title: str, svg_files: list[str]) -> str:
    svg_refs = "".join(
        f'<file href="slides/{f}"/>\n      ' for f in svg_files
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
      {svg_refs}
    </resource>
  </resources>
</manifest>"""


# ─────────────────────────────────────────────
# Vue 3 SCORM Player (index.html)
# ─────────────────────────────────────────────

def _build_player_html(title: str, slide_titles: list[str], slide_html_list: list[str], slide_counts: list[int] = None, display_mode: str = "paginated", sidebar_flags: list[bool] = None, theme: str = "light") -> str:
    counts = slide_counts or [1] * len(slide_html_list)
    show_flags = sidebar_flags or [True] * len(slide_html_list)
    slides_json = json.dumps([
        {"content": h, "title": t, "slideCount": c, "showInSidebar": s}
        for h, t, c, s in zip(slide_html_list, slide_titles, counts, show_flags)
    ])

    return f"""<!DOCTYPE html>
<html lang="en" data-theme="{theme}">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>{title}</title>
  <script src="https://unpkg.com/vue@3/dist/vue.global.prod.js"></script>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

    [data-theme="light"] {{
      --bg-body: #f8fafc;
      --bg-surface: #ffffff;
      --bg-hover: #f1f5f9;
      --bg-active: #eff6ff;
      --border: #e2e8f0;
      --text-primary: #0f172a;
      --text-secondary: #1e293b;
      --text-muted: #475569;
      --text-dim: #64748b;
      --text-placeholder: #94a3b8;
      --shadow: rgba(0,0,0,0.08);
      --overlay-bg: rgba(248,250,252,0.9);
    }}
    [data-theme="dark"] {{
      --bg-body: #0f172a;
      --bg-surface: #1e293b;
      --bg-hover: #263548;
      --bg-active: #1d3a5e;
      --border: #334155;
      --text-primary: #f1f5f9;
      --text-secondary: #e2e8f0;
      --text-muted: #94a3b8;
      --text-dim: #64748b;
      --text-placeholder: #475569;
      --shadow: rgba(0,0,0,0.5);
      --overlay-bg: rgba(15,23,42,0.85);
    }}

    html, body {{
      height: 100%;
      overflow: hidden;
    }}

    body {{
      font-family: 'Segoe UI', system-ui, sans-serif;
      background: var(--bg-body);
      color: var(--text-secondary);
      display: flex;
      flex-direction: column;
    }}

    #app {{
      display: flex;
      flex-direction: column;
      height: 100%;
      overflow: hidden;
    }}

    /* ── Header ── */
    .header {{
      background: var(--bg-surface);
      border-bottom: 1px solid var(--border);
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
      color: var(--text-primary);
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
      background: var(--border);
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
      color: var(--text-muted);
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
      position: relative;
    }}

    /* ── Sidebar ── */
    .sidebar {{
      width: 220px;
      background: var(--bg-surface);
      border-right: 1px solid var(--border);
      overflow-y: auto;
      flex-shrink: 0;
      transition: width 0.2s ease, padding 0.2s ease;
    }}
    .sidebar.collapsed {{
      width: 0;
      overflow: hidden;
      border-right: none;
    }}
    .sidebar-header {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 14px 4px;
    }}
    .sidebar-label {{
      font-size: 10px;
      font-weight: 700;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: var(--text-dim);
    }}
    /* ── Hamburger menu button ── */
    .hamburger-btn {{
      background: none;
      border: none;
      cursor: pointer;
      padding: 6px;
      border-radius: 4px;
      display: flex;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
      transition: background 0.15s;
    }}
    .hamburger-btn:hover {{
      background: var(--border);
    }}
    .hamburger-icon {{
      display: block;
      width: 18px;
      height: 2px;
      background: var(--text-muted);
      position: relative;
      transition: background 0.15s;
    }}
    .hamburger-icon::before,
    .hamburger-icon::after {{
      content: '';
      display: block;
      width: 18px;
      height: 2px;
      background: var(--text-muted);
      position: absolute;
      left: 0;
      transition: background 0.15s;
    }}
    .hamburger-icon::before {{ top: -6px; }}
    .hamburger-icon::after {{ top: 6px; }}
    .hamburger-btn:hover .hamburger-icon,
    .hamburger-btn:hover .hamburger-icon::before,
    .hamburger-btn:hover .hamburger-icon::after {{
      background: var(--text-primary);
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
    .slide-thumb:hover {{ background: var(--bg-hover); }}
    .slide-thumb.active {{
      background: var(--bg-active);
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
      background: var(--border);
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
      color: var(--text-muted);
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .slide-thumb.active .thumb-title {{ color: var(--text-secondary); }}

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
      border-radius: 8px;
      box-shadow: 0 20px 60px var(--shadow);
      overflow: hidden;
      flex-shrink: 0;
    }}

    /* ── SVG slide content ── */
    .slide-svg-content {{
      width: 100%;
    }}
    .slide-svg-content .slide-svg-img {{
      width: 100%;
      height: auto;
      display: block;
    }}
    .slide-svg-content .slide-svg-img + .slide-svg-img {{
      margin-top: 4px;
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
      background: var(--bg-surface);
      color: var(--text-muted);
      border: 1px solid var(--border);
    }}
    .btn-ghost:not(:disabled):hover {{ background: var(--bg-hover); color: var(--text-secondary); }}

    .slide-counter {{
      font-size: 13px;
      color: var(--text-dim);
      min-width: 90px;
      text-align: center;
    }}

    /* ── Completed overlay ── */
    .complete-overlay {{
      position: absolute;
      inset: 0;
      background: var(--overlay-bg);
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
      color: var(--text-muted);
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

    /* ── Close button ── */
    .btn-close {{
      background: transparent;
      border: none;
      color: var(--text-muted);
      font-size: 18px;
      cursor: pointer;
      padding: 4px 8px;
      border-radius: 4px;
      line-height: 1;
      flex-shrink: 0;
      transition: color 0.15s, background 0.15s;
    }}
    .btn-close:hover {{
      color: var(--text-primary);
      background: var(--border);
    }}

    /* Scrollbar */
    .sidebar::-webkit-scrollbar {{ width: 4px; }}
    .sidebar::-webkit-scrollbar-track {{ background: transparent; }}
    .sidebar::-webkit-scrollbar-thumb {{ background: var(--border); border-radius: 2px; }}

    /* ── Continuous scroll mode ── */
    .viewer.continuous {{
      align-items: center;
    }}
    .viewer.continuous .continuous-slide {{
      margin-bottom: 24px;
    }}
  </style>
</head>
<body>
<div id="app">
  <!-- Header -->
  <div class="header">
    <button v-if="showSidebar" class="hamburger-btn" @click="sidebarOpen = !sidebarOpen" title="Toggle slides menu">
      <span class="hamburger-icon"></span>
    </button>
    <span class="course-title">{title}</span>
    <div class="progress-wrap">
      <div class="progress-bar">
        <div class="progress-fill" :style="{{width: progressPct + '%'}}"></div>
      </div>
      <span class="progress-label">{{{{progressPct}}}}%</span>
    </div>
    <span v-if="completed" class="completion-badge">✓ Completed</span>
    <button class="btn-close" @click="closeCourse" title="Close course">✕</button>
  </div>

  <!-- Body -->
  <div class="main">
    <!-- Sidebar -->
    <div v-if="showSidebar" class="sidebar" :class="{{collapsed: !sidebarOpen}}">
      <div class="sidebar-header">
        <div class="sidebar-label">Slides</div>
      </div>
      <div v-for="(slide, i) in slides" :key="i"
           v-show="slide.showInSidebar"
           class="slide-thumb"
           :class="{{active: isActiveGroup(i), visited: visited.has(i)}}"
           @click="goTo(i)">
        <div class="thumb-num">{{{{sidebarNumber(i)}}}}</div>
        <div class="thumb-title">{{{{slide.title}}}}</div>
      </div>
    </div>

    <!-- Paginated viewer -->
    <div v-if="displayMode === 'paginated'" class="viewer">
      <div class="slide-container">
        <div class="slide-svg-content" v-html="slides[current].content"></div>

        <!-- Completion overlay -->
        <div v-if="showComplete" class="complete-overlay">
          <div class="complete-icon">🎉</div>
          <div class="complete-title">Course Complete!</div>
          <div class="complete-sub">You've viewed all {{{{totalSlides}}}} slides.</div>
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

    <!-- Continuous scroll viewer -->
    <div v-else class="viewer continuous" ref="scrollViewer">
      <div v-for="(slide, i) in slides" :key="i"
           :id="'slide-' + i"
           class="slide-container continuous-slide">
        <div class="slide-svg-content" v-html="slide.content"></div>
      </div>
    </div>
  </div>
</div>

<script src="scorm.js"></script>
<script>
const {{ createApp, ref, computed, onMounted, onBeforeUnmount, nextTick }} = Vue;

const SLIDES = {slides_json};
const DISPLAY_MODE = '{display_mode}';

createApp({{
  setup() {{
    const slides = ref(SLIDES);
    const current = ref(0);
    const visited = ref(new Set([0]));
    const showComplete = ref(false);
    const completed = ref(false);
    const sidebarOpen = ref(true);
    const displayMode = ref(DISPLAY_MODE);
    const scrollViewer = ref(null);

    const totalSlides = slides.value.reduce((sum, s) => sum + s.slideCount, 0);
    const showSidebar = computed(() => slides.value.filter(s => s.showInSidebar).length > 1);
    const progressPct = computed(() => {{
      let viewedCount = 0;
      visited.value.forEach(i => {{
        viewedCount += slides.value[i].slideCount;
      }});
      return Math.round((viewedCount / totalSlides) * 100);
    }});

    function saveBookmark() {{
      SCORM.setLocation(String(current.value));
      const data = JSON.stringify(Array.from(visited.value));
      SCORM.setSuspendData(data);
    }}

    function markVisited(idx) {{
      visited.value.add(idx);
      let viewedCount = 0;
      visited.value.forEach(i => {{
        viewedCount += slides.value[i].slideCount;
      }});
      const measure = viewedCount / totalSlides;
      SCORM.setProgress(measure);
      saveBookmark();

      if (viewedCount === totalSlides && !completed.value && displayMode.value !== 'continuous') {{
        completed.value = true;
        SCORM.complete();
        showComplete.value = true;
      }}
    }}

    function goTo(idx) {{
      current.value = idx;
      markVisited(idx);
      showComplete.value = false;
      if (displayMode.value === 'continuous') {{
        const el = document.getElementById('slide-' + idx);
        if (el) el.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
      }}
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

    function closeCourse() {{
      SCORM.terminate();
      if (window.parent && window.parent !== window) {{
        window.parent.postMessage({{ type: 'scorm-close' }}, '*');
      }}
      window.close();
    }}

    function sidebarNumber(idx) {{
      // Count only slides with showInSidebar up to and including idx
      let n = 0;
      for (let i = 0; i <= idx; i++) {{
        if (slides.value[i].showInSidebar) n++;
      }}
      return n;
    }}

    function isActiveGroup(idx) {{
      // In paginated mode or if this slide is current, simple match
      if (current.value === idx) return true;
      if (displayMode.value !== 'continuous') return false;
      // In continuous mode, highlight this sidebar item if current slide
      // falls between this item and the next sidebar item
      if (!slides.value[idx].showInSidebar) return false;
      // Find next sidebar-visible slide after idx
      let nextVisible = slides.value.length;
      for (let i = idx + 1; i < slides.value.length; i++) {{
        if (slides.value[i].showInSidebar) {{ nextVisible = i; break; }}
      }}
      return current.value >= idx && current.value < nextVisible;
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

      // Continuous mode: track scroll position for progress and completion
      if (displayMode.value === 'continuous') {{
        nextTick(() => {{
          const viewer = scrollViewer.value;
          if (!viewer) return;

          // Don't observe slides until first scroll — prevents auto-marking on load
          let observerStarted = false;
          const observer = new IntersectionObserver((entries) => {{
            entries.forEach(entry => {{
              if (entry.isIntersecting) {{
                const idx = parseInt(entry.target.id.replace('slide-', ''), 10);
                if (!isNaN(idx)) {{
                  current.value = idx;
                  markVisited(idx);
                }}
              }}
            }});
          }}, {{
            root: viewer,
            threshold: 0.3,
          }});

          const startObserving = () => {{
            if (observerStarted) return;
            observerStarted = true;
            slides.value.forEach((_, i) => {{
              const el = document.getElementById('slide-' + i);
              if (el) observer.observe(el);
            }});
          }};

          viewer.addEventListener('scroll', () => {{
            startObserving();
            // Complete when scrolled to the bottom
            if (!completed.value && viewer.scrollTop + viewer.clientHeight >= viewer.scrollHeight - 20) {{
              for (let i = 0; i < slides.value.length; i++) {{
                visited.value.add(i);
              }}
              SCORM.setProgress(1);
              completed.value = true;
              SCORM.complete();
              saveBookmark();
            }}
          }});
        }});
      }}
    }});

    onBeforeUnmount(() => {{
      SCORM.terminate();
    }});

    window.addEventListener('beforeunload', () => SCORM.terminate());

    return {{ slides, current, visited, progressPct, showComplete, completed,
              sidebarOpen, showSidebar, totalSlides, displayMode, scrollViewer,
              goTo, next, prev, restart, closeCourse,
              sidebarNumber, isActiveGroup }};
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
    display_mode: str = "paginated",
    theme: str = "light",
) -> str:
    """
    Convert a PPTX/PPT file to a SCORM 2004 ZIP package.

    Args:
        pptx_path: Absolute path to the .pptx/.ppt file
        output_dir: Directory where the ZIP will be written
        course_title: Override the course title (uses file metadata if None)
        display_mode: 'paginated' (one slide at a time) or 'continuous' (scrollable)
        theme: 'light' or 'dark' color scheme for the player

    Returns:
        Path to the generated .zip file
    """
    with tempfile.TemporaryDirectory() as work_dir:
        # 0. Convert .ppt to .pptx if needed (python-pptx only reads .pptx)
        converted_pptx = None
        if pptx_path.lower().endswith(".ppt"):
            pptx_path = _convert_ppt_to_pptx(pptx_path)
            converted_pptx = pptx_path  # track for cleanup

        # 1. Extract metadata
        title = course_title or _get_presentation_title(pptx_path)
        slide_titles = _get_slide_titles(pptx_path)
        has_title_flags = _slides_have_titles(pptx_path)

        # 2. Convert slides to SVG (PPTX → PDF → SVG per page)
        slide_html_list = _extract_slide_svgs(pptx_path, work_dir)
        if not slide_html_list:
            raise RuntimeError("No slides were found in the presentation.")

        # Pad slide titles if count differs
        while len(slide_titles) < len(slide_html_list):
            slide_titles.append(f"Slide {len(slide_titles) + 1}")
        while len(has_title_flags) < len(slide_html_list):
            has_title_flags.append(False)

        # 2b. Group slides: merge headless slides into previous headed group
        # (only in paginated mode — continuous mode needs individual slides for scroll tracking)
        if display_mode == "paginated":
            grouped = []
            for i, (s_html, s_title, has_t) in enumerate(
                zip(slide_html_list, slide_titles, has_title_flags)
            ):
                if has_t or not grouped:
                    grouped.append({"title": s_title, "content": s_html, "count": 1})
                else:
                    grouped[-1]["content"] += s_html
                    grouped[-1]["count"] += 1

            slide_titles = [g["title"] for g in grouped]
            slide_html_list = [g["content"] for g in grouped]
            slide_counts = [g["count"] for g in grouped]
        else:
            slide_counts = [1] * len(slide_html_list)

        # Build per-slide showInSidebar flags
        # Paginated: all groups show (headless slides already merged)
        # Continuous: only slides with real titles show in sidebar
        if display_mode == "paginated":
            sidebar_flags = [True] * len(slide_html_list)
        else:
            sidebar_flags = list(has_title_flags[:len(slide_html_list)])
            # First slide always shows even if untitled
            if sidebar_flags and not sidebar_flags[0]:
                sidebar_flags[0] = True

        # 3. Build SCORM package directory
        pkg_dir = os.path.join(work_dir, "package")
        os.makedirs(pkg_dir)

        # 3a. Copy SVG slides to package
        src_slides = os.path.join(work_dir, "slides")
        dst_slides = os.path.join(pkg_dir, "slides")
        shutil.copytree(src_slides, dst_slides)
        svg_files = sorted(f for f in os.listdir(dst_slides) if f.endswith(".svg"))

        # 3b. Write imsmanifest.xml
        course_id = re.sub(r"[^A-Za-z0-9]", "_", title)[:40].upper()
        manifest = _build_manifest(course_id, title, svg_files)
        with open(os.path.join(pkg_dir, "imsmanifest.xml"), "w", encoding="utf-8") as f:
            f.write(manifest)

        # 3c. Write Vue player (index.html)
        player_html = _build_player_html(title, slide_titles, slide_html_list, slide_counts, display_mode, sidebar_flags, theme)
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

        # Clean up temp .pptx from .ppt conversion
        if converted_pptx:
            shutil.rmtree(os.path.dirname(converted_pptx), ignore_errors=True)

        return zip_path
