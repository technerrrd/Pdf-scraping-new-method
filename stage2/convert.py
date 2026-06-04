#!/usr/bin/env python3
"""
convert.py — DOCX to LaTeX/LyX converter for pdf-conversion output files.

Usage:
    python convert.py                     # convert all .docx in current dir
    python convert.py Chapter1.docx       # convert specific file
    python convert.py Ch1.docx Ch2.docx   # multiple files
"""

import sys
import os
import re
import shutil
import zipfile
import logging
import textwrap
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

TEMPLATE_DIR = Path(__file__).parent.parent / 'Final-lyx_template'

# ---------------------------------------------------------------------------
# XML namespaces
# ---------------------------------------------------------------------------
W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

# ---------------------------------------------------------------------------
# Font size → heading level (w:sz is in half-points)
# ---------------------------------------------------------------------------
SIZE_MAP = {
    32: 'section',       # 16pt
    28: 'subsection',    # 14pt
}

# Width scale choices for images (fraction of \textwidth)
SCALE_OPTIONS = [0.25, 0.40, 0.50, 0.60, 0.75]

def pick_scale(cx_emu: int) -> float:
    """Choose the closest \textwidth fraction from EMU width."""
    text_width_emu = 6.5 * 914400  # assume 6.5in text width
    ratio = cx_emu / text_width_emu
    return min(SCALE_OPTIONS, key=lambda s: abs(s - ratio))

# ---------------------------------------------------------------------------
# Bold detection
# ---------------------------------------------------------------------------
def is_run_bold(run) -> bool:
    """True only if <w:b> present in rPr and w:val is not 0/false."""
    b = run.find(f'{{{W}}}rPr/{{{W}}}b')
    if b is None:
        return False
    val = b.get(f'{{{W}}}val', '1')
    return val not in ('0', 'false')

def para_segments(para) -> list:
    """
    Return list of (text, is_bold) tuples from a paragraph's runs.
    Consecutive runs with the same bold value are merged into one segment.
    """
    segs = []  # [[text, bold], ...]
    for run in para.findall(f'.//{{{W}}}r'):
        run_texts = [t.text for t in run.findall(f'{{{W}}}t') if t.text]
        if not run_texts:
            continue
        run_text = ''.join(run_texts)
        bold = is_run_bold(run)
        if segs and segs[-1][1] == bold:
            segs[-1][0] += run_text
        else:
            segs.append([run_text, bold])
    return [(s[0], s[1]) for s in segs]

def segments_text(segments) -> str:
    """Extract flat text from a segments list."""
    return ''.join(s[0] for s in segments)

# ---------------------------------------------------------------------------
# Noise / TOC / chapter / heading-numbering detection
# ---------------------------------------------------------------------------
TOC_RE     = re.compile(r'^\d+\.\s+\S')
CHAPTER_RE = re.compile(r'^Chapter\s+Notes?\s*[:\-]\s*(.+)', re.IGNORECASE)
HEADING_NUM_RE = re.compile(
    r'^(?:'
    r'Q\d+\.?\s+'            # Q1 / Q1. / Q12
    r'|\d+(?:\.\d+)*\.?\s+'  # 1. / 1.1 / 1.6.9
    r'|\([a-zA-Z0-9]+\)\s+'  # (a) / (i) / (1)
    r')'
)

def is_toc_line(text: str) -> bool:
    return bool(TOC_RE.match(text.strip()))

def chapter_name(text: str):
    """Return chapter name if text matches 'Chapter Notes: <name>', else None."""
    m = CHAPTER_RE.match(text.strip())
    return m.group(1).strip() if m else None

def strip_heading_numbering(text: str) -> str:
    """Strip leading numeric/alpha prefixes like '1.', '1.1', 'Q1', '(a)'."""
    return HEADING_NUM_RE.sub('', text).strip()

# ---------------------------------------------------------------------------
# LaTeX helpers
# ---------------------------------------------------------------------------
def tex_escape(text: str) -> str:
    """Escape LaTeX special characters."""
    replacements = [
        ('\\', r'\textbackslash{}'),
        ('&',  r'\&'),
        ('%',  r'\%'),
        ('$',  r'\$'),
        ('#',  r'\#'),
        ('{',  r'\{'),
        ('}',  r'\}'),
        ('~',  r'\textasciitilde{}'),
        ('^',  r'\^{}'),
        ('_',  r'\_'),
        ('₹',  r'\rupee~'),
        ('Rs.', r'\rupee~'),
        ('INR', r'\rupee~'),
    ]
    result = text
    for old, new in replacements:
        if old == '\\':
            result = result.replace(old, new)
            break
    for old, new in replacements[1:]:
        result = result.replace(old, new)
    return result

def render_tex_segments(segments) -> str:
    """Render (text, bold) segments to LaTeX, applying \\textbf{} where bold."""
    parts = []
    for text, bold in segments:
        escaped = tex_escape(text)
        parts.append(f'\\textbf{{{escaped}}}' if bold else escaped)
    return ''.join(parts)

# ---------------------------------------------------------------------------
# Parse DOCX
# ---------------------------------------------------------------------------
def parse_docx(docx_path: Path):
    """
    Returns list of elements in document order.
    Each element is a dict with keys:
      type: 'heading' | 'body' | 'image' | 'mcq_question'
      level: 'section' | 'subsection'  (headings only)
      text: str  (flat text, used for MCQ detection and join conditions)
      segments: list of (text, is_bold)  (body elements only)
      filename: str   (images only)
      scale: float    (images only)
    """
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
        try:
            rels_xml = z.read('word/_rels/document.xml.rels').decode('utf-8')
        except KeyError:
            rels_xml = '<Relationships/>'
        media_files = [f for f in z.namelist() if f.startswith('word/media/')]

    # Build rId → filename
    rid_map = {}
    for rel in ET.fromstring(rels_xml):
        rid = rel.get('Id', '')
        tgt = rel.get('Target', '')
        if 'media' in tgt:
            rid_map[rid] = Path(tgt).name

    root = ET.fromstring(xml)
    body = root.find(f'{{{W}}}body')

    raw = []   # list of (sz, text, segments, img_fname, cx_emu)

    for para in body.findall(f'{{{W}}}p'):
        # --- font size ---
        sz_el = para.find(f'.//{{{W}}}sz')
        try:
            sz = int(sz_el.get(f'{{{W}}}val')) if sz_el is not None else 24
        except (TypeError, ValueError):
            sz = 24

        # --- run-level segments (text + bold) ---
        segs = para_segments(para)
        text = segments_text(segs).strip()

        # --- inline image ---
        blip = para.find(f'.//{{{A}}}blip')
        img_fname = None
        cx_emu = 0
        if blip is not None:
            rid = blip.get(f'{{{R}}}embed', '')
            img_fname = rid_map.get(rid)
            ext = para.find(f'.//{{{WP}}}extent')
            if ext is not None:
                try:
                    cx_emu = int(ext.get('cx', 0))
                except (TypeError, ValueError):
                    cx_emu = 0

        raw.append((sz, text, segs, img_fname, cx_emu))

    # --- Skip TOC block ---
    in_toc = False
    elements = []
    seen_first_heading = False

    for sz, text, segs, img_fname, cx_emu in raw:
        # Image always included
        if img_fname:
            scale = pick_scale(cx_emu)
            elements.append({'type': 'image', 'filename': img_fname,
                              'scale': scale, 'text': ''})
            continue

        if not text:
            continue

        level = SIZE_MAP.get(sz)

        # First sz=32 that is NOT a chapter marker = document title → skip + enter TOC zone
        if level == 'section' and not seen_first_heading and not chapter_name(text):
            seen_first_heading = True
            in_toc = True
            continue  # skip title
        seen_first_heading = True

        # In TOC zone: skip numbered TOC entries
        if in_toc:
            if is_toc_line(text):
                continue
            else:
                in_toc = False  # TOC ended

        # Classify — chapter pattern wins over font-size level
        ch = chapter_name(text)
        if ch:
            elements.append({'type': 'heading', 'level': 'chapter',
                              'text': strip_heading_numbering(ch)})
        elif level == 'section':
            elements.append({'type': 'heading', 'level': 'section',
                              'text': strip_heading_numbering(text)})
        elif level == 'subsection':
            elements.append({'type': 'heading', 'level': 'subsection',
                              'text': strip_heading_numbering(text)})
        else:
            # MCQ question detection
            if text.startswith('Try yourself:') or text.startswith('Try yourself :'):
                question = re.sub(r'^Try yourself\s*:\s*', '', text).strip()
                elements.append({'type': 'mcq_question', 'text': question,
                                  'segments': segs})
            else:
                elements.append({'type': 'body', 'text': text, 'segments': segs})

    return elements, rid_map, media_files

# ---------------------------------------------------------------------------
# Reconstruct paragraphs (join split lines)
# ---------------------------------------------------------------------------
def reconstruct_paragraphs(elements):
    """Join consecutive body lines that are continuations of the same paragraph."""
    out = []
    buf = None

    def flush():
        nonlocal buf
        if buf:
            out.append(buf)
            buf = None

    for el in elements:
        if el['type'] != 'body':
            flush()
            out.append(el)
            continue

        text = el['text']
        if buf is None:
            buf = dict(el)
            buf['segments'] = list(el['segments'])
        else:
            prev_text = buf['text']
            # Join if previous line doesn't end a sentence and current looks like continuation
            if (not prev_text.endswith(('.', ':', '?', '!'))
                    and (text and (text[0].islower() or text[0] == '('))):
                buf['text'] = prev_text + ' ' + text
                buf['segments'] = buf['segments'] + [(' ', False)] + list(el['segments'])
            else:
                flush()
                buf = dict(el)
                buf['segments'] = list(el['segments'])

    flush()
    return out

# ---------------------------------------------------------------------------
# Group into MCQ blocks
# ---------------------------------------------------------------------------
def group_mcq_blocks(elements):
    """
    Detect MCQ blocks: mcq_question followed by exactly 4 body lines (options A-D).
    Returns new element list where MCQ blocks are collapsed into type='mcq'.
    """
    out = []
    i = 0
    while i < len(elements):
        el = elements[i]
        if el['type'] == 'mcq_question':
            # Collect next 4 body elements as options
            opts = []
            j = i + 1
            while j < len(elements) and len(opts) < 4:
                if elements[j]['type'] == 'body':
                    opts.append(elements[j]['text'])
                    j += 1
                else:
                    break
            if len(opts) == 4:
                out.append({'type': 'mcq', 'text': el['text'], 'options': opts})
                i = j
                continue
            else:
                # Not enough options — treat as body with no bold
                out.append({'type': 'body', 'text': el['text'],
                             'segments': el.get('segments', [(el['text'], False)])})
        else:
            out.append(el)
        i += 1
    return out

# ---------------------------------------------------------------------------
# LaTeX writer
# ---------------------------------------------------------------------------
TEX_PREAMBLE = r"""\documentclass[12pt,a4paper]{book}
\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage{graphicx}
\usepackage{amssymb}
\usepackage{enumitem}
\usepackage{rupee}
\usepackage{geometry}
\geometry{margin=2.5cm}
\setlength{\parindent}{0pt}
\setlength{\parskip}{4pt}

\begin{document}
"""

def write_tex(elements, out_path: Path, media_dir: Path):
    lines = [TEX_PREAMBLE]
    in_enum = False

    def close_enum():
        nonlocal in_enum
        if in_enum:
            lines.append(r'\end{enumerate}' + '\n')
            in_enum = False

    def open_enum():
        nonlocal in_enum
        if not in_enum:
            lines.append(r'\begin{enumerate}' + '\n')
            in_enum = True

    for el in elements:
        t = el['type']

        if t == 'heading':
            close_enum()
            lvl = el['level']
            txt = tex_escape(el['text'])
            cmd = {'chapter': 'chapter', 'section': 'section',
                   'subsection': 'subsection', 'subsubsection': 'subsubsection'}.get(lvl, 'subsection')
            lines.append(f'\\{cmd}{{{txt}}}\n')

        elif t == 'body':
            close_enum()
            rendered = render_tex_segments(el['segments'])
            lines.append(f'\\par {rendered}\n\n')

        elif t == 'mcq':
            open_enum()
            q = tex_escape(el['text'])
            opts = [tex_escape(o) for o in el['options']]
            lines.append(f'\\item \\textbf{{{q}}}\\\\[0.13cm]\n')
            lines.append(r'\begin{tabular}{@{}p{0.45\textwidth} p{0.45\textwidth}@{}}' + '\n')
            lines.append(f'$\\square$ A) {opts[0]} & $\\square$ B) {opts[1]} \\\\\n')
            lines.append(f'$\\square$ C) {opts[2]} & $\\square$ D) {opts[3]}\n')
            lines.append(r'\end{tabular}' + '\n')

        elif t == 'image':
            close_enum()
            fname = el['filename']
            scale = el['scale']
            img_rel = f'media/{fname}'
            lines.append('\\begin{figure}[h]\n')
            lines.append('\\centering\n')
            lines.append(f'\\includegraphics[width={scale}\\textwidth]{{{img_rel}}}\n')
            lines.append('\\end{figure}\n\n')

        elif t == 'list':
            close_enum()
            env = 'enumerate' if el.get('ordered') else 'itemize'
            depth = 0
            for d, segs in el['items']:
                while depth < d + 1:
                    lines.append(f'\\begin{{{env}}}\n')
                    depth += 1
                while depth > d + 1:
                    lines.append(f'\\end{{{env}}}\n')
                    depth -= 1
                lines.append(f'\\item {render_tex_segments(segs)}\n')
            while depth > 0:
                lines.append(f'\\end{{{env}}}\n')
                depth -= 1
            lines.append('\n')

        elif t == 'table':
            close_enum()
            rows = el['rows']
            ncol = max(len(r) for r in rows)
            lines.append('\\begin{center}\n')
            lines.append('\\begin{tabular}{|' + 'l|' * ncol + '}\n\\hline\n')
            for row in rows:
                row = row + [''] * (ncol - len(row))
                lines.append(' & '.join(tex_escape(c) for c in row) + ' \\\\\n\\hline\n')
            lines.append('\\end{tabular}\n\\end{center}\n\n')

    close_enum()
    lines.append(r'\end{document}' + '\n')

    out_path.write_text(''.join(lines), encoding='utf-8')

# ---------------------------------------------------------------------------
# LyX writer
# ---------------------------------------------------------------------------
LYX_HEADER = """\
#LyX 2.3 created this file.
\\lyxformat 544
\\begin_document
\\begin_header
\\save_transient_properties true
\\origin unavailable
\\textclass book
\\use_default_options true
\\maintain_unincluded_children false
\\language english
\\language_package default
\\inputencoding auto
\\fontencoding global
\\font_roman "default" "default"
\\font_sans "default" "default"
\\font_typewriter "default" "default"
\\font_math "auto" "auto"
\\font_default_family default
\\use_non_tex_fonts false
\\font_sc false
\\font_osf false
\\font_sf_scale 100 100
\\font_tt_scale 100 100
\\use_microtype false
\\use_dash_ligatures true
\\graphics default
\\default_output_format default
\\output_sync 0
\\bibtex_command default
\\index_command default
\\paperfontsize 12
\\spacing single
\\use_hyperref false
\\papersize a4paper
\\use_geometry true
\\use_package amsmath 1
\\use_package amssymb 1
\\use_package cancel 1
\\use_package esint 1
\\use_package mathdots 1
\\use_package mathtools 1
\\use_package mhchem 1
\\use_package stackrel 1
\\use_package stmaryrd 1
\\use_package undertilde 1
\\cite_engine basic
\\cite_engine_type default
\\biblio_style plain
\\use_bibtopic false
\\use_indices false
\\paperorientation portrait
\\suppress_date false
\\justification true
\\use_refstyle 1
\\use_minted 0
\\index Index
\\shortcut idx
\\color #008000
\\end_index
\\leftmargin 2.5cm
\\topmargin 2.5cm
\\rightmargin 2.5cm
\\bottommargin 2.5cm
\\secnumdepth 3
\\tocdepth 3
\\paragraph_separation indent
\\paragraph_indentation default
\\is_math_indent 0
\\math_numbering_side default
\\quotes_style english
\\dynamic_quotes 0
\\papercolumns 1
\\papersides 1
\\paperpagestyle default
\\tracking_changes false
\\output_changes false
\\html_math_output 0
\\html_css_as_file 0
\\html_be_strict false
\\end_header

\\begin_body
"""

LYX_FOOTER = """\
\\end_body
\\end_document
"""

# ---------------------------------------------------------------------------
# Template integration helpers
# ---------------------------------------------------------------------------
def load_template_prefix(template_dir: Path) -> str:
    """Read main.lyx verbatim up to and including the TOC/pagestyle block."""
    text = (template_dir / 'main.lyx').read_text(encoding='utf-8')
    marker = '% Print headers again'
    idx = text.find(marker)
    if idx == -1:
        raise ValueError("Marker '% Print headers again' not found in main.lyx")
    # Skip past: \end_layout (Plain Layout) → \end_inset → \end_layout (Standard)
    end_inset = text.find('\\end_inset', idx)
    end_layout = text.find('\\end_layout', end_inset)
    return text[:end_layout + len('\\end_layout')] + '\n\n'


def chapterimage_block(filename: str) -> str:
    """Return the LyX ERT block for \\chapterimage{filename}."""
    return (
        '\\begin_layout Standard\n'
        '\\begin_inset ERT\n'
        'status open\n'
        '\n'
        '\\begin_layout Plain Layout\n'
        '\n'
        '\n'
        '\\backslash\n'
        'chapterimage\n'
        '\\end_layout\n'
        '\n'
        '\\end_inset\n'
        '\n'
        '\n'
        '\\begin_inset ERT\n'
        'status collapsed\n'
        '\n'
        '\\begin_layout Plain Layout\n'
        '\n'
        '{\n'
        '\\end_layout\n'
        '\n'
        '\\end_inset\n'
        '\n'
        f'{filename}\n'
        '\\begin_inset ERT\n'
        'status collapsed\n'
        '\n'
        '\\begin_layout Plain Layout\n'
        '\n'
        '}\n'
        '\\end_layout\n'
        '\n'
        '\\end_inset\n'
        '\n'
        ' \n'
        '\\begin_inset ERT\n'
        'status collapsed\n'
        '\n'
        '\\begin_layout Plain Layout\n'
        '\n'
        '% Chapter heading image\n'
        '\\end_layout\n'
        '\n'
        '\\end_inset\n'
        '\n'
        '\n'
        '\\end_layout\n'
        '\n'
    )


# ---------------------------------------------------------------------------
# Chapter banner diagram helpers (matplotlib)
# ---------------------------------------------------------------------------
BG       = '#f8f8f8'
OCRE_HEX = '#F36619'
WHITE    = '#ffffff'   # used only for title text on orange box
FG       = '#1a1a2e'   # lines, labels, axes (was WHITE on dark theme)
LGRAY    = '#555555'   # secondary labels
BLUE     = '#2c6fad'
GREEN    = '#2e7d32'

def _setup_fig():
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(18.4, 9.2), dpi=100)
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis('off')
    return fig, ax

def _add_title(ax, chapter_name: str):
    import matplotlib.pyplot as plt
    ax.text(5, 0.38, chapter_name, ha='center', va='center',
            fontsize=26, fontweight='bold', color=WHITE,
            bbox=dict(facecolor=OCRE_HEX, alpha=0.88,
                      boxstyle='round,pad=0.45', edgecolor='none'),
            zorder=10)

def _draw_electricity(ax):
    import matplotlib.patches as mpatches
    import numpy as np
    # Circuit rectangle: left=1.5, right=8.5, top=3.8, bottom=1.8
    lx, rx, ty, by = 1.5, 8.5, 3.8, 1.8
    cx = (lx + rx) / 2
    # Wires
    ax.plot([lx, rx], [ty, ty], color=OCRE_HEX, lw=3)
    ax.plot([lx, rx], [by, by], color=OCRE_HEX, lw=3)
    ax.plot([lx, lx], [by, ty], color=OCRE_HEX, lw=3)
    ax.plot([rx, rx], [by, ty], color=OCRE_HEX, lw=3)
    # Battery on left vertical (lx)
    bmy = (ty + by) / 2
    for offset, sign in [(-0.18, '+'), (0.18, '−')]:
        ax.plot([lx - 0.15, lx + 0.15], [bmy + offset, bmy + offset],
                color=WHITE, lw=3 if sign == '+' else 1.5)
    ax.text(lx - 0.45, bmy + 0.18, '+', color=WHITE, fontsize=14, ha='center')
    ax.text(lx - 0.45, bmy - 0.18, '−', color=WHITE, fontsize=14, ha='center')
    # Bulb on right vertical (rx) — circle with X
    bulb_y = bmy
    circ = mpatches.Circle((rx, bulb_y), 0.28, color=OCRE_HEX, fill=False, lw=2.5)
    ax.add_patch(circ)
    ax.plot([rx - 0.2, rx + 0.2], [bulb_y - 0.2, bulb_y + 0.2], color=OCRE_HEX, lw=2)
    ax.plot([rx - 0.2, rx + 0.2], [bulb_y + 0.2, bulb_y - 0.2], color=OCRE_HEX, lw=2)
    # Current arrows on top wire
    for xpos in [3.0, 5.0, 7.0]:
        ax.annotate('', xy=(xpos + 0.5, ty), xytext=(xpos, ty),
                    arrowprops=dict(arrowstyle='->', color=FG, lw=1.5))
    # Magnetic field circles above top wire
    for xpos in [3.2, 5.0, 6.8]:
        for r, alpha in [(0.35, 0.8), (0.6, 0.5), (0.85, 0.3)]:
            theta = np.linspace(0, 2 * np.pi, 80)
            ax.plot(xpos + r * np.cos(theta), ty + 0.3 + r * np.sin(theta),
                    color=OCRE_HEX, lw=1, alpha=alpha, linestyle='--')
    ax.text(cx, ty + 1.3, 'Magnetic Field Lines', color=LGRAY,
            fontsize=11, ha='center')
    ax.text(cx, by - 0.35, 'Electric Circuit', color=LGRAY,
            fontsize=11, ha='center')

def _draw_light(ax):
    import numpy as np
    my = 2.5
    # Faint divider
    ax.axvline(x=5.0, color=LGRAY, lw=1, linestyle=':', alpha=0.4)

    # ── LEFT: Concave Mirror ──────────────────────────────────────────
    # Mirror arc at x≈4.2, curving leftward (concave toward incoming light)
    theta_m = np.linspace(-0.55, 0.55, 80)
    mx_arc  = 4.2 - 0.35 * (1 - np.cos(theta_m * 2.8))
    my_arc  = my + 1.45 * np.sin(theta_m)
    ax.plot(mx_arc, my_arc, color=FG, lw=3)
    # Principal axis
    ax.plot([0.4, 4.3], [my, my], color=LGRAY, lw=1, linestyle='--')
    f_m = 1.5
    ax.plot(4.2 - f_m, my, 'o', color=OCRE_HEX, ms=7)
    ax.text(4.2 - f_m, my - 0.32, 'F', color=OCRE_HEX, fontsize=11, ha='center')
    ax.plot(4.2 - 2*f_m, my, 'o', color=LGRAY, ms=5)
    ax.text(4.2 - 2*f_m, my - 0.32, 'C', color=LGRAY, fontsize=10, ha='center')
    # Parallel rays → reflect to F
    for ray_y in [my + 0.75, my, my - 0.75]:
        ax.annotate('', xy=(4.2, ray_y), xytext=(0.6, ray_y),
                    arrowprops=dict(arrowstyle='->', color=OCRE_HEX, lw=1.6))
        ax.plot([4.2, 4.2 - f_m], [ray_y, my], color=OCRE_HEX, lw=1.6)
    ax.text(2.3, 0.55, 'Concave Mirror', color=FG, fontsize=12,
            ha='center', fontweight='bold')

    # ── RIGHT: Convex Lens ────────────────────────────────────────────
    rcx = 7.5
    theta_l = np.linspace(-0.6, 0.6, 60)
    ax.plot(rcx + 0.42 * np.sin(theta_l * 3),
            my + 1.4 * np.sin(theta_l), color=FG, lw=2.5)
    ax.plot(rcx - 0.42 * np.sin(theta_l * 3),
            my + 1.4 * np.sin(theta_l), color=FG, lw=2.5)
    # Principal axis
    ax.plot([5.2, 9.6], [my, my], color=LGRAY, lw=1, linestyle='--')
    f_l = 1.5
    ax.plot(rcx + f_l, my, 'o', color=BLUE, ms=7)
    ax.plot(rcx - f_l, my, 'o', color=BLUE, ms=7)
    ax.text(rcx + f_l, my - 0.32, 'F', color=BLUE, fontsize=11, ha='center')
    # Parallel rays → converge at F
    for ray_y in [my + 0.75, my, my - 0.75]:
        ax.annotate('', xy=(rcx, ray_y), xytext=(5.4, ray_y),
                    arrowprops=dict(arrowstyle='->', color=BLUE, lw=1.6))
        ax.plot([rcx, rcx + f_l], [ray_y, my], color=BLUE, lw=1.6)
    ax.text(7.5, 0.55, 'Convex Lens', color=FG, fontsize=12,
            ha='center', fontweight='bold')

    ax.text(5.0, 4.65, 'Light: Mirrors & Lenses', color=LGRAY,
            fontsize=12, ha='center')

def _draw_particles(ax):
    import numpy as np
    panels = [
        (1.5,  'Solid',  'grid'),
        (5.0,  'Liquid', 'close'),
        (8.5,  'Gas',    'spread'),
    ]
    rng = np.random.default_rng(42)
    for cx, label, ptype in panels:
        # Panel box
        box = __import__('matplotlib').patches.FancyBboxPatch(
            (cx - 1.1, 1.1), 2.2, 2.8,
            boxstyle='round,pad=0.05',
            facecolor='#eeeeee', edgecolor=OCRE_HEX, lw=1.5)
        ax.add_patch(box)
        ax.text(cx, 0.75, label, color=FG, fontsize=13,
                ha='center', fontweight='bold')
        if ptype == 'grid':
            for row in range(4):
                for col in range(4):
                    ax.plot(cx - 0.75 + col * 0.5,
                            1.45 + row * 0.6,
                            'o', color=OCRE_HEX, ms=9)
        elif ptype == 'close':
            positions = [(cx + rng.uniform(-0.8, 0.8),
                          1.45 + rng.uniform(0, 2.4)) for _ in range(16)]
            for px, py in positions:
                ax.plot(px, py, 'o', color=OCRE_HEX, ms=9)
        else:
            positions = [(cx + rng.uniform(-0.85, 0.85),
                          1.45 + rng.uniform(0, 2.4)) for _ in range(8)]
            for px, py in positions:
                ax.plot(px, py, 'o', color=OCRE_HEX, ms=9)
    ax.text(5.0, 4.55, 'Particulate Nature of Matter', color=LGRAY,
            fontsize=12, ha='center')

def _draw_elements(ax):
    import numpy as np
    groups = [
        (1.8,  'Element',  [(OCRE_HEX, OCRE_HEX)]),
        (5.0,  'Compound', [(OCRE_HEX, BLUE)]),
        (8.2,  'Mixture',  [(OCRE_HEX, None), (BLUE, None)]),
    ]
    rng = np.random.default_rng(7)
    for cx, label, colors in groups:
        box = __import__('matplotlib').patches.FancyBboxPatch(
            (cx - 1.2, 1.0), 2.4, 3.0,
            boxstyle='round,pad=0.05',
            facecolor='#eeeeee', edgecolor=OCRE_HEX, lw=1.5)
        ax.add_patch(box)
        ax.text(cx, 0.65, label, color=FG, fontsize=13,
                ha='center', fontweight='bold')
        if label == 'Element':
            for row in range(3):
                for col in range(3):
                    ax.plot(cx - 0.6 + col * 0.6, 1.5 + row * 0.8,
                            'o', color=OCRE_HEX, ms=14)
        elif label == 'Compound':
            for row in range(3):
                for col in range(3):
                    c = OCRE_HEX if (row + col) % 2 == 0 else BLUE
                    ax.plot(cx - 0.6 + col * 0.6, 1.5 + row * 0.8,
                            'o', color=c, ms=14)
                    if col < 2:
                        ax.plot([cx - 0.6 + col * 0.6,
                                 cx - 0.6 + (col + 1) * 0.6],
                                [1.5 + row * 0.8, 1.5 + row * 0.8],
                                color=FG, lw=1.5)
        else:
            for _ in range(12):
                c = OCRE_HEX if rng.random() > 0.5 else BLUE
                ax.plot(cx + rng.uniform(-0.9, 0.9),
                        1.3 + rng.uniform(0, 2.2),
                        'o', color=c, ms=12)
    ax.text(5.0, 4.55, 'Elements, Compounds & Mixtures', color=LGRAY,
            fontsize=12, ha='center')

def _draw_solution(ax):
    import matplotlib.patches as mpatches
    import numpy as np
    rng = np.random.default_rng(3)
    # Beaker outline
    bx, bw, by_bot, by_top = 2.8, 4.4, 0.9, 4.2
    ax.plot([bx, bx, bx + bw, bx + bw], [by_top, by_bot, by_bot, by_top],
            color=FG, lw=3)
    ax.plot([bx - 0.2, bx + bw + 0.2], [by_top, by_top],
            color=FG, lw=2, linestyle='--', alpha=0.4)
    # Solvent (blue) and solute (orange) particles
    for _ in range(40):
        px = bx + 0.3 + rng.random() * (bw - 0.6)
        py = by_bot + 0.3 + rng.random() * (by_top - by_bot - 0.5)
        c = OCRE_HEX if rng.random() < 0.3 else BLUE
        ms = 10 if c == OCRE_HEX else 7
        ax.plot(px, py, 'o', color=c, ms=ms, alpha=0.85)
    # Labels
    ax.text(bx + bw / 2, by_bot - 0.4, 'Solution', color=FG,
            fontsize=13, ha='center', fontweight='bold')
    ax.plot(1.0, 3.2, 'o', color=OCRE_HEX, ms=10)
    ax.text(1.35, 3.2, '= Solute', color=OCRE_HEX, fontsize=11, va='center')
    ax.plot(1.0, 2.7, 'o', color=BLUE, ms=7)
    ax.text(1.35, 2.7, '= Solvent', color=BLUE, fontsize=11, va='center')
    ax.text(5.0, 4.6, 'Solutes, Solvents & Solutions', color=LGRAY,
            fontsize=12, ha='center')

def _draw_health(ax):
    import matplotlib.patches as mpatches
    import numpy as np
    segments = [
        ('Physical',      OCRE_HEX),
        ('Mental',        BLUE),
        ('Social',        GREEN),
        ('Emotional',     '#9c27b0'),
        ('Nutritional',   '#00897b'),
        ('Environmental', '#e65100'),
    ]
    cx, cy = 5.0, 2.6
    r_outer, r_inner = 1.85, 0.62
    n = len(segments)
    for i, (label, color) in enumerate(segments):
        theta1 = 90 + i * (360 / n)
        theta2 = 90 + (i + 1) * (360 / n)
        wedge = mpatches.Wedge((cx, cy), r_outer, theta1, theta2,
                               width=r_outer - r_inner,
                               facecolor=color, edgecolor=BG, lw=2, alpha=0.88)
        ax.add_patch(wedge)
        mid_angle = np.radians((theta1 + theta2) / 2)
        label_r = r_inner + (r_outer - r_inner) * 0.55
        lx = cx + label_r * np.cos(mid_angle)
        ly = cy + label_r * np.sin(mid_angle)
        ax.text(lx, ly, label, ha='center', va='center',
                fontsize=9.5, fontweight='bold', color=WHITE, rotation=0)
    # Centre circle
    centre = mpatches.Circle((cx, cy), r_inner,
                              facecolor=OCRE_HEX, edgecolor=FG, lw=2)
    ax.add_patch(centre)
    ax.text(cx, cy, 'WELLNESS\nWHEEL', ha='center', va='center',
            fontsize=8.5, fontweight='bold', color=WHITE)
    ax.text(5.0, 4.68, 'Dimensions of Health & Well-being',
            color=LGRAY, fontsize=11, ha='center')

def _draw_atom(ax):
    import matplotlib.patches as mpatches
    import numpy as np
    cx, cy = 5.0, 2.7
    # Nucleus
    ax.plot(cx, cy, 'o', color=OCRE_HEX, ms=18, zorder=5)
    ax.text(cx, cy - 0.55, 'Nucleus', color=OCRE_HEX,
            fontsize=11, ha='center')
    # Orbits at different tilts
    for a, b, angle in [(1.8, 0.5, 0), (1.6, 0.45, 60), (2.0, 0.4, 120)]:
        theta = np.linspace(0, 2 * np.pi, 120)
        cos_a, sin_a = np.cos(np.radians(angle)), np.sin(np.radians(angle))
        xe = a * np.cos(theta)
        ye = b * np.sin(theta)
        xr = cx + xe * cos_a - ye * sin_a
        yr = cy + xe * sin_a + ye * cos_a
        ax.plot(xr, yr, color=FG, lw=1.2, linestyle='--', alpha=0.7)
        # Electron at a point on the orbit
        idx = int(len(theta) * 0.25)
        ax.plot(xr[idx], yr[idx], 'o', color=BLUE, ms=8, zorder=5)
    ax.text(cx, 4.55, 'Atomic Model', color=LGRAY,
            fontsize=12, ha='center')


def generate_chapter_image(chapter_name: str, out_path: Path) -> None:
    """Generate a 1840×920 chapter banner with a thematic scientific diagram."""
    import matplotlib.pyplot as plt

    fig, ax = _setup_fig()
    name_l = chapter_name.lower()

    if any(k in name_l for k in ('electric', 'magnetic', 'current', 'circuit', 'heating')):
        _draw_electricity(ax)
    elif any(k in name_l for k in ('light', 'mirror', 'lens', 'optic', 'reflect', 'refract')):
        _draw_light(ax)
    elif any(k in name_l for k in ('particle', 'particulate')):
        _draw_particles(ax)
    elif any(k in name_l for k in ('element', 'compound', 'mixture')):
        _draw_elements(ax)
    elif any(k in name_l for k in ('solution', 'solute', 'solvent', 'dissolve')):
        _draw_solution(ax)
    elif any(k in name_l for k in ('health', 'disease', 'body', 'treasure')):
        _draw_health(ax)
    else:
        _draw_atom(ax)

    plt.tight_layout(pad=0)
    plt.savefig(str(out_path), dpi=100, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    plt.close(fig)


def copy_template_assets(template_dir: Path, out_dir: Path) -> None:
    """Copy Legrand template assets into the output directory."""
    for asset in ('background.pdf', 'chapterhead1.pdf', 'bibliography.bib',
                  'StyleInd.ist', 'placeholder.jpg'):
        src = template_dir / asset
        if src.exists():
            shutil.copy2(src, out_dir / asset)

    # Copy structure.tex with patches:
    # 1. graphicspath → current directory
    # 2. Remove biblatex block (no citations in notes docs → suppresses BibTeX errors)
    structure = (template_dir / 'structure.tex').read_text(encoding='utf-8')
    structure = structure.replace(
        r'\graphicspath{{Pictures/}}',
        r'\graphicspath{{./}}'
    )
    structure = re.sub(
        r'\\usepackage\{csquotes\}.*?\\defbibheading\{bibempty\}\{\}',
        '',
        structure,
        flags=re.DOTALL
    )
    (out_dir / 'structure.tex').write_text(structure, encoding='utf-8')


def lyx_escape(text: str) -> str:
    """Escape backslashes for LyX plain layout."""
    return text.replace('\\', '\\backslash\n')

def render_lyx_segments(segments) -> str:
    """
    Render (text, bold) segments as LyX inline markup.
    Uses \\series bold / \\series default switches only where needed.
    """
    parts = []
    prev_bold = False
    for text, bold in segments:
        if bold and not prev_bold:
            parts.append('\\series bold\n')
        elif not bold and prev_bold:
            parts.append('\\series default\n')
        parts.append(lyx_escape(text))
        prev_bold = bold
    if prev_bold:
        parts.append('\n\\series default\n')
    return ''.join(parts)

def write_lyx(elements, out_path: Path, media_dir: Path, template_dir: Path = None):
    if template_dir is not None:
        lines = [load_template_prefix(template_dir)]
    else:
        lines = [LYX_HEADER]
    in_enum = False
    chapter_idx = 0

    def close_enum():
        nonlocal in_enum
        in_enum = False  # LyX Enumerate auto-closes at next non-Enumerate layout

    def ert(content: str) -> str:
        """Wrap content in an ERT inset."""
        return (
            '\\begin_inset ERT\nstatus open\n\n'
            '\\begin_layout Plain Layout\n'
            f'{content}\n'
            '\\end_layout\n\n'
            '\\end_inset\n\n'
        )

    for el in elements:
        t = el['type']

        if t == 'heading':
            close_enum()
            lvl = el['level']
            layout = {'chapter': 'Chapter', 'section': 'Section',
                      'subsection': 'Subsection',
                      'subsubsection': 'Subsubsection'}.get(lvl, 'Section')
            if lvl == 'chapter' and template_dir is not None:
                chapter_idx += 1
                img_name = f'chapterhead_Ch{chapter_idx}.png'
                generate_chapter_image(el['text'], out_path.parent / img_name)
                lines.append(chapterimage_block(img_name))
            lines.append(f'\\begin_layout {layout}\n')
            lines.append(f'{el["text"]}\n')
            lines.append('\\end_layout\n\n')

        elif t == 'body':
            close_enum()
            lines.append('\\begin_layout Standard\n')
            lines.append(render_lyx_segments(el['segments']))
            lines.append('\n\\end_layout\n\n')

        elif t == 'mcq':
            in_enum = True
            q = el['text']
            opts = el['options']
            lines.append('\\begin_layout Enumerate\n')
            q_ert = (
                f'\\backslash\ntextbf{{{q}}}\\backslash\n\\backslash\n[0.13cm]\n'
            )
            lines.append(ert(q_ert))
            tab = (
                '\\backslash\nbegin{tabular}{@{}p{0.45\\backslash\ntextwidth} '
                'p{0.45\\backslash\ntextwidth}@{}}\n'
                f'$\\backslash\nsquare$ A) {opts[0]} & '
                f'$\\backslash\nsquare$ B) {opts[1]} \\backslash\n\\backslash\n\n'
                f'$\\backslash\nsquare$ C) {opts[2]} & '
                f'$\\backslash\nsquare$ D) {opts[3]}\n'
                '\\backslash\nend{tabular}\n'
            )
            lines.append(ert(tab))
            lines.append('\\end_layout\n\n')

        elif t == 'image':
            close_enum()
            fname = el['filename']
            scale_pct = int(el['scale'] * 100)
            lines.append('\\begin_layout Standard\n')
            lines.append('\\align center\n')
            lines.append('\\begin_inset Graphics\n')
            lines.append(f'\tfilename media/{fname}\n')
            lines.append(f'\twidth {scale_pct}text%\n')
            lines.append('\\end_inset\n\n')
            lines.append('\\end_layout\n\n')

        elif t == 'list':
            close_enum()
            layout = 'Enumerate' if el.get('ordered') else 'Itemize'
            for d, segs in el['items']:
                lines.append('\\begin_deeper\n' * d)
                lines.append(f'\\begin_layout {layout}\n')
                lines.append(render_lyx_segments(segs))
                lines.append('\n\\end_layout\n\n')
                lines.append('\\end_deeper\n' * d)

        elif t == 'table':
            close_enum()
            rows = el['rows']
            ncol = max(len(r) for r in rows)
            tex = ['\\begin{tabular}{|' + 'l|' * ncol + '}', '\\hline']
            for row in rows:
                row = row + [''] * (ncol - len(row))
                tex.append(' & '.join(tex_escape(c) for c in row) + ' \\\\')
                tex.append('\\hline')
            tex.append('\\end{tabular}')
            lines.append('\\begin_layout Standard\n\\align center\n')
            lines.append(ert(lyx_escape('\n'.join(tex)) + '\n'))
            lines.append('\\end_layout\n\n')

    if template_dir is not None:
        lines.append('\\end_body\n\\end_document\n')
    else:
        lines.append(LYX_FOOTER)
    out_path.write_text(''.join(lines), encoding='utf-8')

# ---------------------------------------------------------------------------
# Image extraction
# ---------------------------------------------------------------------------
def extract_images(docx_path: Path, media_dir: Path, media_files: list):
    media_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path) as z:
        for f in media_files:
            fname = Path(f).name
            dest = media_dir / fname
            dest.write_bytes(z.read(f))

# ---------------------------------------------------------------------------
# Main conversion
# ---------------------------------------------------------------------------
def convert(docx_path: Path, base_dir: Path, logger: logging.Logger):
    stem = docx_path.stem
    out_dir = base_dir / 'output' / stem
    out_dir.mkdir(parents=True, exist_ok=True)
    media_dir = out_dir / 'media'

    logger.info(f'Converting: {docx_path.name}')

    elements, rid_map, media_files = parse_docx(docx_path)
    elements = reconstruct_paragraphs(elements)
    elements = group_mcq_blocks(elements)

    extract_images(docx_path, media_dir, media_files)
    logger.info(f'  Extracted {len(media_files)} images to {media_dir}')

    tex_path = out_dir / f'{stem}.tex'
    write_tex(elements, tex_path, media_dir)
    logger.info(f'  Wrote {tex_path}')

    lyx_path = out_dir / f'{stem}.lyx'
    t_dir = TEMPLATE_DIR if TEMPLATE_DIR.exists() else None
    write_lyx(elements, lyx_path, media_dir, template_dir=t_dir)
    logger.info(f'  Wrote {lyx_path}')

    if t_dir is not None:
        copy_template_assets(t_dir, out_dir)
        logger.info(f'  Copied template assets to {out_dir}')

    counts = {t: sum(1 for e in elements if e['type'] == t)
              for t in ('heading', 'body', 'mcq', 'image')}
    logger.info(f'  Elements: {counts}')
    return counts

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    base_dir = Path(__file__).parent
    logs_dir = base_dir / 'logs'
    logs_dir.mkdir(exist_ok=True)

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = logs_dir / f'convert_{ts}.log'

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout),
        ]
    )
    logger = logging.getLogger(__name__)

    # Determine files to process
    if len(sys.argv) > 1:
        files = []
        for arg in sys.argv[1:]:
            p = Path(arg)
            if not p.is_absolute():
                p = base_dir / p
            if not p.exists():
                logger.error(f'File not found: {p}')
            elif p.suffix.lower() != '.docx':
                logger.error(f'Not a .docx file: {p}')
            else:
                files.append(p)
    else:
        files = [f for f in base_dir.glob('*.docx')
                 if not f.name.startswith('~$')]

    if not files:
        logger.info('No .docx files found.')
        return

    logger.info(f'Processing {len(files)} file(s)')
    for f in files:
        try:
            convert(f, base_dir, logger)
        except Exception as e:
            logger.error(f'Error converting {f.name}: {e}', exc_info=True)

    logger.info('Done.')

if __name__ == '__main__':
    main()
