#!/usr/bin/env python3
"""PROTOTYPE: build .tex + .lyx for one EduRev chapter straight from its web page.

Evaluates a web-only pipeline (no PDF/DOCX): parses the chapter HTML into ordered
content — headings, paragraphs (with bold), bullet/numbered lists (incl. nesting),
inline images, tables — downloads the images, and emits standalone .tex and .lyx
using Stage 2's formatting conventions, plus itemize support (which Stage 2 lacks).

Outputs to stage0/prototype_out/:  preview.md, chapter.tex, chapter.lyx, media/

Usage:
    python stage0/prototype_html_to_content.py [<chapter-url>]
"""

import os
import re
import sys

import requests
from bs4 import BeautifulSoup, Tag

_UA = ('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 '
       '(KHTML, like Gecko) Chrome/120.0 Safari/537.36')
_DEFAULT_URL = 'https://edurev.in/t/424649/Chapter-Notes-Health-The-Ultimate-Treasure/'

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_OUT_DIR = os.path.join(_SCRIPT_DIR, 'prototype_out')
_MEDIA_DIR = os.path.join(_OUT_DIR, 'media')
os.makedirs(_MEDIA_DIR, exist_ok=True)

# Typographic Unicode → ASCII/LaTeX-safe (pdflatex utf8 chokes on thin/zero-width spaces)
_UNICODE_MAP = {
    ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ',
    '​': '', '‌': '', '‍': '', '﻿': '', '­': '',
    '‘': "'", '’': "'", '“': '"', '”': '"',
    '–': '--', '—': '---', '…': '...', '•': '',
}
_UNICODE_RE = re.compile('|'.join(map(re.escape, _UNICODE_MAP)))


def clean_text(s):
    return _UNICODE_RE.sub(lambda m: _UNICODE_MAP[m.group()], s)


_HEADINGS = {'h2': 1, 'h3': 2, 'h4': 3}   # → section / subsection / subsubsection
_BLOCK = set(_HEADINGS) | {'p', 'ul', 'ol', 'table', 'img'}
_SCALES = [0.25, 0.4, 0.5, 0.6, 0.75]
_NOISE_RE = re.compile(r'^(view more|view solution|join for free|table of contents'
                       r'|explore courses|download)\b', re.IGNORECASE)


# --------------------------------------------------------------------------- #
# Fetch + locate content
# --------------------------------------------------------------------------- #
def fetch(url):
    cache = os.path.join(_OUT_DIR, 'page.html')
    if os.path.exists(cache):
        with open(cache, encoding='utf-8') as fh:
            return fh.read()
    html = requests.get(url, headers={'User-Agent': _UA}, timeout=30).text
    with open(cache, 'w', encoding='utf-8') as fh:
        fh.write(html)
    return html


def find_content_root(soup):
    """Smallest ancestor containing (almost) all _lg.jpg content images."""
    imgs = soup.find_all('img', src=lambda s: s and '_lg.jpg' in s)
    if not imgs:
        return soup.body or soup
    cur, best = imgs[0], imgs[0]
    for _ in range(12):
        par = cur.find_parent()
        if par is None:
            break
        best = par
        if len(par.find_all('img', src=lambda s: s and '_lg.jpg' in s)) >= max(2, int(0.8 * len(imgs))):
            break
        cur = par
    return best


# --------------------------------------------------------------------------- #
# Inline / list / table helpers
# --------------------------------------------------------------------------- #
def runs_with_bold(el):
    """[(text, is_bold)] for inline content; keeps inter-node spacing."""
    runs = []
    for node in el.descendants:
        if isinstance(node, Tag):
            continue
        text = clean_text(str(node))
        if not text.strip():
            continue
        bold = any(p.name in ('strong', 'b') for p in node.parents)
        runs.append((re.sub(r'\s+', ' ', text), bold))
    return runs


def runs_text(runs):
    return re.sub(r'\s+', ' ', ''.join(t for t, _ in runs)).strip()


def list_items(list_el, depth=0):
    """Yield (depth, runs) per <li>, recursing into nested lists."""
    for li in list_el.find_all('li', recursive=False):
        nested = li.find_all(['ul', 'ol'], recursive=False)
        runs = []
        for node in li.descendants:
            if isinstance(node, Tag):
                continue
            if any(anc.name in ('ul', 'ol') for anc in node.parents if anc in li.descendants):
                continue
            t = clean_text(str(node))
            if t.strip():
                bold = any(p.name in ('strong', 'b') for p in node.parents)
                runs.append((re.sub(r'\s+', ' ', t), bold))
        if runs:
            yield depth, runs
        for sub in nested:
            yield from list_items(sub, depth + 1)


def table_rows(table):
    rows = []
    for tr in table.find_all('tr'):
        cells = [clean_text(re.sub(r'\s+', ' ', c.get_text(' ', strip=True))) for c in tr.find_all(['td', 'th'])]
        if cells:
            rows.append(cells)
    return rows


def pick_scale(width_px):
    if not width_px:
        return 0.6
    frac = width_px / 700.0   # EduRev content column ≈ 700px
    return min(_SCALES, key=lambda s: abs(s - frac))


def img_width_px(el):
    m = re.search(r'width:\s*([0-9.]+)px', el.get('style', '') or '')
    if m:
        return float(m.group(1))
    w = el.get('width')
    return float(w) if w and w.isdigit() else None


# --------------------------------------------------------------------------- #
# Extract ordered elements
# --------------------------------------------------------------------------- #
def extract(root):
    elements, consumed, img_seen = [], [], set()

    def inside_consumed(el):
        return any(c in el.parents for c in consumed)

    for el in root.descendants:
        name = getattr(el, 'name', None)
        if name not in _BLOCK:
            continue
        if name == 'img':
            src = el.get('src', '')
            if '_lg.jpg' in src and src not in img_seen:
                img_seen.add(src)
                url = 'https:' + src if src.startswith('//') else src
                elements.append(('image', src.rsplit('/', 1)[-1], url, pick_scale(img_width_px(el))))
            continue
        if inside_consumed(el):
            continue
        if name in _HEADINGS:
            txt = clean_text(el.get_text(' ', strip=True))
            if txt and not _NOISE_RE.match(txt):
                elements.append(('heading', _HEADINGS[name], txt))
        elif name == 'p':
            runs = runs_with_bold(el)
            if runs_text(runs) and not _NOISE_RE.match(runs_text(runs)):
                elements.append(('para', runs))
        elif name in ('ul', 'ol'):
            items = [(d, r) for d, r in list_items(el) if not _NOISE_RE.match(runs_text(r))]
            if items:
                elements.append(('list', name, items))
            consumed.append(el)
        elif name == 'table':
            rows = table_rows(el)
            flat = ' '.join(' '.join(r) for r in rows).lower()
            if rows and 'table of contents' not in flat:   # skip auto-TOC
                elements.append(('table', rows))
            consumed.append(el)
    return elements


def download_images(elements):
    sess = requests.Session()
    sess.headers.update({'User-Agent': _UA})
    for e in elements:
        if e[0] != 'image':
            continue
        fname, url = e[1], e[2]
        dest = os.path.join(_MEDIA_DIR, fname)
        if os.path.exists(dest) and os.path.getsize(dest) > 0:
            continue
        try:
            r = sess.get(url, timeout=30)
            r.raise_for_status()
            with open(dest, 'wb') as fh:
                fh.write(r.content)
        except requests.RequestException as exc:
            print(f"  ! image download failed {fname}: {exc}")


# --------------------------------------------------------------------------- #
# Markdown preview
# --------------------------------------------------------------------------- #
def to_markdown(elements):
    out = []
    for e in elements:
        kind = e[0]
        if kind == 'heading':
            out.append('#' * (e[1] + 1) + ' ' + e[2])
        elif kind == 'para':
            out.append(runs_md(e[1]))
        elif kind == 'list':
            for depth, runs in e[2]:
                out.append('  ' * depth + '- ' + runs_md(runs))
        elif kind == 'image':
            out.append(f'![{e[1]}](media/{e[1]})')
        elif kind == 'table':
            for i, row in enumerate(e[1]):
                out.append('| ' + ' | '.join(row) + ' |')
                if i == 0:
                    out.append('| ' + ' | '.join(['---'] * len(row)) + ' |')
        out.append('')
    return '\n'.join(out)


def runs_md(runs):
    parts = [f'**{t.strip()}**' if b and t.strip() else t for t, b in runs]
    return re.sub(r'\s+', ' ', ''.join(parts)).strip()


# --------------------------------------------------------------------------- #
# LaTeX writer (mirrors stage2 TEX_PREAMBLE + adds itemize)
# --------------------------------------------------------------------------- #
TEX_PREAMBLE = r"""\documentclass[12pt,a4paper]{book}
\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage{graphicx}
\usepackage{amssymb}
\usepackage{enumitem}
\IfFileExists{rupee.sty}{\usepackage{rupee}}{\providecommand{\rupee}{Rs.\,}}
\usepackage{geometry}
\geometry{margin=2.5cm}
\setlength{\parindent}{0pt}
\setlength{\parskip}{4pt}

\begin{document}
"""


def tex_escape(text):
    text = text.replace('\\', r'\textbackslash{}')
    for old, new in [('&', r'\&'), ('%', r'\%'), ('$', r'\$'), ('#', r'\#'),
                     ('{', r'\{'), ('}', r'\}'), ('~', r'\textasciitilde{}'),
                     ('^', r'\^{}'), ('_', r'\_'), ('₹', r'\rupee~')]:
        text = text.replace(old, new)
    return text


def render_tex(runs):
    return ''.join(f'\\textbf{{{tex_escape(t)}}}' if b else tex_escape(t) for t, b in runs)


def write_tex(elements, path):
    L = [TEX_PREAMBLE]
    list_depth = 0

    def close_lists(to=0):
        nonlocal list_depth
        while list_depth > to:
            L.append('\\end{itemize}\n')
            list_depth -= 1

    for e in elements:
        k = e[0]
        if k == 'heading':
            close_lists()
            cmd = {1: 'section', 2: 'subsection', 3: 'subsubsection'}[e[1]]
            L.append(f'\\{cmd}*{{{tex_escape(e[2])}}}\n')
        elif k == 'para':
            close_lists()
            L.append(f'\\par {render_tex(e[1])}\n\n')
        elif k == 'list':
            for depth, runs in e[2]:
                want = depth + 1
                while list_depth < want:
                    L.append('\\begin{itemize}\n')
                    list_depth += 1
                close_lists(want)
                L.append(f'\\item {render_tex(runs)}\n')
            close_lists()
        elif k == 'image':
            close_lists()
            L.append('\\begin{figure}[h]\n\\centering\n')
            L.append(f'\\includegraphics[width={e[3]}\\textwidth]{{media/{e[1]}}}\n')
            L.append('\\end{figure}\n\n')
        elif k == 'table':
            close_lists()
            ncol = max(len(r) for r in e[1])
            L.append('\\begin{tabular}{|' + 'l|' * ncol + '}\n\\hline\n')
            for row in e[1]:
                row = row + [''] * (ncol - len(row))
                L.append(' & '.join(tex_escape(c) for c in row) + ' \\\\\n\\hline\n')
            L.append('\\end{tabular}\n\n')
    close_lists()
    L.append('\\end{document}\n')
    with open(path, 'w', encoding='utf-8') as fh:
        fh.write(''.join(L))


# --------------------------------------------------------------------------- #
# LyX writer (standalone book; mirrors stage2 LYX_HEADER + adds Itemize)
# --------------------------------------------------------------------------- #
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

LYX_FOOTER = "\\end_body\n\\end_document\n"


def lyx_escape(text):
    return text.replace('\\', '\\backslash\n')


def render_lyx(runs):
    parts, prev = [], False
    for text, bold in runs:
        if bold and not prev:
            parts.append('\n\\series bold\n')
        elif not bold and prev:
            parts.append('\n\\series default\n')
        parts.append(lyx_escape(text))
        prev = bold
    if prev:
        parts.append('\n\\series default\n')
    return ''.join(parts)


def write_lyx(elements, path):
    L = [LYX_HEADER]
    for e in elements:
        k = e[0]
        if k == 'heading':
            layout = {1: 'Section', 2: 'Subsection', 3: 'Subsubsection'}[e[1]]
            L.append(f'\\begin_layout {layout}\n{e[2]}\n\\end_layout\n\n')
        elif k == 'para':
            L.append('\\begin_layout Standard\n' + render_lyx(e[1]) + '\n\\end_layout\n\n')
        elif k == 'list':
            for depth, runs in e[2]:
                # LyX nests Itemize by depth via \begin_deeper
                L.append('\\begin_deeper\n' * depth)
                L.append('\\begin_layout Itemize\n' + render_lyx(runs) + '\n\\end_layout\n\n')
                L.append('\\end_deeper\n' * depth)
        elif k == 'image':
            L.append('\\begin_layout Standard\n\\align center\n')
            L.append('\\begin_inset Graphics\n')
            L.append(f'\tfilename media/{e[1]}\n')
            L.append(f'\twidth {int(e[3] * 100)}text%\n')
            L.append('\\end_inset\n\n\\end_layout\n\n')
        elif k == 'table':
            # Prototype: render table as ERT tabular for fidelity preview
            ncol = max(len(r) for r in e[1])
            body = '\\backslash\nbegin{tabular}{|' + 'l|' * ncol + '}\n'
            for row in e[1]:
                row = row + [''] * (ncol - len(row))
                body += ' & '.join(row) + ' \\backslash\n\\backslash\n\n'
            body += '\\backslash\nend{tabular}\n'
            L.append('\\begin_layout Standard\n\\begin_inset ERT\nstatus open\n\n'
                     '\\begin_layout Plain Layout\n' + body + '\\end_layout\n\n'
                     '\\end_inset\n\n\\end_layout\n\n')
    L.append(LYX_FOOTER)
    with open(path, 'w', encoding='utf-8') as fh:
        fh.write(''.join(L))


# --------------------------------------------------------------------------- #
def main():
    url = sys.argv[1] if len(sys.argv) > 1 else _DEFAULT_URL
    soup = BeautifulSoup(fetch(url), 'html5lib')
    elements = extract(find_content_root(soup))

    counts = {}
    for e in elements:
        counts[e[0]] = counts.get(e[0], 0) + 1
    n_items = sum(len(e[2]) for e in elements if e[0] == 'list')

    download_images(elements)
    with open(os.path.join(_OUT_DIR, 'preview.md'), 'w', encoding='utf-8') as fh:
        fh.write(f'<!-- source: {url} -->\n\n{to_markdown(elements)}\n')
    write_tex(elements, os.path.join(_OUT_DIR, 'chapter.tex'))
    write_lyx(elements, os.path.join(_OUT_DIR, 'chapter.lyx'))

    print('Element counts:', counts, '| list items:', n_items)
    print('Wrote:', os.path.join(_OUT_DIR, 'chapter.tex'))
    print('       ', os.path.join(_OUT_DIR, 'chapter.lyx'))
    print('       ', os.path.join(_OUT_DIR, 'preview.md'))
    print('Images in:', _MEDIA_DIR)


if __name__ == '__main__':
    main()
