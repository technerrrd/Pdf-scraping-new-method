#!/usr/bin/env python3
"""v4.0 HTML front-end: parse an EduRev chapter page into Stage 2 element dicts.

The scraped EduRev page is the original source the PDF/DOCX were derived from — it is
server-rendered, complete, and carries real semantics (h2/h3/h4 headings, <strong>
bold, <ul><li> lists incl. nesting, inline _lg.jpg images in reading order, <table>s).
This module reconstructs that into the SAME element schema Stage 2's writers consume,
plus a new 'list' element type, so `write_tex` / `write_lyx` (template + banners) can
be reused unchanged apart from list support.

Element schema produced (matches stage2/convert.py):
  {'type':'heading','level':'section'|'subsection'|'subsubsection','text':str}
  {'type':'body','text':str,'segments':[(text,bold)..]}
  {'type':'list','ordered':bool,'items':[(depth,segments)..]}      # NEW
  {'type':'image','filename':str,'url':str,'scale':float}
  {'type':'table','rows':[[cell,..],..]}
"""

import os
import re

import requests
from bs4 import BeautifulSoup, Tag

_UA = ('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 '
       '(KHTML, like Gecko) Chrome/120.0 Safari/537.36')

_HEADINGS = {'h2': 'section', 'h3': 'subsection', 'h4': 'subsubsection'}
_BLOCK = set(_HEADINGS) | {'p', 'ul', 'ol', 'table', 'img'}
_SCALES = [0.25, 0.4, 0.5, 0.6, 0.75]
_NOISE_RE = re.compile(r'^(view more|view solution|join for free|table of contents'
                       r'|explore courses|download|attempt test)\b', re.IGNORECASE)
# Leading numbering prefixes to strip from headings (LaTeX auto-numbers) — mirrors
# stage2/convert.py HEADING_NUM_RE: '1.' / '1.1' / 'Q1' / '(a)'
_HEADING_NUM_RE = re.compile(r'^(?:Q\d+\.?\s+|\d+(?:\.\d+)*\.?\s+|\([a-zA-Z0-9]+\)\s+)')

# Typographic Unicode → ASCII/LaTeX-safe (pdflatex utf8 chokes on thin/zero-width spaces)
_UNICODE_MAP = {
    ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ',
    '​': '', '‌': '', '‍': '', '﻿': '', '­': '',
    '‘': "'", '’': "'", '“': '"', '”': '"',
    '–': '--', '—': '---', '…': '...', '•': '',
}
# sub/superscripts and math/science symbols → ASCII (plain text; no $ _ ^ \ {} which
# tex_escape would re-escape). Faithful math typesetting is a known follow-up.
_UNICODE_MAP.update({
    '₀': '0', '₁': '1', '₂': '2', '₃': '3', '₄': '4',
    '₅': '5', '₆': '6', '₇': '7', '₈': '8', '₉': '9',
    '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
    '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9',
    '→': ' -> ', '←': ' <- ', '↔': ' <-> ', '⇌': ' <=> ', '⟶': ' -> ',
    '×': 'x', '÷': '/', '±': '+/-', '≈': '~', '≤': '<=', '≥': '>=', '≠': '!=',
    '°': ' deg', '·': '.', '′': "'", '″': '"', '∴': 'therefore', '∞': 'infinity',
})
_UNICODE_RE = re.compile('|'.join(map(re.escape, _UNICODE_MAP)))


def clean_text(s):
    return _UNICODE_RE.sub(lambda m: _UNICODE_MAP[m.group()], s)


# --------------------------------------------------------------------------- #
# Fetch (cache) + locate content
# --------------------------------------------------------------------------- #
def fetch(url, cache_path=None):
    """Fetch a page, caching to cache_path so converts are reproducible offline."""
    if cache_path and os.path.exists(cache_path):
        with open(cache_path, encoding='utf-8') as fh:
            return fh.read()
    html = requests.get(url, headers={'User-Agent': _UA}, timeout=30).text
    if cache_path:
        os.makedirs(os.path.dirname(cache_path), exist_ok=True)
        with open(cache_path, 'w', encoding='utf-8') as fh:
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
# Inline runs / lists / tables
# --------------------------------------------------------------------------- #
def _segments(el):
    """[(text, is_bold)] for inline content, preserving inter-element spacing.

    Whitespace-only text nodes between inline tags are collapsed to a single trailing
    space on the previous run (rather than dropped) — this fixes joins like
    'active</strong>and' that lost the source space.
    """
    runs = []
    for node in el.descendants:
        if isinstance(node, Tag):
            continue
        raw = str(node)
        if not raw.strip():
            if runs and not runs[-1][0].endswith(' '):
                runs[-1] = (runs[-1][0] + ' ', runs[-1][1])
            continue
        text = re.sub(r'\s+', ' ', clean_text(raw))
        bold = any(p.name in ('strong', 'b') for p in node.parents)
        runs.append((text, bold))
    return _merge_runs(runs)


def _merge_runs(runs):
    """Coalesce adjacent runs of the same weight."""
    out = []
    for text, bold in runs:
        if out and out[-1][1] == bold:
            out[-1] = (out[-1][0] + text, bold)
        else:
            out.append((text, bold))
    return out


def _segments_text(segments):
    return re.sub(r'\s+', ' ', ''.join(t for t, _ in segments)).strip()


def list_items(list_el, depth=0):
    """Yield (depth, segments) per <li>, recursing into nested lists (keeps depth)."""
    for li in list_el.find_all('li', recursive=False):
        nested = li.find_all(['ul', 'ol'], recursive=False)
        runs = []
        for node in li.descendants:
            if isinstance(node, Tag):
                continue
            if any(anc.name in ('ul', 'ol') for anc in node.parents if anc in li.descendants):
                continue
            raw = str(node)
            if not raw.strip():
                if runs and not runs[-1][0].endswith(' '):
                    runs[-1] = (runs[-1][0] + ' ', runs[-1][1])
                continue
            text = re.sub(r'\s+', ' ', clean_text(raw))
            bold = any(p.name in ('strong', 'b') for p in node.parents)
            runs.append((text, bold))
        runs = _merge_runs(runs)
        if _segments_text(runs):
            yield depth, runs
        for sub in nested:
            yield from list_items(sub, depth + 1)


def table_rows(table):
    rows = []
    for tr in table.find_all('tr'):
        cells = [clean_text(re.sub(r'\s+', ' ', c.get_text(' ', strip=True)))
                 for c in tr.find_all(['td', 'th'])]
        if cells:
            rows.append(cells)
    return rows


def pick_scale(width_px):
    if not width_px:
        return 0.6
    return min(_SCALES, key=lambda s: abs(s - width_px / 700.0))


def _img_width_px(el):
    m = re.search(r'width:\s*([0-9.]+)px', el.get('style', '') or '')
    if m:
        return float(m.group(1))
    w = el.get('width')
    return float(w) if w and str(w).isdigit() else None


# --------------------------------------------------------------------------- #
# Parse → elements
# --------------------------------------------------------------------------- #
def parse_html(html):
    """Return Stage 2 element dicts (in reading order) from chapter page HTML."""
    soup = BeautifulSoup(html, 'html5lib')   # HTML5 implied end-tags → clean sibling tree
    root = find_content_root(soup)
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
                elements.append({'type': 'image', 'filename': src.rsplit('/', 1)[-1],
                                 'url': url, 'scale': pick_scale(_img_width_px(el))})
            continue
        if inside_consumed(el):
            continue
        if name in _HEADINGS:
            txt = _HEADING_NUM_RE.sub('', clean_text(el.get_text(' ', strip=True))).strip()
            if txt and not _NOISE_RE.match(txt):
                elements.append({'type': 'heading', 'level': _HEADINGS[name], 'text': txt})
        elif name == 'p':
            segs = _segments(el)
            txt = _segments_text(segs)
            if txt and not _NOISE_RE.match(txt):
                elements.append({'type': 'body', 'text': txt, 'segments': segs})
        elif name in ('ul', 'ol'):
            items = [(d, s) for d, s in list_items(el) if not _NOISE_RE.match(_segments_text(s))]
            if items:
                elements.append({'type': 'list', 'ordered': name == 'ol', 'items': items})
            consumed.append(el)
        elif name == 'table':
            rows = table_rows(el)
            flat = ' '.join(' '.join(r) for r in rows).lower()
            if rows and 'table of contents' not in flat:   # skip auto-generated TOC
                elements.append({'type': 'table', 'rows': rows})
            consumed.append(el)
    return elements
