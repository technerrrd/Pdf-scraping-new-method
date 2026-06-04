#!/usr/bin/env python3
"""v4.0 validation oracle: confirm the scraped HTML is complete against the module PDF.

Read-only QA. The HTML is authoritative for the output; this script never edits or
overrides it. It only measures, per chapter, what fraction of the PDF's content lines
are present in that chapter's scraped HTML, and flags any chapter below a threshold for
manual review. The PDF is used purely as a completeness oracle — to catch a failed
scrape, gated content, or a future site change — not as a text source.

Skips automatically when no module PDF is present in input/.

Usage:
    python stage0/validate_against_pdf.py
"""

import glob
import html as _html
import os
import re
import sys

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_COMBINED_DIR = os.path.dirname(_SCRIPT_DIR)
_INPUT_DIR = os.path.join(_COMBINED_DIR, 'input')
_SCRAPED_DIR = os.path.join(_INPUT_DIR, 'scraped')
_LINKS_FILE = os.path.join(_INPUT_DIR, 'CHAPTER-LINKS')

sys.path.insert(0, _SCRIPT_DIR)
sys.path.insert(0, os.path.join(_COMBINED_DIR, 'stage1'))
from scrape_images import parse_links          # noqa: E402
from analyze_document import analyze_pdf        # noqa: E402

import pymupdf as fitz                          # noqa: E402

_THRESHOLD = 0.95
# PDF-only artifacts that never appear in the HTML (don't count them as "missing")
_ARTIFACT_RE = re.compile(
    r'edurev|durev|^\d+\s*/\s*\d+$|^\d{2}/\d{2}/\d{2}|view (solution|more)'
    r'|^https?://|chapter notes \| ', re.IGNORECASE)


def norm(s):
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9 ]+', ' ', s.lower())).strip()


def html_text(path):
    with open(path, encoding='utf-8') as fh:
        h = fh.read()
    h = re.sub(r'<(script|style)\b[^>]*>.*?</\1>', ' ', h, flags=re.I | re.S)
    h = _html.unescape(re.sub(r'<[^>]+>', ' ', h))
    return norm(h)


def pdf_chapter_lines(pdf_doc, start_page, end_page):
    """Normalized PDF content lines (1-based inclusive range), artifacts removed."""
    lines = []
    for pno in range(start_page - 1, min(end_page, len(pdf_doc))):
        for b in pdf_doc[pno].get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = "".join(s["text"] for s in l.get("spans", [])).strip()
                if len(t) >= 12 and not _ARTIFACT_RE.search(t):
                    lines.append(norm(t))
    return [x for x in lines if len(x) >= 8]


def main():
    pdfs = glob.glob(os.path.join(_INPUT_DIR, '*.pdf'))
    if not pdfs:
        print("No module PDF in input/ — skipping validation (HTML-only run).")
        return
    pdf_path = pdfs[0]

    chapters = parse_links(_LINKS_FILE)
    markers = analyze_pdf(pdf_path)
    pdf_doc = fitz.open(pdf_path)
    total_pages = len(pdf_doc)

    if len(markers) != len(chapters):
        print(f"⚠️  PDF chapter markers ({len(markers)}) != CHAPTER-LINKS ({len(chapters)}); "
              f"mapping by document order anyway.")

    print("=" * 78)
    print(f"PDF validation oracle — {os.path.basename(pdf_path)}")
    print("=" * 78)

    flagged = []
    for i, ch in enumerate(chapters):
        start = markers[i]['page'] if i < len(markers) else 1
        end = (markers[i + 1]['page'] - 1) if (i + 1) < len(markers) else total_pages
        cache = os.path.join(_SCRAPED_DIR, f"Chapter{ch['num']}", 'page.html')
        if not os.path.exists(cache):
            print(f"  Ch{ch['num']:<3} {ch['name'][:34]:<34}  no cached HTML — run scrape_chapters first")
            flagged.append(ch['num'])
            continue

        pdf_lines = pdf_chapter_lines(pdf_doc, start, end)
        htext = html_text(cache)
        present = sum(1 for t in pdf_lines if t[:40] in htext)
        cov = present / len(pdf_lines) if pdf_lines else 1.0
        flag = '  ⚠️ REVIEW' if cov < _THRESHOLD else ''
        if cov < _THRESHOLD:
            flagged.append(ch['num'])
        print(f"  Ch{ch['num']:<3} {ch['name'][:34]:<34}  pages {start:>3}-{end:<3}  "
              f"coverage {present:>3}/{len(pdf_lines):<3} = {cov*100:5.1f}%{flag}")

    pdf_doc.close()
    print("=" * 78)
    if flagged:
        print(f"⚠️  {len(flagged)} chapter(s) below {int(_THRESHOLD*100)}% — review: {flagged}")
        sys.exit(2)
    print(f"✓ All chapters ≥ {int(_THRESHOLD*100)}% — HTML is complete against the PDF.")


if __name__ == '__main__':
    main()
