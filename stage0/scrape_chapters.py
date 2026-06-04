#!/usr/bin/env python3
"""v4.0 orchestrator: build a themed combined LyX/TeX book straight from the web.

Reads input/CHAPTER-LINKS, fetches+caches each chapter page, parses it to Stage 2
element dicts (stage2/html_source.parse_html), downloads the inline images, prepends a
chapter heading per URL, then reuses Stage 2's themed writers (write_tex / write_lyx
with the Final-lyx_template → chapter banners + theme) to emit one combined document.

No PDF/DOCX is used. Run stage0/validate_against_pdf.py separately to confirm the HTML
text is complete against the module PDF.

Usage:
    python stage0/scrape_chapters.py [<module-name>]
"""

import logging
import os
import sys
from pathlib import Path

import requests

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))            # .../stage0/
_COMBINED_DIR = os.path.dirname(_SCRIPT_DIR)                        # .../final-combined-pdf-to-lyx/
_STAGE2_DIR = os.path.join(_COMBINED_DIR, 'stage2')
_INPUT_DIR = os.path.join(_COMBINED_DIR, 'input')
_LINKS_FILE = os.path.join(_INPUT_DIR, 'CHAPTER-LINKS')
_SCRAPED_DIR = os.path.join(_INPUT_DIR, 'scraped')
_LOGS_DIR = os.path.join(_SCRIPT_DIR, 'logs')
os.makedirs(_LOGS_DIR, exist_ok=True)

sys.path.insert(0, _STAGE2_DIR)
sys.path.insert(0, _SCRIPT_DIR)
import html_source           # noqa: E402  (stage2/html_source.py)
import convert               # noqa: E402  (stage2/convert.py)
from scrape_images import parse_links  # noqa: E402  (reuse v3.0 CHAPTER-LINKS parser)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler(os.path.join(_LOGS_DIR, 'scrape_chapters.log')),
              logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

_UA = html_source._UA
_DEFAULT_MODULE = 'Science_Class8th_Module2_v4'


def download_images(elements, media_dir, session):
    """Download every image element's URL into media_dir/<filename> (cached)."""
    n = 0
    for el in elements:
        if el.get('type') != 'image':
            continue
        dest = os.path.join(media_dir, el['filename'])
        if os.path.exists(dest) and os.path.getsize(dest) > 0:
            n += 1
            continue
        try:
            r = session.get(el['url'], timeout=30)
            r.raise_for_status()
            with open(dest, 'wb') as fh:
                fh.write(r.content)
            n += 1
        except requests.RequestException as exc:
            logger.error(f"  image download failed {el['filename']}: {exc}")
    return n


def main():
    module = sys.argv[1] if len(sys.argv) > 1 else _DEFAULT_MODULE
    chapters = parse_links(_LINKS_FILE)
    if not chapters:
        logger.error("No chapters parsed from CHAPTER-LINKS.")
        sys.exit(1)

    out_dir = Path(_STAGE2_DIR) / 'output' / module
    media_dir = out_dir / 'media'
    media_dir.mkdir(parents=True, exist_ok=True)

    logger.info("=" * 70)
    logger.info(f"v4.0 web-only build: {len(chapters)} chapters → {out_dir}")
    logger.info("=" * 70)

    session = requests.Session()
    session.headers.update({'User-Agent': _UA})

    combined = []
    summary = []
    for ch in chapters:
        cache = os.path.join(_SCRAPED_DIR, f"Chapter{ch['num']}", 'page.html')
        logger.info(f"Chapter {ch['num']} ({ch['name']}): {ch['url']}")
        html = html_source.fetch(ch['url'], cache_path=cache)
        els = html_source.parse_html(html)

        combined.append({'type': 'heading', 'level': 'chapter', 'text': ch['name']})
        combined.extend(els)

        n_img = download_images(els, str(media_dir), session)
        counts = {}
        for e in els:
            counts[e['type']] = counts.get(e['type'], 0) + 1
        items = sum(len(e['items']) for e in els if e['type'] == 'list')
        summary.append((ch['num'], ch['name'], counts, items, n_img))
        logger.info(f"  elements: {counts} | list items: {items} | images: {n_img}")

    template_dir = convert.TEMPLATE_DIR if convert.TEMPLATE_DIR.exists() else None
    convert.write_tex(combined, out_dir / f'{module}.tex', media_dir)
    convert.write_lyx(combined, out_dir / f'{module}.lyx', media_dir, template_dir=template_dir)
    if template_dir is not None:
        convert.copy_template_assets(template_dir, out_dir)

    logger.info("=" * 70)
    for num, name, counts, items, n_img in summary:
        logger.info(f"  Ch{num:<3} {name[:40]:<40} "
                    f"head={counts.get('heading',0):>2} body={counts.get('body',0):>3} "
                    f"lists={counts.get('list',0):>2}/{items:<3} imgs={n_img:>2} "
                    f"tables={counts.get('table',0)}")
    logger.info("=" * 70)
    logger.info(f"✓ Wrote {module}.tex and {module}.lyx (+ media/, theme={'yes' if template_dir else 'no'})")
    logger.info(f"  Output: {out_dir}")
    logger.info("  Open the .lyx in LyX, or pdflatex the .tex (needs rupee.sty).")


if __name__ == '__main__':
    main()
