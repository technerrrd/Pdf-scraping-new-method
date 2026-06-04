#!/usr/bin/env python3
"""Stage 0: scrape full-size chapter images from EduRev source pages.

The Stage 1 PDF is an EduRev page printed from Firefox, so images that straddle
a page break get split in half. The original web pages host every content image
full-size on a CDN, in reading order, in server-rendered HTML. This script reads
input/CHAPTER-LINKS, fetches each chapter URL, extracts the content images, and
caches them per chapter so Stage 1 can use them instead of the DOCX-zip images.

Content images are the ONLY refs matching:
    cn.edurev.in/ApplicationImages/Temp/<UUID>_lg.jpg
All page chrome (icons/ads/course thumbnails/favicons) lives on other paths, so a
single path filter cleanly separates content from junk.

Output layout (consumed by Stage 1):
    input/scraped/Chapter<N>/image1.jpg, image2.jpg, ...
    input/scraped/manifest.json

Run, inspect the per-chapter counts and the cached images, prune anything stray,
then run Stage 1.
"""

import json
import logging
import os
import re
import sys
import time

import requests

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))        # .../stage0/
_COMBINED_DIR = os.path.dirname(_SCRIPT_DIR)                     # .../final-combined-pdf-to-lyx/
_INPUT_DIR = os.path.join(_COMBINED_DIR, 'input')               # .../input/
_LINKS_FILE = os.path.join(_INPUT_DIR, 'CHAPTER-LINKS')
_SCRAPED_DIR = os.path.join(_INPUT_DIR, 'scraped')              # .../input/scraped/
_LOGS_DIR = os.path.join(_SCRIPT_DIR, 'logs')
os.makedirs(_SCRAPED_DIR, exist_ok=True)
os.makedirs(_LOGS_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(_LOGS_DIR, 'scrape.log')),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

_UA = ('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 '
       '(KHTML, like Gecko) Chrome/120.0 Safari/537.36')
_REQUEST_DELAY = 1.5  # polite delay between page fetches (seconds)

# Content image on the EduRev CDN: ApplicationImages/Temp/<UUID>_lg.jpg
_CONTENT_IMG_RE = re.compile(r'ApplicationImages/Temp/[a-f0-9-]+_lg\.jpg', re.IGNORECASE)
# Chapter number from a label like "- Chapter3" / "- Chapter 4" / "-Chapter9"
_CHAPTER_NUM_RE = re.compile(r'chapter\s*(\d+)', re.IGNORECASE)


def parse_links(path):
    """Parse CHAPTER-LINKS into [{num, name, url}], in document order.

    Each line: <idx>\t<url> - Chapter<N>
    """
    if not os.path.exists(path):
        logger.error(f"CHAPTER-LINKS not found: {path}")
        return []

    chapters = []
    with open(path, encoding='utf-8') as fh:
        for raw in fh:
            line = raw.strip()
            if not line:
                continue
            url_match = re.search(r'https?://\S+', line)
            if not url_match:
                continue
            url = url_match.group(0).rstrip(',')
            # Strip the trailing " - Chapter N" label off the URL if glued on
            url = re.split(r'\s+-\s*[Cc]hapter', url)[0].rstrip()

            num_match = _CHAPTER_NUM_RE.search(line[url_match.end():]) or _CHAPTER_NUM_RE.search(line)
            num = int(num_match.group(1)) if num_match else len(chapters) + 1

            # Name slug from URL path: .../t/<id>/Chapter-Notes-<Name>/...
            slug_match = re.search(r'/t/\d+/([^/#?]+)', url)
            slug = slug_match.group(1) if slug_match else f'Chapter{num}'
            name = re.sub(r'^Chapter-Notes-', '', slug, flags=re.IGNORECASE)
            name = name.replace('-', ' ').strip()

            chapters.append({'num': num, 'name': name, 'url': url})
    return chapters


def extract_image_urls(html):
    """Return ordered, de-duplicated full content-image URLs from page HTML.

    The Temp/<UUID>_lg.jpg path uniquely identifies uploaded content images; all
    page chrome lives on other CDN paths, so the path filter alone is sufficient.
    (Note: the visible content container appears LATE in the markup, after the
    images, so it is not usable as an upper bound — do not scope by it.)
    """
    seen = set()
    urls = []
    for m in _CONTENT_IMG_RE.finditer(html):
        path = m.group(0)
        if path in seen:
            continue
        seen.add(path)
        urls.append(f'https://cn.edurev.in/{path}')
    return urls


def scrape_chapter(session, ch):
    """Fetch one chapter page and download its content images. Returns file list."""
    out_dir = os.path.join(_SCRAPED_DIR, f"Chapter{ch['num']}")
    os.makedirs(out_dir, exist_ok=True)

    logger.info(f"Chapter {ch['num']} ({ch['name']}): fetching {ch['url']}")
    try:
        resp = session.get(ch['url'], timeout=30)
        resp.raise_for_status()
    except requests.RequestException as exc:
        logger.error(f"  ✗ fetch failed: {exc}")
        return []

    img_urls = extract_image_urls(resp.text)
    logger.info(f"  found {len(img_urls)} content images")

    files = []
    for i, img_url in enumerate(img_urls, start=1):
        dest = os.path.join(out_dir, f'image{i}.jpg')
        if os.path.exists(dest) and os.path.getsize(dest) > 0:
            logger.debug(f"  cached image{i}.jpg")
            files.append(dest)
            continue
        try:
            r = session.get(img_url, timeout=30)
            r.raise_for_status()
            with open(dest, 'wb') as fh:
                fh.write(r.content)
            files.append(dest)
            logger.info(f"  ↓ image{i}.jpg ({len(r.content)} bytes)")
        except requests.RequestException as exc:
            logger.error(f"  ✗ image{i} download failed ({img_url}): {exc}")
        time.sleep(0.3)
    return files


def main():
    chapters = parse_links(_LINKS_FILE)
    if not chapters:
        logger.error("No chapters parsed from CHAPTER-LINKS — nothing to do.")
        sys.exit(1)

    logger.info("=" * 70)
    logger.info(f"Stage 0: scraping images for {len(chapters)} chapters")
    logger.info("=" * 70)

    session = requests.Session()
    session.headers.update({'User-Agent': _UA})

    manifest = []
    for n, ch in enumerate(chapters):
        files = scrape_chapter(session, ch)
        manifest.append({
            'num': ch['num'],
            'name': ch['name'],
            'url': ch['url'],
            'folder': f"Chapter{ch['num']}",
            'image_count': len(files),
            'files': [os.path.basename(f) for f in files],
        })
        if n < len(chapters) - 1:
            time.sleep(_REQUEST_DELAY)

    manifest_path = os.path.join(_SCRAPED_DIR, 'manifest.json')
    with open(manifest_path, 'w', encoding='utf-8') as fh:
        json.dump(manifest, fh, indent=2)

    logger.info("=" * 70)
    logger.info("Per-chapter image counts:")
    for entry in manifest:
        flag = '  ⚠️  ZERO IMAGES' if entry['image_count'] == 0 else ''
        logger.info(f"  Chapter{entry['num']:<3} {entry['name'][:45]:<45} {entry['image_count']:>3} imgs{flag}")
    logger.info("=" * 70)
    logger.info(f"Cache: {_SCRAPED_DIR}")
    logger.info(f"Manifest: {manifest_path}")
    logger.info("REVIEW the images above, prune any stray files, then run Stage 1.")


if __name__ == '__main__':
    main()
