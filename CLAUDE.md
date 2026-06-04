# CLAUDE.md <!-- version: v4.0 -->

## Maintenance Rule

**Whenever this file is modified:**
1. Increment the version number in the `<!-- version: vX.Y -->` tag on line 1 (bump minor by 0.1 each time: v1.1 → v1.2 → v1.3, etc. — only bump the major version when explicitly instructed, e.g. to v2.0 or v3.0)

---

## Git & Versioning

- **Repository:** https://github.com/technerrrd/combined-pdf-tex-scraping
- **Versioning scheme:** Minor version bumps of 0.1 per change (v1.1, v1.2, v1.3, …) until explicitly instructed to jump to a new major version (v2.0, v3.0, etc.)
- **Token:** To be provided by user before any push operations

---

## Pipeline Overview

This project converts EduRev teaching notes into clean, typeset `.tex` and `.lyx` files.

**As of v4.0 the primary path is web-only.** The scraped EduRev page is the original
source the PDF/DOCX were derived from — it is server-rendered (no JS), complete, and
carries real semantics (h2/h3/h4 headings, `<strong>` bold, `<ul><li>` lists incl.
nesting, inline full-size images in reading order, tables). Converting it directly is
more faithful than the PDF path (which lost bullet markers entirely — they survived as
neither text nor vector graphics, only indentation) and removes the fragile PDF
machinery (font-size heading guessing, ligature hacks, split-image detection,
artifact stripping). The PDF is kept **only as a validation oracle**.

```
v4.0 PRIMARY — web-only:
[Input: input/CHAPTER-LINKS  (+ module PDF, optional, for validation only)]
        ↓
  stage0/scrape_chapters.py
  (per chapter URL: fetch + cache HTML → stage2/html_source.parse_html →
   download inline images → prepend a chapter heading)
        ↓
  reuse stage2/convert.py themed writers: write_tex / write_lyx(template)
  (chapter banners + Legrand theme, real itemize lists, full-size images)
        ↓
[Output: combined .tex + .lyx + media/ in stage2/output/<module>/]
        ↓
  stage0/validate_against_pdf.py  (read-only QA)
  (per-chapter % of PDF content lines present in the HTML; flags < 95%.
   NEVER edits/overrides the HTML; skipped if no PDF present)

v3.0 FALLBACK — PDF→DOCX (retained, no longer primary):
  stage0/scrape_images.py (web images) → stage1/convert_document_v4.py (PDF text +
  web/DOCX images → DOCX) → stage2/convert.py (DOCX → .tex/.lyx). Use only when a
  chapter cannot be scraped as HTML.
```

**Why both still exist:** v4.0 depends on the page being reachable; the cached HTML
(`input/scraped/Chapter<N>/page.html`) makes converts reproducible offline, and the
PDF oracle flags any silent content loss. The v3.0 PDF pipeline remains as a fallback.

**Original projects (DO NOT MODIFY):**
- Stage 1 source: `/home/anirudh/pdf-conversion/`
- Stage 2 source: `/home/anirudh/tex_output/`

**This project's scripts are independent copies.** Changes here do not affect the originals.

### Project Structure

```
final-combined-pdf-to-lyx/
├── CLAUDE.md
├── stage0/                    ← web scraping + validation
│   ├── scrape_chapters.py     ← v4.0 PRIMARY: HTML → themed combined .tex/.lyx
│   ├── validate_against_pdf.py ← v4.0 read-only PDF coverage oracle
│   ├── scrape_images.py       ← v3.0 fallback: web image scraper
│   ├── prototype_html_to_content.py ← standalone HTML→tex/lyx demo (reference)
│   └── logs/
├── stage1/                    ← v3.0 fallback: PDF → DOCX
│   ├── convert_document_v4.py
│   ├── analyze_document.py    ← chapter detection (reused by validation oracle)
│   ├── output/
│   ├── logs/
│   └── venv/                  ← shared venv (requests, beautifulsoup4, html5lib, …)
├── stage2/                    ← DOCX/HTML → LyX/TeX
│   ├── convert.py             ← themed writers (write_tex/write_lyx) + template/banners
│   ├── html_source.py         ← v4.0 HTML front-end: parse_html() → element dicts
│   ├── output/                ← .tex, .lyx, media/ land here
│   └── logs/
├── Final-lyx_template/        ← Legrand LyX theme + assets (banners, structure.tex)
└── input/                     ← inputs
    ├── CHAPTER-LINKS          ← chapter → EduRev URL map (primary input)
    ├── scraped/Chapter<N>/    ← cached page.html (+ v3.0 image cache, manifest.json)
    └── <module>.pdf/.docx     ← optional; PDF used for validation only
```

### Common Commands

```bash
# Activate the shared venv (requests, beautifulsoup4, html5lib, pymupdf, python-docx, Pillow)
source stage1/venv/bin/activate

# === v4.0 PRIMARY: web-only ===
# Build the themed combined .tex/.lyx straight from input/CHAPTER-LINKS
python stage0/scrape_chapters.py [<module-name>]
# → stage2/output/<module>/<module>.{tex,lyx} + media/ + chapter banners + theme assets

# Validate the scrape is complete against the module PDF (read-only; skips if no PDF)
python stage0/validate_against_pdf.py
# → per-chapter coverage %; flags any chapter < 95% for review

# === v3.0 FALLBACK: PDF → DOCX (only if a chapter can't be scraped) ===
# Scrape full-size chapter images from EduRev pages listed in input/CHAPTER-LINKS
python stage0/scrape_images.py
# → downloads to input/scraped/Chapter<N>/, writes manifest.json, prints per-chapter
#   counts, then STOPS. Review/prune the images before running Stage 1.

# === STAGE 1 (fallback) ===
# Activate Stage 1 venv
source stage1/venv/bin/activate

# Analyze document structure (detect chapters before converting)
python stage1/analyze_document.py

# Run Stage 1 conversion — splits into chapter files
python stage1/convert_document_v4.py

# === STAGE 2 ===
# Convert all DOCX files placed in stage2/
python stage2/convert.py

# Convert specific file(s)
python stage2/convert.py Chapter12.docx
python stage2/convert.py Chapter10.docx Chapter11.docx
```

---

# v4.0: Web-Only Pipeline (HTML → LyX/TeX)

The primary path. `stage0/scrape_chapters.py` orchestrates; `stage2/html_source.py` is
the HTML front-end; `stage2/convert.py`'s themed writers are the reused back-end.

## How it works
1. **Parse `CHAPTER-LINKS`** (reuses `scrape_images.parse_links`) → ordered chapters.
2. **Fetch + cache** each chapter page to `input/scraped/Chapter<N>/page.html`
   (`html_source.fetch`) — reproducible offline.
3. **`html_source.parse_html(html)`** → Stage 2 element dicts, in reading order:
   - `heading` (h2→section, h3→subsection, h4→subsubsection; leading numbering stripped)
   - `body` (`<p>` with `(text,bold)` segments from `<strong>`)
   - `list` **(new element type)** — `<ul>/<ol>`, nested via depth
   - `image` — `<img …_lg.jpg>` (downloaded; not extracted from a DOCX)
   - `table` — `<table>` (the auto-TOC table is skipped)
4. **Prepend a `chapter` heading** per URL, concatenate all chapters.
5. **Reuse `write_tex` / `write_lyx(template_dir=Final-lyx_template)`** → combined themed
   `.tex`/`.lyx`; `generate_chapter_image` draws a banner per chapter (keyed by name),
   `copy_template_assets` copies theme files.

## Key implementation rules (do not regress)
- **Parse with `html5lib`**, not `html.parser`: EduRev omits closing `</p>`/`</li>` tags;
  only html5lib applies HTML5 implied-end-tag rules to give a clean sibling tree. With
  `html.parser` one `<p>` wrongly wraps the whole chapter.
- **Content scope:** `find_content_root` = smallest ancestor holding ≥80% of the
  `_lg.jpg` images. Do **not** scope by the visible content `<div>` (`explr_htmlcnt_dv`)
  — it appears *after* the images in the markup. The `_lg.jpg` path filter cleanly
  separates content images from page chrome (see v3.0 notes).
- **Inline spacing:** whitespace-only text nodes between inline tags must be collapsed to
  a single trailing space, not dropped — otherwise `<strong>active</strong>and` joins to
  "activeand". (`_segments` / `list_items` handle this.)
- **Unicode normalization** (`clean_text` + `_UNICODE_MAP`): map thin/zero-width spaces,
  curly quotes, dashes, and — for science chapters — sub/superscripts and arrows to
  plain ASCII, since pdflatex's `utf8` rejects them. Values must be plain text (no
  `$ _ ^ \ {}`) because `tex_escape` runs afterward.
- **`write_lyx` requires the template** (`Final-lyx_template/main.lyx`) for the theme and
  for `generate_chapter_image` banners; without it you get a bare standalone book.

## Known limitations / follow-ups
- **Math & inline sub/superscripts** are flattened to ASCII (e.g. `CO₂`→`CO2`). Formulas
  on EduRev are images and pass through as images (same as the PDF path — not a
  regression). Faithful math typesetting (`CO$_2$`, `$\to$`) is a future improvement.
- **Tables:** these chapter-notes pages have no content tables (only the auto-TOC, which
  is skipped); the `table` handler exists but is unproven on real content tables.
- **MCQ:** chapter-notes pages have none; question-bank pages would need MCQ detection in
  `parse_html` to reuse Stage 2's MCQ writer.

## PDF validation oracle (`stage0/validate_against_pdf.py`)
Read-only. Uses `analyze_pdf` (Stage 1) for per-chapter PDF page ranges, normalizes both
sides, and reports the % of PDF content lines (artifacts removed) present in each
chapter's cached HTML. Flags any chapter < 95%; never edits output; skipped when no PDF
is in `input/`. Measured on the Class 8 module: 97.7–100% per chapter.

---

# Stage 0: Web Image Scrape (NEW in v3.0)

## Why

EduRev PDFs are pages printed from Firefox. Images that straddle a page break get
split into two halves, and Stage 1 carried complex split-detection logic to cope —
still imperfect. The **original EduRev web pages** host every content image full-size
on a CDN, in reading order, in server-rendered HTML (no JavaScript). Stage 0 fetches
those clean images so Stage 1 never has to deal with split halves.

## Input: `input/CHAPTER-LINKS`

One line per chapter: `<idx>\t<EduRev URL> - Chapter<N>`. The scraper extracts the
URL, the chapter number `<N>`, and the chapter name (from the URL slug).

## How `stage0/scrape_images.py` works

1. Parse `CHAPTER-LINKS` into ordered `{num, name, url}` entries.
2. `GET` each page with a desktop User-Agent (polite ~1.5s delay between pages).
3. Extract content image URLs by the path filter
   **`ApplicationImages/Temp/<UUID>_lg.jpg`** — this uniquely identifies uploaded
   content images. ALL page chrome (icons/ads/course thumbnails/favicons) lives on
   other CDN paths (`cdn_assets/`, `cdn_lib/`, `CourseImages/`, `..._icon.jpg`), so the
   single path filter cleanly separates content from junk. Dedupe, preserve order.
   **Do NOT scope to the visible content `<div>`** (`explr_htmlcnt_dv`) — it appears
   *after* the images in the markup, so scoping by it drops everything.
4. Download to `input/scraped/Chapter<N>/image1.jpg, image2.jpg, …` (numeric order so
   Stage 1's sort works). Skips files already present (cache).
5. Write `input/scraped/manifest.json` (`[{num, name, url, folder, image_count, files}]`)
   and print per-chapter counts, then **STOP** (review gate). Inspect/prune, then
   run Stage 1.

## Notes

- Dependency: `requests` (in the Stage 1 venv).
- Images are downloaded as-is; these EduRev JPEGs sometimes start with a DQT marker
  and lack a JFIF/EXIF header, which makes `python-docx` raise `UnrecognizedImageError`.
  Stage 1's `insert_image()` handles this by re-encoding through PIL on failure — do
  not remove that fallback.
- If a chapter's scrape fails or yields zero images, Stage 1 falls back to DOCX-zip
  images for the whole document (see Stage 1 image-source rule).

---

# Stage 1: PDF → DOCX

## Why Both Input Files Are Needed

These documents are EduRev teaching notes printed from Firefox to PDF, then converted to DOCX. This creates specific structural issues:

- **DOCX text is inaccessible:** All text lives inside text boxes/shapes which `python-docx` cannot read. DOCX paragraphs contain only images and empty lines — zero readable text.
- **PDF is the only text source:** All text must be extracted from the PDF using `pymupdf`.
- **Image source (v3.0):** Images come from **Stage 0 web scrape** (`input/scraped/`) when available — these are full-size and never split. If no scrape exists, Stage 1 falls back to the **DOCX zip** (legacy behavior). PDF images are never used as a source (they are sometimes split in half across page boundaries; the PDF only supplies image *positions*).

**Text:** Extract from PDF ONLY  
**Images:** Stage-0 scraped web images (primary) → DOCX zip (fallback). Never from the PDF.

**Workflow:** Build a brand new DOCX from scratch — text from PDF, images from `input/scraped/` (or DOCX zip), processed page by page in reading order. The PDF's image slots (page + reading-order position) determine *where* each image lands; the scraped list supplies *which* image (1:1, in document order).

## Chapter Detection and Processing

**Process:**
1. **Analyze document structure** — detect chapter boundaries by identifying Heading 1 text (19–21pt) containing "Chapter Notes:" or "Chapter Notes-" (colon or dash separator).
2. **Confirm with user** — before processing, report findings:
   - "Found 1 chapter: Chapter 12 - Beyond Earth"
   - "Found 3 chapters: Chapter 10, Chapter 11, Chapter 12"
   - Ask: "Process as single file or split into separate files per chapter?"
3. **Processing options:**
   - **Single file:** `converter.convert()`
   - **Split by chapter:** `converter.convert_chapters(chapters)` where `chapters` is a list of `{name, start_page, end_page, output_path}` dicts
   - **Output naming for splits:** `<Subject>-Ch<N>-<Chapter-Name>.docx` (e.g. `Class6-Ch1-Mindful-Eating.docx`)

## Formatting Specifications

### Font and Color
- **Font:** Times New Roman ONLY
- **Color:** `#000000` (black) throughout — strip all source colors and font families

### Heading Styles
- **Heading 1:** 16pt bold
- **Heading 2:** 14pt bold
- **Heading 3:** 12pt bold
- **Body / captions / table cells / bullets / numbered lists:** 12pt normal

### Page Layout
- **Page size:** A4 (11906 × 16838 DXA)
- **Margins:** 1080 DXA on all four sides (0.75 inches)
- **Line spacing:** Single
- **Paragraph spacing:** 0pt before, 8pt after

### Content Rules
- **Images:** Center-align all images
- **Page/section breaks:** Let content flow naturally — do not preserve original breaks
- **Tables:** Preserve structure; apply 12pt Times New Roman to all cells; simple grid borders (black, single line)

### Heading Detection (PDF font size → output heading)
- **19–21pt** → Heading 1 (16pt bold black). Covers 21pt colored section titles and 19.5pt black "Chapter Notes: …" / "Chapter Notes- …" headings (colon or dash separator)
- **16–17pt** → Heading 2 (14pt bold black). Range covers 16.5pt used in newer documents
- **~13.5pt + bold + standalone** → Heading 3 (12pt bold black)
- **~13.5pt normal** → Body text (12pt normal black)

### Elements to Strip
- **EduRev branding:** text containing "edurev" or "durev" (case-insensitive), images under 1KB, EduRev URLs, EduRev logo PNG (~301×112px)
- **Browser print artifacts:** "Firefox", `^\d+\s+of\s+\d+$` (page counter), `^\d{2}/\d{2}/\d{2},\s+\d{2}:\d{2}$` (timestamp)
- **MCQ blocks:** "MULTIPLE CHOICE QUESTION", "Try yourself:" (start marker), question text, "View Solution" (end marker), option labels A/B/C/D, "View More"
- **Headers and footers:** strip entirely
- **Hyperlinks:** convert to plain text
- **Footnotes and endnotes:** remove entirely
- **TOC / Index:** remove any auto-generated TOC pages (detect by `\.{3,}\s*\d+$` pattern)

## Technical Implementation

### Python Environment

```bash
source stage1/venv/bin/activate
# Dependencies: pymupdf>=1.23.0, python-docx>=1.1.0, Pillow>=10.0.0, requests (Stage 0)
```

### Image Source Resolution (v3.0)

Before building, Stage 1 picks an image source:
- **`convert()` (single combined file):** `_collect_all_scraped_images()` concatenates
  every chapter's `input/scraped/Chapter<N>/` images in `manifest.json` (document)
  order. If the manifest is missing or any folder is empty → falls back to DOCX-zip.
  The concatenated list is assigned 1:1 onto `active_slots_ordered`.
- **`convert_chapters()` (split files):** `_resolve_scraped_images()` maps each chapter
  to its scraped folder **positionally** (both the PDF chapters and the manifest are in
  document order — chapter *names* are only a sanity check, since PDF headings are
  sometimes truncated mid-line). Each chapter gets its own scraped deque, popped 1:1 at
  that chapter's PDF image slots. All-or-nothing: if any chapter can't resolve, the
  whole run falls back to DOCX-zip.

When scraped images are used, the split-detection/dimension-matching logic still runs
on the PDF to locate image *slots*, but scraped images are placed by simple in-order
pop (`_insert_scraped_image_item`) — no dimension matching, because they are clean and
never split.

### Implementation Approach

`convert_document_v4.py` builds a brand new DOCX from scratch:

1. **Resolve image source** — scraped web images (`input/scraped/`) if available, else
   **extract from DOCX zip** — unzip, read `word/media/`, sort by filename number; filter out files under 1KB (watermarks) and PNG files with 270–330 × 90–130px (EduRev logo).

2. **Pre-scan PDF images** — walk all pages with `pdf_doc.extract_image(xref)` to get pixel dimensions. Build `split_set` for: consecutive same-dimension slots (Firefox page-break splits) and logo-dimension slots. Build `pdf_img_dims` dict: `(page_num, xref) → (px_w, px_h, pdf_pt_w)`. Build the image queue (per-chapter scraped deques, or one DOCX `docx_queue`).

3. **Process PDF page by page** — for each page, extract text blocks AND image positions via `pymupdf`. Sort by y-coordinate. For text: skip artifacts and EduRev text, detect heading level, sanitize, add to document. For images: check split_set and queue.

4. **Image insertion (queue-based):**
   - If `(page_num, xref)` is in `split_set` → skip (Firefox split continuation)
   - **Scraped mode:** pop the next scraped image in order and insert (no dimension match)
   - **DOCX-zip mode:** flush orphan images (if front of queue doesn't dimension-match, pop and insert as orphan — handles images in shapes pymupdf doesn't enumerate), then pop and insert the matched image
   - After all pages: insert any remaining queued images (trailing — per chapter in scraped mode)
   - Insert with `space_before=Pt(0)`, `space_after=Pt(0)` — do NOT set `line_spacing_rule` on image paragraphs
   - **`insert_image()` re-encodes through PIL on `add_picture` failure** — scraped EduRev JPEGs may lack the JFIF/EXIF header python-docx requires. Keep this fallback.

5. **Image sizing** — `width = Inches(min(r.width / 72, 6.77))` where `6.77` is A4 content width (inches). Orphan/trailing images default to full content width. Only `width` is passed to `add_picture()`; height auto-scales.

6. **Text formatting** — detect heading levels from PDF font size/color/bold. Apply Times New Roman, black, correct sizes. Inline bold preserved via parallel `bold_chars` map (one entry per character from rawdict spans), propagated through ligature expansion, grouped into `runs = [(text, is_bold), …]`. Headings force all runs bold.

### Key Rules

- **Never read DOCX paragraphs for text** — all text is in text boxes. Use PDF exclusively.
- **Bold detection:** `is_bold = bool(span["flags"] & 16) or bool(_BOLD_FONT_RE.search(span.get("font", "")))` where `_BOLD_FONT_RE = re.compile(r'bold|black|heavy|semibold', re.IGNORECASE)`. Never rely on the flags bit alone — EduRev PDFs use `Lato-Black` (flags=0) for visually bold body text.
- **Run boundary spaces:** use `re.sub(r'\s+', ' ', seg)` without `.strip()` when building run segments — stripping drops boundary spaces and joins words.
- **MCQ state machine:** set `in_mcq_block=True` on "Try yourself:", clear on "View Solution". MCQ option texts appear AFTER "View Solution" in extraction order — track `after_option_label=True` when A/B/C/D label is seen and skip the next text block.
- **Split image detection:** identical pixel dimensions on consecutive pages = Firefox split. Do NOT use xref deduplication or y-position heuristics.

---

# Stage 2: DOCX → LyX/TeX

## Core Principle

**You are a transcription worker, not a content generator.**

Your sole job is to faithfully convert the content of the input DOCX into LaTeX/LyX format. Do not:
- Add, infer, or rewrite any text
- Fill in missing content or "improve" phrasing
- Generate section summaries, captions, or explanations

If something is unclear or missing in the source, flag it — do not invent a replacement.

### Subject Detection

Before beginning conversion, scan the document title, chapter heading, or metadata for subject indicators. If the subject is not immediately obvious, ask: **"Is this a Math or Science document?"** and wait for confirmation before applying subject-specific spacing rules.

### Input File Characteristics

Input DOCX files are clean and formatted (produced by Stage 1):
- **Text:** ALL paragraphs use `Normal` style — heading level is determined by **font size**, not style name
- **Font:** Times New Roman throughout — bold is **run-level**. `<w:b val="0"/>` = explicitly not bold. Many paragraphs have mixed bold/plain runs. Do **not** use bold as a heading indicator
- **Images:** Embedded inline, center-aligned
- **Tables:** Simple grid borders, 12pt Times New Roman cell text
- **No watermarks, headers/footers, or hyperlinks** — stripped in Stage 1
- **MCQ blocks are present** and must be formatted per MCQ Rules below

## Tool Chain

**Custom Python XML parser** (`convert.py`) reads `word/document.xml` directly from the DOCX ZIP.  
Pandoc is not used — it cannot reconstruct heading hierarchy because all paragraphs are `Normal` style.

### Font Size → Structure Mapping

| Font size (`w:sz`) | Point size | LaTeX output | LyX layout |
|--------------------|------------|--------------|------------|
| 32 + `Chapter Notes:` prefix | 16pt | `\chapter{}` | `Chapter` |
| 32 | 16pt | `\section{}` | `Section` |
| 28 | 14pt | `\subsection{}` | `Subsection` |
| 24 | 12pt | body paragraph (`\par`) | `Standard` |
| other | varies | body paragraph (`\par`) | `Standard` |

**Chapter detection:** An sz=32 paragraph whose text matches `Chapter Notes: <name>` (case-insensitive) is treated as a chapter heading. The "Chapter Notes:" prefix is stripped; only `<name>` appears in the output. Document class is `book`.

**Single-chapter files** (no `Chapter Notes:` headings) still use `book` class.

## Numbering Rules

**Strip all original numbering prefixes** from heading and body text. Patterns stripped: `1.`, `1.1`, `1.6.9`, `Q1`, `Q1.`, `(a)`, `(i)`, `(1)`. LyX and LaTeX provide their own auto-numbering.

### Notes documents (default)

All sz=24 body paragraphs are **Standard paragraphs** — no enumerate, no numbering. Bold formatting preserved run-by-run.

**`.tex` body paragraph:**
```latex
\par Plain text here.
\par \textbf{Bold label:} Plain description here.
```

**`.lyx` body paragraph:**
```
\begin_layout Standard
Plain text here.
\end_layout

\begin_layout Standard
\series bold
Bold label:
\series default
 Plain description here.
\end_layout
```

### Question bank documents

Enumerate blocks apply only when the document is a question bank (MCQ, short/long Q&A). In that case:
- **Strip all original numbering** — remove `Q1`, `2.`, `1.1`, `(a)`, etc.
- **Under every subsection**, open a fresh `enumerate` block immediately after the heading
- **Before the next subsection or section**, close the `enumerate` block

MCQ questions always use enumerate regardless of document type.

## MCQ Rules

### All formats
- **Strip all original option markers** — remove `(a)`, `(i)`, `1)`, `A.`, etc. from option text
- **NEVER wrap the option table in `\begin{center}`**
- The question itself is an `\item` in the enclosing `enumerate` block

### `.tex` format

```latex
\item \textbf{<QUESTION>}\\[0.13cm]
\begin{tabular}{@{}p{0.45\textwidth} p(0.45\textwidth)@{}}
$\square$ A) <OPTION_1> & $\square$ B) <OPTION_2> \\
$\square$ C) <OPTION_3> & $\square$ D) <OPTION_4>
\end{tabular}
```

### `.lyx` format

```
\begin_layout Enumerate
\begin_inset ERT
status open

\begin_layout Plain Layout
\backslash
textbf{<QUESTION>}
\backslash
\backslash
[0.13cm]
\end_layout

\end_inset

\begin_inset ERT
status open

\begin_layout Plain Layout
\backslash
begin{tabular}{@{}p{0.45
\backslash
textwidth} p{0.45
\backslash
textwidth}@{}}
\end_layout

\begin_layout Plain Layout
$
\backslash
square$ A) <OPTION_1> & $
\backslash
square$ B) <OPTION_2>
\backslash
\backslash

\end_layout

\begin_layout Plain Layout
$
\backslash
square$ C) <OPTION_3> & $
\backslash
square$ D) <OPTION_4>
\end_layout

\begin_layout Plain Layout
\backslash
end{tabular}
\end_layout

\end_inset

\end_layout
```

### Assertion-Reasoning Lock

Every Assertion-Reasoning pair **must** be followed by these exact fixed options — never alter the wording.

**`.tex` format:**

```latex
\item \textbf{Assertion (A):} <ASSERTION TEXT>\\[0.06cm]
\textbf{Reason (R):} <REASON TEXT>\\[0.13cm]
\begin{tabular}{@{}p{0.45\textwidth} p{0.45\textwidth}@{}}
$\square$ A) Both Assertion and Reason are correct; Reason is correct explanation. &
$\square$ B) Both Assertion and Reason are correct; Reason is NOT correct explanation. \\
$\square$ C) Assertion is correct; Reason is incorrect. &
$\square$ D) Assertion is incorrect; Reason is correct.
\end{tabular}
```

**`.lyx` format:**

```
\begin_layout Enumerate
\begin_inset ERT
status open

\begin_layout Plain Layout
\backslash
textbf{Assertion (A):} <ASSERTION TEXT>
\backslash
\backslash
[0.06cm]
\backslash
textbf{Reason (R):} <REASON TEXT>
\backslash
\backslash
[0.13cm]
\end_layout

\end_inset

\begin_inset ERT
status open

\begin_layout Plain Layout
\backslash
begin{tabular}{@{}p{0.45
\backslash
textwidth} p{0.45
\backslash
textwidth}@{}}
\end_layout

\begin_layout Plain Layout
$
\backslash
square$ A) Both Assertion and Reason are correct; Reason is correct explanation. &
$
\backslash
square$ B) Both Assertion and Reason are correct; Reason is NOT correct explanation.
\backslash
\backslash

\end_layout

\begin_layout Plain Layout
$
\backslash
square$ C) Assertion is correct; Reason is incorrect. &
$
\backslash
square$ D) Assertion is incorrect; Reason is correct.
\end_layout

\begin_layout Plain Layout
\backslash
end{tabular}
\end_layout

\end_inset

\end_layout
```

## Subject-Specific Spacing Rules

### Short vs Long Question Detection

The source document will explicitly label sections as **"Long Question Answers"** or **"Short Question Answers"** — use that heading to determine spacing. For **case-based questions**, treat as Short Question or MCQ as indicated in the question itself.

### IF MATH

Do **not** add rules or large vertical spaces.

| Type | `.tex` | `.lyx` |
|------|--------|--------|
| Regular question | `\item \textbf{<QUESTION>}` | `Enumerate` layout, bold via `\series bold` |
| True/False | `\item \textbf{<STATEMENT>} \hfill ________` | ERT for `\hfill ________` after bold text |

### IF SCIENCE

For every question that is **not** an MCQ or fill-in-the-blank, and where **no answer is provided** in the source text, append writing lines.

**Short questions — 1 rule line:**

`.tex`:
```latex
\item \textbf{<QUESTION>}
\par \vspace{0.3cm} \noindent\rule{\linewidth}{0.4pt} \vspace{0.5cm}
```

`.lyx`:
```
\begin_layout Enumerate
\series bold
<QUESTION>
\series default

\begin_inset ERT
status open

\begin_layout Plain Layout
\backslash
par
\backslash
vspace{0.3cm}
\backslash
noindent
\backslash
rule{
\backslash
linewidth}{0.4pt}
\backslash
vspace{0.5cm}
\end_layout

\end_inset

\end_layout
```

**Long questions — 3 rule lines:**

`.tex`:
```latex
\item \textbf{<QUESTION>}
\par \vspace{0.3cm} \noindent\rule{\linewidth}{0.4pt}
\par \vspace{0.4cm} \noindent\rule{\linewidth}{0.4pt}
\par \vspace{0.4cm} \noindent\rule{\linewidth}{0.4pt} \vspace{0.5cm}
```

`.lyx`:
```
\begin_layout Enumerate
\series bold
<QUESTION>
\series default

\begin_inset ERT
status open

\begin_layout Plain Layout
\backslash
par
\backslash
vspace{0.3cm}
\backslash
noindent
\backslash
rule{
\backslash
linewidth}{0.4pt}
\backslash
par
\backslash
vspace{0.4cm}
\backslash
noindent
\backslash
rule{
\backslash
linewidth}{0.4pt}
\backslash
par
\backslash
vspace{0.4cm}
\backslash
noindent
\backslash
rule{
\backslash
linewidth}{0.4pt}
\backslash
vspace{0.5cm}
\end_layout

\end_inset

\end_layout
```

**Fill in the blank** (Math and Science):

`.tex`: `\item \textbf{<TEXT> _________.}`  
`.lyx`: `Enumerate` layout, bold text, underscores inline

**True/False** (Math and Science):

`.tex`: `\item \textbf{<STATEMENT>} \hfill ________`  
`.lyx`:
```
\begin_layout Enumerate
\series bold
<STATEMENT>
\series default

\begin_inset ERT
status open

\begin_layout Plain Layout
\backslash
hfill ________
\end_layout

\end_inset

\end_layout
```

## Math & Currency Lock

- **Math:** Wrap all mathematical expressions, variables, and equations in `$...$` (inline) or `$$...$$` (display). Do NOT wrap plain numeric text like years or counts.
- **Currency:** Convert all ₹, Rs., and INR to `\rupee~<amount>` (e.g. ₹250 → `\rupee~250`)
- **No image transcription:** Strictly forbidden from transcribing text or formulas found inside an image. **Exception:** tables inside images may be transcribed.

**`.lyx` math format:**

Inline: `\begin_inset Formula $<MATH>$\n\end_inset`  
Display: `\begin_inset Formula \n$$<MATH>$$\n\end_inset`

## Automated Image Numbering & Smart Scaling

`convert.py` extracts images and names them automatically (`image1.png`, `image2.jpeg`, etc.). **Do not rename them.** Map each image to its assigned filename in document order.

**Sizing:** Read rendered display size from DOCX XML (`<wp:extent cx="..." cy="..."/>`, in EMUs where 914400 EMU = 1 inch). Pick the closest width from: `0.25`, `0.4`, `0.5`, `0.6`, `0.75`. Do not use raw pixel dimensions.

**`.tex` figure block (no `\caption{}`):**
```latex
\begin{figure}[h]
\centering
\includegraphics[width=<CHOSEN_SCALE>\textwidth]{<filename>}
\end{figure}
```

**`.lyx` figure block (no caption, inline — no float wrapper):**
```
\begin_layout Standard
\align center
\begin_inset Graphics
	filename media/<filename>
	width <SCALE>text%
\end_inset

\end_layout
```

Scale mapping for `.lyx`: `0.25` → `25text%`, `0.4` → `40text%`, `0.5` → `50text%`, `0.6` → `60text%`, `0.75` → `75text%`

## No Image Hallucination

**Never** insert a figure block for an image that does not exist in the source document. Only reference images explicitly present in the user-provided source text.

## Table Transcription Exception

You are allowed to transcribe tables (including those originally presented as images) into standard LaTeX `tabular` environments.

## Answer/Solution Exception

If the source text contains `Ans:`, `Answer:`, or `Solution:`, treat it as provided content. Place it immediately after the question item.

**`.tex`:** `\par \textit{Ans: <TRANSCRIPT_CONTENT>}`

**`.lyx`:**
```
\begin_layout Standard

\shape italic
Ans: <TRANSCRIPT_CONTENT>
\shape default

\end_layout
```

**Critical:** Do **not** append writing lines (`\rule`) or `\vspace` when an answer is already present in the source.

## Rule Priority Order

When multiple rules could apply, use this precedence (highest wins):

1. **Answer/Solution Exception** — answer present → no writing lines, ever
2. **MCQ format** — A/B/C/D options present → use MCQ grid, no writing lines
3. **Fill-in-the-blank** — question has blanks → use underscores, no writing lines
4. **Case-based question** — explicitly stated → follow Short/Long as indicated in the question
5. **Section heading** — "Long Question Answers" / "Short Question Answers" in source
6. **Subject default** — Math: no lines; Science: short = 1 rule, long = 3 rules

## Stage 2 Implementation Notes

- Skip temp files matching `~$*.docx` (created by Word when a file is open)
- Place DOCX files to convert directly in `stage2/` — do not read from subdirectories
- Log to `stage2/logs/convert_YYYYMMDD_HHMMSS.log` and echo to terminal
- Output dir `stage2/output/<stem>/` is created if it does not exist

## Verification

1. Copy a DOCX from `stage1/output/` into `stage2/`
2. Run `python stage2/convert.py`
3. Open `stage2/output/<name>/<name>.tex` — verify headings, body text, image references
4. Open `stage2/output/<name>/<name>.lyx` directly in LyX (File → Open)
5. Confirm images are in `stage2/output/<name>/media/` and referenced correctly

## Required LaTeX Packages

Add to your document preamble in LyX via **Document → Settings → LaTeX Preamble**:

| Package | Required for |
|---------|-------------|
| `graphicx` | `\includegraphics{}` in figure blocks |
| `rupee` | `\rupee` currency symbol |
| `amssymb` | `$\square$` checkbox in MCQ grids |
