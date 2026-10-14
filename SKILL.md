---
name: pdf-to-word
description: Convert a designed PDF into an editable Word .docx that keeps the original layout, colours, artwork and pagination. Use when a PDF (resume, one-pager, report, flyer, form) has to become a Word document. Not for editing a PDF in place, and not for scanned PDFs without a text layer (those need OCR first).
metadata:
  short-description: Convert PDF to editable Word
---

# PDF to Word

Produce a `.docx` that a person can edit in Word or WPS and that still looks like
the source PDF: same page count and page breaks, same text, same colours and
artwork in the same places, with every sentence as real editable text.

## Boundaries

* A text layer is required. `inspect_pdf.py` reports when a page has none - that
  is a scan, so say OCR is needed instead of quietly shipping pictures of text.
* Never rasterise text to imitate the design, and never summarise or "improve"
  the wording: the user asked for a conversion.
* Content the source got wrong (a dangling `、`, a truncated sentence, an odd
  quote mark) stays as it is, and is mentioned in the hand-off.

## Workflow

Work in a scratch folder per document; keep the build script there so the
conversion can be re-run after an edit.

### 1. Inspect the source

```bash
python scripts/inspect_pdf.py INPUT.pdf --out work --json --render
```

Read `work/report.txt`. It lists, per page: page size, images and links with
coordinates, vector shapes with fill colours (ribbons, rules, bands, margin
bars), and text rows with per-run bold/size/font/colour/x-range.

Decide before writing any code:

* what is prose (becomes flowing paragraphs) versus what is artwork (photo,
  logo, icon badge - gets cropped and re-inserted);
* the page geometry: margins, the text box, the line rhythm;
* whether the source's own label columns, bands or tables need real Word
  constructs. `references/layout-fidelity.md` maps each PDF construct to the
  Word equivalent with the measured numbers.

### 2. Rebuild in Word

Copy `scripts/build_template.py` into the scratch folder, replace its CONTENT
block with the inspected text, and adjust the geometry constants. It already
demonstrates the whole recipe (header block, accent band, section bars, form
grid, floating picture, bullets) on top of `scripts/ooxml.py`, which holds the
OOXML primitives python-docx lacks.

Crop any artwork with `scripts/crop_pdf_image.py` and place it with the PDF's own
image bbox so nothing has to be eyeballed:

```bash
python scripts/crop_pdf_image.py INPUT.pdf --page 1 --auto --out-dir work
python scripts/crop_pdf_image.py INPUT.pdf --page 1 --bbox 400,30,500,60 --out work/badges.png
```

### 3. Verify by rendering

```bash
python scripts/render_docx.py work/out.docx
python scripts/verify_conversion.py --source INPUT.pdf --docx work/out.docx --rendered work/out.pdf
```

Do not deliver on an unrendered `VERDICT: FAIL`. The expected result is
identical text, the same page count with the same section starting each page,
and artwork within a few points of the source. Iterate on the cause of any
drift - `references/rendering-and-verification.md` explains how to read the
report and what to change.

Use `inspect_pdf.py --ascii` to confirm the visual structure in text form (bands,
bars, photo placement) when images cannot be viewed directly.

### 4. Hand off

Put the `.docx` next to the source file unless the user asked for a different
location, and state plainly:

* what is faithful (text, page breaks, artwork positions, links);
* what was approximated (slanted ribbons straightened, decorative artwork
  re-inserted as a picture, sub-point artefacts dropped);
* anything in the source that looks like an authoring mistake.

## Reference material

* `references/layout-fidelity.md` - PDF construct → Word construct mapping, the
  geometry conversions, a worked two-page resume example, and the OOXML pitfalls
  that cost real debugging time (run shading, merged tables, floating pictures
  inside cells, East Asian fonts, hanging punctuation).
* `references/rendering-and-verification.md` - runtimes, how to drive Word
  headlessly when its COM type library is broken, what the verification report
  means, and how to iterate on drift.
