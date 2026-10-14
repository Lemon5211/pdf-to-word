# pdf-to-word - a Codex skill

Turn a designed PDF into an editable Word `.docx` that still looks like the
original: same page count and page breaks, same text, same colours, artwork in
the same places, and every sentence as real editable text - never an image of
text pretending to be text.

## What it does

* **Inspects** the PDF: text runs with font/size/colour/bold, vector shapes
  (ribbons, rules, bands, margin bars) with fill colours, image bounding boxes
  and hyperlinks - as a readable report plus JSON.
* **Rebuilds** it in Word with constructs that match the design: paragraph
  borders, run shading, fixed tables, character-spacing label fields, floating
  pictures placed at the source coordinates, Wingdings bullets, real hyperlinks.
* **Verifies** by rendering the result with Word or LibreOffice and comparing it
  against the source: character-for-character content, page count and page
  breaks, picture positions, link targets.
* **Reports** what was approximated (straightened ribbons, artwork re-inserted as
  pictures, sub-point artefacts dropped) instead of guessing silently.

## Requirements

* Python 3 with `pdfplumber`, `python-docx`, `Pillow`
* Microsoft Word or LibreOffice for the render-and-verify step

## Install

Clone into your Codex skills directory:

```bash
git clone https://github.com/Lemon5211/pdf-to-word.git ~/.codex/skills/pdf-to-word
```

On Windows:

```powershell
git clone https://github.com/Lemon5211/pdf-to-word.git "$env:USERPROFILE\.codex\skills\pdf-to-word"
```

## Use

Ask your agent to "convert this PDF to Word", or invoke the skill explicitly.
The workflow is:

```bash
python scripts/inspect_pdf.py INPUT.pdf --out work --json --render
python scripts/crop_pdf_image.py INPUT.pdf --page 1 --auto --out-dir work
python scripts/build_template.py --out work/out.docx    # copy and edit per document
python scripts/render_docx.py work/out.docx
python scripts/verify_conversion.py --source INPUT.pdf --docx work/out.docx --rendered work/out.pdf
```

## Layout

```
SKILL.md                              workflow, boundaries, hand-off rules
references/layout-fidelity.md         PDF construct -> Word construct mapping + pitfalls
references/rendering-and-verification.md  headless rendering, verification criteria
scripts/ooxml.py                      the OOXML primitives python-docx lacks
scripts/inspect_pdf.py                analyse a PDF (text, vectors, images, links)
scripts/crop_pdf_image.py             extract artwork at high DPI
scripts/build_template.py             per-document build skeleton
scripts/render_docx.py                docx -> PDF via Word or LibreOffice
scripts/verify_conversion.py          content + geometry verification
```

## Scope

* Scanned PDFs without a text layer are out of scope - run OCR first.
* Slanted ribbons and similar shapes are approximated with straight bands unless
  you ask for the exact angle.

## License

See `LICENSE` if present.