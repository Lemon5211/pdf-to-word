# Rendering and verifying the result

A `.docx` that "looks right" in the building script can still paginate wrong,
clip a heading, or drop a picture. The only cheap way to know is to let a real
layout engine render it and then compare coordinates against the source PDF.

## Runtimes

The scripts need `pdfplumber`, `python-docx` and `Pillow`.

* Prefer the runtime the host already provides: call the workspace-dependencies
  tool and use its Python path (it ships all three libraries).
* Otherwise use any Python 3 with the packages, installing if needed:
  `python -m pip install python-docx pdfplumber pillow`.
* `render_docx.py` additionally needs Microsoft Word (Windows/macOS) or
  LibreOffice.

## Render

```bash
python scripts/render_docx.py out.docx              # -> out.pdf
python scripts/render_docx.py out.docx --engine soffice
```

Then read the produced PDF: page count, where the page break falls, whether
anything bleeds off the page, whether the pictures are where the source had
them.

Two environment quirks are already handled inside the script - keep them in mind
because the raw errors are misleading:

* **Word COM type library failure.** Some Office installs raise
  `TYPE_E_CANTLOADLIBRARY` / "Unable to cast COM object ... to interface type
  _Application" on the first property access. Driving Word through IDispatch
  (`InvokeMember`) works in that state, which is what the generated PowerShell
  does. Do not conclude "Word is broken" and skip verification.
* **Windows PowerShell 5.1 reads `.ps1` as ANSI.** A script containing a
  non-ASCII path is silently mangled, and Word then reports "the document may be
  corrupted" for a file that is perfectly fine. Stage the `.docx` under an
  ASCII-only temp name.

`Documents.Open(...).ComputeStatistics(2)` can report a stale page count right
after opening. Count the pages of the rendered PDF instead.

If no renderer exists, say so: deliver the `.docx` and state that layout
verification was not possible. Do not claim the layout matches.

## Verify

```bash
python scripts/verify_conversion.py --source in.pdf --docx out.docx --rendered out.pdf
```

The script prints `VERDICT: PASS|FAIL` and checks:

* **content** - every character of the source against the `.docx`, whitespace
  insensitive; the multiset diff catches dropped or duplicated characters even
  when reading order legitimately differs, the sequence diff catches reordering.
  Content must be identical (ratio `1.00000`); any `MISSING from docx` entry is a
  real defect.
* **pagination** - same page count; the same section must start the same page.
* **geometry** - pictures within a few points of the source position, hyperlinks
  preserved, and the vertical drift of rows matched by text.

A row-count difference is expected and fine: re-joined paragraphs re-wrap, so a
paragraph may occupy one line more or less than in the PDF. Row drift of a few
points is also normal; drift that grows down the page means the paragraph rhythm
is wrong (`space_before`/`space_after`/line height).

## Looking at the page without image support

If images cannot be viewed directly, both inspection tools can print ASCII art,
which is enough to confirm bands, bars, blocks and picture placement:

```bash
python scripts/inspect_pdf.py out.pdf --ascii --cols 110
python scripts/inspect_pdf.py in.pdf  --ascii --crop 0,0,595,150     # header area
```

`--render` writes page PNGs if a human wants to look at them.

## Iterating

Read the drift numbers from `verify_conversion.py` and correct the cause rather
than the symptom:

* everything below a certain block is shifted → that block's height is wrong
  (header row height, band thickness, `space_before`)
* the last line of page 1 spills onto page 2 → trim a few points of spacing or
  relax a row height; forcing a page break does not fix overflow
* one picture is at the wrong offset → it is anchored in a table cell
* a specific run is in the wrong typeface → `w:rFonts/@w:eastAsia` or the Latin
  font is missing on that run

Finish by putting the `.docx` where the user expects it (next to the source file
unless they said otherwise) and stating plainly what was approximated - slanted
ribbons, decorative artwork re-inserted as pictures, dropped sub-point artefacts.
