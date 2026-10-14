# Rebuilding a PDF's design in Word

Read this when deciding *how* a piece of the source PDF becomes Word content.
The inspection report (`inspect_pdf.py`) gives the inputs: text runs with
bold/size/font/colour, vector shapes with fill colours, image bounding boxes.

## Two rules that decide everything

1. **Prose stays prose.** Every sentence becomes a normal paragraph that Word
   wraps itself. Re-joining PDF line breaks into paragraphs is what makes the
   result editable; forcing the original hard breaks produces a file that
   shatters as soon as it is edited.
2. **Artwork becomes a picture.** Photos, logos, icon badges, and textured
   backgrounds have no text equivalent. Crop them out of the PDF at high DPI
   (`crop_pdf_image.py`) and re-insert them at the same page coordinates.

Everything else is a judgement call between *flow* (paragraphs/tables, editable)
and *absolute position* (floating pictures). Prefer flow; use floating only for
things that must not move.

## Construct mapping

| In the PDF | In Word | API |
| --- | --- | --- |
| Full-width or two-tone accent band | borderless table row with shaded cells, `trHeight` in points | `fixed_table`, `shade_cell`, `row_height` |
| Coloured ribbon behind a section title | run-level shading (`w:shd` on `rPr`) + `\u3000` padding runs | `style_run(..., shading=)` |
| Thin rule running under a section title | paragraph border on the heading paragraph (`w:sz` = eighths of a point, 6 = 0.75pt) | `para_border` |
| Vertical accent line down the margin | single-sided page border, `offsetFrom="text"`, `space` in pt | `page_border` |
| Label/value form fields | fixed 4-column table (`label \| value \| label \| value`) | `fixed_table`, `row_height` |
| Labels that look 分散对齐 (e.g. `姓`…`名` spread over 3 cells) | character spacing on the label run: `pt` = field_width − chars×size, applied per character | `style_run(..., spacing_pt=)` |
| Photo / logo / icon badge | floating picture at the PDF's image bbox | `add_floating_picture` |
| Bullet list | Wingdings `U+F0B2` + hanging indent (this is exactly how Word stores that bullet) | `symbol_run`, `set_para(hanging=)` |
| Blue underlined e-mail/URL | real hyperlink run | `add_hyperlink` |
| Two-column body text | table cells, or explicit tab stops on one line | `set_para(tabs=)` |
| Slanted parallelogram / rotated ribbon | straight band of the same colour and thickness, unless the client asks for the exact angle | approximation - say so in the report |

## Geometry conversion

PDF user-space units and Word's `Pt()` are both points at 72/inch, so measured
coordinates pass through unchanged for page-anchored objects:

* margins and picture placement: `Pt()` straight from the PDF
* table/indent widths: `Cm(pt / 28.3465)`
* table widths and cell margins are *twips* internally: `pt * 20`
* border weights: eighths of a point (`w:sz`) → 0.75pt = `6`
* character spacing: twentieths of a point → `pt * 20`

Page setup follows the PDF: A4 is `(21.0, 29.7)` cm. Margins should reproduce
the source's *text* box, which is not always symmetric (a designed resume can
have a 0.7cm right margin to make room for a margin bar).

## Paragraph rhythm

Sample the vertical distance between consecutive text rows in the report and use
it as the exact line height (Chinese business documents are usually 10.5pt text
on an 18pt rhythm). For the gap between blocks:

```
space_before ≈ (gap to previous row) − line_height
space_after  ≈ (gap to next row)     − line_height
```

The same rhythm applied uniformly reads better than copying every measured gap,
which is usually the same value plus rounding noise.

## Worked example - a two-page designed resume

A generic A4 two-page resume, rounded to whole points (`rect x0,top -> x1,bottom`),
and the Word construct each measurement turns into:

| Measurement from the PDF | Word construct |
| --- | --- |
| `rect 45,107 -> 575,108` fill `#4E7282` | heading paragraph, bottom border `sz=6`, colour `4E7282` |
| ribbon `30,92 -> 140,108` with a white 12pt title at x=46 | run shading on a 12pt white run, `\u3000` padding, `left_indent = -0.4cm` so the bar starts outside the text margin |
| a 2-glyph label whose glyphs sit 21pt apart (`姓` 50, `名` 71, `：` 92) | label run with `spacing_pt = 10.55` (2 glyphs spread over 3 cells), value in the next table cell |
| field columns start at x=50 and x=290 | columns `1.9 \| 4.82 \| 1.9 \| 9.93` cm |
| photo bbox `460,115 -> 540,228` | `add_floating_picture(..., 460, 115, 80, 113)` anchored in the *body-level* heading paragraph |
| three badges at `455,32 -> 548,55` | cropped image, same helper |
| `rect 41,92 -> 42,800` fill `#4E7282` | `page_border(sides=("left",), space=9)` |
| body rows every 18pt, ~10.5pt text | `line_pt=18.0`, body size 10.5pt |
| heading gaps of 22-27pt | `before=6, after=7` on the heading paragraph |

Result: 2 pages, page 2 starting at the same section, identical text, pictures
within 0.1pt - reported by `verify_conversion.py`.

## Pitfalls

* **East Asian font is not inherited.** Set `w:rFonts/@w:eastAsia` on every run
  (`style_run` does). Otherwise Word uses the theme's East Asian font and the
  document silently renders in a different typeface.
* **Word merges adjacent tables.** Two tables with no paragraph between them
  become one table. Insert a 1pt spacer paragraph.
* **A floating picture anchored inside a table cell is offset.** Word adds the
  cell's origin to `posOffset`, shifting the picture by the page margins
  (measured: exactly the top/left margin). Anchor floating pictures in a
  body-level paragraph.
* **Exact line spacing clips large glyphs.** A 36pt title needs ~44pt of line;
  use `row_height(..., rule="atLeast")` or give the title line enough room.
* **Shading rectangles follow the run's font size.** Padding spaces at a smaller
  size make a stepped bar; use the same size as the text.
* **Trailing CJK punctuation may hang past the right margin.** Word's default
  允许标点溢出边界 lets `、`/`，` overhang, so lines can end a few points beyond
  the text box. That is normal Word behaviour, not a bug - do not chase an exact
  right edge, and do not treat it as overflow when comparing renders. Add
  `w:overflowPunct @w:val="0"` to the paragraph properties only if the client
  needs a hard right edge.
* **Empty cells collapse.** To reproduce a fixed-height band, set the row height
  (`rule="exact"`) *and* give the cell a tiny explicit paragraph.
* **`w:pBdr`, `w:shd`, `w:spacing` and friends are not modelled by python-docx.**
  Insert them in schema order (`insert_ordered`) or Word may ignore the file.
* **Don't invent content.** If the source has a typo, a dangling `、` or a
  truncated sentence, keep it and mention it - the user asked for a conversion.
