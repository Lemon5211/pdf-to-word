"""Describe a PDF's text, vector design and images so it can be rebuilt in Word.

Prints a readable report (and optionally a machine-readable JSON) containing,
for every page:

* page size, text-layer presence, image and annotation inventory
* text rows grouped by baseline, with per-run bold flag, font, size, colour and
  x-range -- enough to reconstruct paragraphs, tables and label columns
* vector shapes (rects/curves/lines) with fill colour and bbox -- the source of
  ribbons, rules, bands and decorative bars

It can also rasterise pages (``--render``) or print them as ASCII art
(``--ascii``) when a visual read of the design helps.

Examples
--------
python inspect_pdf.py resume.pdf --out work --render
python inspect_pdf.py resume.pdf --ascii --crop 40,80,560,250
"""

import argparse
import json
import os
import sys

import pdfplumber


def to_hex(color):
    """Normalise pdfplumber colour tuples (rgb / gray / cmyk) to '#RRGGBB'."""
    if color is None:
        return None
    if isinstance(color, (int, float)):
        value = int(round(color * 255))
        return f"#{value:02X}{value:02X}{value:02X}"
    parts = list(color)
    if len(parts) == 1:
        value = int(round(parts[0] * 255))
        return f"#{value:02X}{value:02X}{value:02X}"
    if len(parts) == 3:
        return "#" + "".join(f"{int(round(max(0, min(1, p)) * 255)):02X}" for p in parts)
    if len(parts) == 4:  # cmyk
        c, m, y, k = [max(0.0, min(1.0, p)) for p in parts]
        return "#" + "".join(
            f"{int(round(255 * (1 - min(1, v + k)))):02X}" for v in (c, m, y)
        )
    return str(color)


def is_bold(fontname):
    lowered = (fontname or "").lower()
    return "bold" in lowered or "black" in lowered or "heavy" in lowered


def group_rows(chars, tol):
    rows = []
    for ch in sorted(chars, key=lambda c: (c["top"], c["x0"])):
        for row in rows:
            if abs(row["top"] - ch["top"]) <= tol:
                row["chars"].append(ch)
                break
        else:
            rows.append({"top": ch["top"], "chars": [ch]})
    return sorted(rows, key=lambda r: r["top"])


def row_runs(row):
    """Collapse a row into consecutive runs sharing bold/size/colour/font."""
    chars = sorted(row["chars"], key=lambda c: c["x0"])
    runs = []
    for ch in chars:
        key = (is_bold(ch["fontname"]), round(ch["size"], 2), str(ch["fontname"]),
               to_hex(ch.get("non_stroking_color")))
        if runs and runs[-1]["key"] == key:
            runs[-1]["text"] += ch["text"]
            runs[-1]["x1"] = ch["x1"]
        else:
            runs.append({"key": key, "text": ch["text"], "x0": ch["x0"],
                         "x1": ch["x1"]})
    return runs


def describe(path, tol=2.0):
    data = {"path": os.path.abspath(path), "pages": []}
    with pdfplumber.open(path) as pdf:
        for index, page in enumerate(pdf.pages, 1):
            rows = group_rows(page.chars, tol)
            entry = {
                "index": index,
                "width": round(page.width, 1),
                "height": round(page.height, 1),
                "chars": len(page.chars),
                "has_text_layer": bool(page.chars),
                "rows": [],
                "images": [],
                "annots": [],
                "shapes": [],
            }
            for row in rows:
                runs = row_runs(row)
                entry["rows"].append(
                    {
                        "top": round(row["top"], 1),
                        "x0": round(runs[0]["x0"], 1),
                        "x1": round(runs[-1]["x1"], 1),
                        "size": max(r["key"][1] for r in runs),
                        "text": "".join(r["text"] for r in runs),
                        "runs": [
                            {
                                "text": r["text"],
                                "bold": r["key"][0],
                                "size": r["key"][1],
                                "font": r["key"][2],
                                "color": r["key"][3],
                                "x0": round(r["x0"], 1),
                                "x1": round(r["x1"], 1),
                            }
                            for r in runs
                        ],
                    }
                )
            for image in page.images:
                entry["images"].append(
                    {
                        "x0": round(image["x0"], 1),
                        "top": round(image["top"], 1),
                        "x1": round(image["x1"], 1),
                        "bottom": round(image["bottom"], 1),
                        "srcsize": image.get("srcsize"),
                    }
                )
            for annot in page.annots or []:
                entry["annots"].append(
                    {
                        "type": str(annot.get("object_type") or annot.get("data", {}).get("Subtype")),
                        "uri": annot.get("uri"),
                        "x0": round(annot["x0"], 1),
                        "top": round(annot["top"], 1),
                        "x1": round(annot["x1"], 1),
                        "bottom": round(annot["bottom"], 1),
                    }
                )
            for kind, shapes in (("rect", page.rects), ("curve", page.curves),
                                 ("line", page.lines)):
                for shape in shapes:
                    width = shape["x1"] - shape["x0"]
                    height = shape["bottom"] - shape["top"]
                    fill = to_hex(shape.get("non_stroking_color"))
                    if width < 1.0 and height < 1.0:
                        continue  # line-cap artefacts
                    if fill == "#FFFFFF" and (width >= 0.9 * page.width
                                              or height >= 0.9 * page.height):
                        continue  # full-page white background layer
                    if not fill and not shape.get("stroking_color"):
                        continue
                    entry["shapes"].append(
                        {
                            "kind": kind,
                            "x0": round(shape["x0"], 1),
                            "top": round(shape["top"], 1),
                            "x1": round(shape["x1"], 1),
                            "bottom": round(shape["bottom"], 1),
                            "w": round(width, 1),
                            "h": round(height, 1),
                            "fill": fill,
                            "stroke": to_hex(shape.get("stroking_color")),
                            "linewidth": shape.get("linewidth"),
                        }
                    )
            data["pages"].append(entry)
    return data


def print_report(data, stream=sys.stdout, max_shapes=200):
    for page in data["pages"]:
        print("=" * 78, file=stream)
        print(
            f"PAGE {page['index']}  {page['width']}x{page['height']}pt  "
            f"chars={page['chars']}  rows={len(page['rows'])}",
            file=stream,
        )
        if not page["has_text_layer"]:
            print("  !! no text layer -- scanned page, OCR is required", file=stream)
        for image in page["images"]:
            print(f"  image bbox=({image['x0']}, {image['top']}, {image['x1']}, "
                  f"{image['bottom']}) px={image['srcsize']}", file=stream)
        for annot in page["annots"]:
            print(f"  annot {annot['type']} uri={annot['uri']} "
                  f"bbox=({annot['x0']}, {annot['top']}, {annot['x1']}, {annot['bottom']})",
                  file=stream)
        print("  -- vectors --", file=stream)
        shown = page["shapes"][:max_shapes]
        for shape in shown:
            print(
                f"    {shape['kind']:5} x0={shape['x0']:7.1f} top={shape['top']:7.1f} "
                f"x1={shape['x1']:7.1f} bottom={shape['bottom']:7.1f} "
                f"w={shape['w']:6.1f} h={shape['h']:5.1f} fill={shape['fill']} "
                f"stroke={shape['stroke']}",
                file=stream,
            )
        if len(page["shapes"]) > len(shown):
            print(f"    ... {len(page['shapes']) - len(shown)} more shapes "
                  "(mostly table rules; raise --max-shapes to see them all)",
                  file=stream)
        print("  -- text --", file=stream)
        for row in page["rows"]:
            print(f"    top={row['top']:7.1f} x0={row['x0']:7.1f} x1={row['x1']:7.1f} "
                  f"sz={row['size']:5.2f}", file=stream)
            for run in row["runs"]:
                print(
                    f"        [{run['x0']:7.1f}-{run['x1']:7.1f}] "
                    f"bold={str(run['bold']):5} sz={run['size']:5.2f} "
                    f"{run['color']} {run['font']:32} | {run['text']}",
                    file=stream,
                )


def ascii_art(pdf_path, page_index, box, cols, resolution=200):
    from PIL import Image

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_index - 1]
        region = (box[0], box[1], box[2], box[3]) if box else (
            0, 0, page.width, page.height)
        image = page.crop(region).to_image(resolution=resolution).original
    image = image.convert("L")
    width, height = image.size
    rows = max(1, round(cols * height / width / 2.1))
    small = image.resize((cols, rows), Image.LANCZOS)
    pixels = small.load()
    ramp = " .:-=+*#%@"
    print(f"### ascii page {page_index} box={region} img={width}x{height}")
    for y in range(rows):
        print("|" + "".join(
            ramp[min(9, (255 - pixels[x, y]) * 10 // 256)] for x in range(cols)
        ) + "|")


def render_pages(pdf_path, out_dir, dpi=200):
    os.makedirs(out_dir, exist_ok=True)
    written = []
    with pdfplumber.open(pdf_path) as pdf:
        for index, page in enumerate(pdf.pages, 1):
            path = os.path.join(out_dir, f"page{index}.png")
            page.to_image(resolution=dpi).save(path)
            written.append(path)
    return written


def parse_box(text):
    parts = [float(p) for p in text.split(",")]
    if len(parts) != 4:
        raise argparse.ArgumentTypeError("crop must be x0,top,x1,bottom")
    return parts


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pdf")
    parser.add_argument("--out", help="directory for report.txt / layout.json / renders")
    parser.add_argument("--json", action="store_true", help="also dump layout.json")
    parser.add_argument("--render", action="store_true", help="write page PNGs")
    parser.add_argument("--dpi", type=int, default=200)
    parser.add_argument("--ascii", action="store_true", help="print ASCII art")
    parser.add_argument("--page", type=int, default=1, help="page for --ascii")
    parser.add_argument("--crop", type=parse_box, help="ascii/render window x0,top,x1,bottom")
    parser.add_argument("--cols", type=int, default=112, help="ascii art width")
    parser.add_argument("--tol", type=float, default=2.0, help="row grouping tolerance (pt)")
    parser.add_argument("--max-shapes", type=int, default=200,
                        help="cap the vector shapes listed per page")
    parser.add_argument("--quiet", action="store_true", help="do not print the report")
    args = parser.parse_args(argv)

    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    data = describe(args.pdf, tol=args.tol)
    if not args.quiet:
        print_report(data, max_shapes=args.max_shapes)
    if args.out:
        os.makedirs(args.out, exist_ok=True)
        with open(os.path.join(args.out, "report.txt"), "w", encoding="utf-8") as fh:
            print_report(data, stream=fh, max_shapes=args.max_shapes)
        if args.json:
            with open(os.path.join(args.out, "layout.json"), "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=1)
        print(f"[inspect] wrote {args.out}/report.txt")
    if args.render:
        target = args.out or "."
        for path in render_pages(args.pdf, target, dpi=args.dpi):
            print(f"[inspect] rendered {path}")
    if args.ascii:
        ascii_art(args.pdf, args.page, args.crop, args.cols)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
