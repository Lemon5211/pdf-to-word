"""Extract artwork from a PDF as high-resolution PNGs.

Photos, logos and icon badges that sit *behind* a designed layout have no
usable text equivalent, so they are lifted out of the PDF and re-inserted as
pictures at the same page coordinates.  Rendering the crop (instead of pulling
the embedded stream) also captures the clipping and blending the PDF applies.

Examples
--------
# every embedded image on page 1, named by position
python crop_pdf_image.py resume.pdf --page 1 --auto --out-dir work

# an arbitrary region, e.g. the icon badges seen in the vectors listing
python crop_pdf_image.py resume.pdf --page 1 --bbox 400,30,500,60 \
    --out work/badges.png --dpi 600
"""

import argparse
import os
import sys


def parse_box(text):
    parts = [float(value) for value in text.split(",")]
    if len(parts) != 4:
        raise argparse.ArgumentTypeError("bbox must be x0,top,x1,bottom")
    return parts


def crop(pdf_path, page_number, box, out_path, dpi=600, pad=0.0):
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number - 1]
        region = (box[0] - pad, box[1] - pad, box[2] + pad, box[3] + pad)
        image = page.crop(region).to_image(resolution=dpi)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    image.save(out_path)
    return out_path


def auto_crop(pdf_path, page_number, out_dir, dpi=600):
    import pdfplumber

    written = []
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number - 1]
        for index, image in enumerate(page.images, 1):
            box = (image["x0"], image["top"], image["x1"], image["bottom"])
            name = (f"page{page_number}_img{index}"
                    f"_x{int(box[0])}_y{int(box[1])}.png")
            written.append(crop(pdf_path, page_number, box,
                                os.path.join(out_dir, name), dpi=dpi))
    return written


def main(argv=None):
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pdf")
    parser.add_argument("--page", type=int, default=1)
    parser.add_argument("--bbox", type=parse_box, help="x0,top,x1,bottom in points")
    parser.add_argument("--out", help="output PNG path (with --bbox)")
    parser.add_argument("--auto", action="store_true",
                        help="crop every embedded image on the page")
    parser.add_argument("--out-dir", help="directory for --auto")
    parser.add_argument("--dpi", type=int, default=600)
    parser.add_argument("--pad", type=float, default=0.0,
                        help="grow the region by this many points")
    args = parser.parse_args(argv)

    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    if args.auto:
        target = args.out_dir or "."
        for path in auto_crop(args.pdf, args.page, target, dpi=args.dpi):
            print(f"[crop] {path}")
        return 0
    if not (args.bbox and args.out):
        parser.error("provide --bbox and --out, or --auto")
    path = crop(args.pdf, args.page, args.bbox, args.out, dpi=args.dpi, pad=args.pad)
    print(f"[crop] {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
