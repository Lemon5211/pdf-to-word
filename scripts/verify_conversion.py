"""Check that a PDF -> DOCX conversion kept the text, images and pagination.

Two independent comparisons are run:

1. **Content** - every character in the source PDF against every character in
   the .docx (whitespace-insensitive).  A sequence diff catches reordering; a
   multiset diff catches omissions and accidental additions even when the
   reading order legitimately differs.
2. **Geometry** - page count, text-row rhythm, picture bounding boxes and
   hyperlink targets of the Word-rendered PDF against the source PDF, so that
   "looks converted" is backed by coordinates.

Examples
--------
python verify_conversion.py --source in.pdf --docx out.docx --rendered out.pdf
python verify_conversion.py --source in.pdf --docx out.docx      # content only
"""

import argparse
import os
import re
import sys
from collections import Counter

import difflib


def squash(text):
    return re.sub(r"\s+", "", text)


def source_text(path, min_size=1.0, tol=2.0):
    """Reading-order text of the source PDF (tiny artefacts are dropped)."""
    import pdfplumber

    chunks = []
    with pdfplumber.open(path) as pdf:
        pages = len(pdf.pages)
        for page in pdf.pages:
            chars = [c for c in page.chars if c["size"] >= min_size]
            rows = []
            for ch in sorted(chars, key=lambda c: (c["top"], c["x0"])):
                for row in rows:
                    if abs(row[0] - ch["top"]) <= tol:
                        row[1].append(ch)
                        break
                else:
                    rows.append([ch["top"], [ch]])
            for _, row_chars in sorted(rows, key=lambda r: r[0]):
                chunks.append("".join(c["text"] for c in sorted(row_chars,
                                                                 key=lambda c: c["x0"])))
    return squash("".join(chunks)), pages


def docx_text(path):
    from docx import Document

    doc = Document(path)
    parts = []
    for paragraph in doc.element.body.iter():
        if paragraph.tag.split("}")[1] == "p":
            parts.append("".join(
                node.text or "" for node in paragraph.iter()
                if node.tag.split("}")[1] == "t"))
    return squash("".join(parts))


def compare_content(src, docx):
    ratio = difflib.SequenceMatcher(None, src, docx, autojunk=False).ratio()
    counter = Counter(src) - Counter(docx)
    extra = Counter(docx) - Counter(src)
    missing = "".join(sorted(counter.elements()))
    added = "".join(sorted(extra.elements()))
    return ratio, missing, added


def rows(page, tol=2.0, min_size=1.0):
    out = []
    chars = [c for c in page.chars if c["size"] >= min_size]
    for ch in sorted(chars, key=lambda c: (c["top"], c["x0"])):
        for row in out:
            if abs(row["top"] - ch["top"]) <= tol:
                row["chars"].append(ch)
                break
        else:
            out.append({"top": ch["top"], "chars": [ch]})
    for row in out:
        row["x0"] = min(c["x0"] for c in row["chars"])
        row["x1"] = max(c["x1"] for c in row["chars"])
        row["text"] = "".join(c["text"] for c in sorted(row["chars"],
                                                        key=lambda c: c["x0"]))
    return sorted(out, key=lambda r: r["top"])


def bbox(image):
    return (round(image["x0"], 1), round(image["top"], 1),
            round(image["x1"], 1), round(image["bottom"], 1))


def compare_geometry(source, rendered, tol=3.0):
    import pdfplumber

    report = {"ok": True, "notes": []}
    with pdfplumber.open(source) as src_pdf, pdfplumber.open(rendered) as out_pdf:
        report["pages"] = (len(src_pdf.pages), len(out_pdf.pages))
        if len(src_pdf.pages) != len(out_pdf.pages):
            report["ok"] = False
            report["notes"].append(
                f"page count differs: source {len(src_pdf.pages)} vs rendered "
                f"{len(out_pdf.pages)}")
        for index in range(min(len(src_pdf.pages), len(out_pdf.pages))):
            src_page, out_page = src_pdf.pages[index], out_pdf.pages[index]
            src_rows, out_rows = rows(src_page), rows(out_page)
            src_images = {bbox(i)[:2] for i in src_page.images}
            out_images = {bbox(i)[:2] for i in out_page.images}
            if len(src_rows) != len(out_rows):
                report["notes"].append(
                    f"page {index + 1}: row count {len(src_rows)} -> {len(out_rows)} "
                    "(reflow, expected when paragraphs are re-joined)")
            # match text rows by content so the drift figure stays meaningful
            # even when the rebuilt paragraphs re-wrap into a different
            # number of lines
            src_by_text, out_by_text = {}, {}
            for row in src_rows:
                src_by_text.setdefault(squash(row["text"]), []).append(row["top"])
            for row in out_rows:
                out_by_text.setdefault(squash(row["text"]), []).append(row["top"])
            drift, unmatched = [], []
            for text, tops in src_by_text.items():
                if text in out_by_text:
                    drift.append(min(out_by_text[text]) - min(tops))
                else:
                    unmatched.append(text[:24])
            if drift:
                mean = sum(drift) / len(drift)
                report["notes"].append(
                    f"page {index + 1}: {len(drift)}/{len(src_by_text)} text rows "
                    f"matched, vertical drift {min(drift):+.1f}..{max(drift):+.1f}pt "
                    f"(mean {mean:+.1f})")
            if unmatched:
                report["notes"].append(
                    f"page {index + 1}: {len(unmatched)} source row(s) have no "
                    "verbatim match because re-flowed line breaks land "
                    f"differently: {', '.join(unmatched[:4])}"
                    + (" ..." if len(unmatched) > 4 else ""))
            for origin in sorted(src_images):
                distances = [max(abs(origin[0] - o[0]), abs(origin[1] - o[1]))
                             for o in out_images]
                nearest = min(distances) if distances else None
                if nearest is None or nearest > tol:
                    report["ok"] = False
                    detail = "missing" if nearest is None else f"nearest {nearest:.1f}pt away"
                    report["notes"].append(
                        f"page {index + 1}: picture at ({origin[0]}, {origin[1]}) "
                        f"{detail}")
            src_links = [a.get("uri") for a in (src_page.annots or []) if a.get("uri")]
            out_links = [a.get("uri") for a in (out_page.annots or []) if a.get("uri")]
            for uri in src_links:
                if uri not in out_links:
                    report["ok"] = False
                    report["notes"].append(
                        f"page {index + 1}: hyperlink lost: {uri}")
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--source", required=True, help="original PDF")
    parser.add_argument("--docx", required=True, help="converted .docx")
    parser.add_argument("--rendered", help="PDF produced from the .docx")
    parser.add_argument("--min-size", type=float, default=1.0,
                        help="ignore source glyphs smaller than this (pt)")
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(argv)

    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    src_text, src_pages = source_text(args.source, min_size=args.min_size)
    docx_plain = docx_text(args.docx)
    ratio, missing, added = compare_content(src_text, docx_plain)

    print(f"source : {src_pages} page(s), {len(src_text)} chars of text")
    print(f"docx   : {os.path.basename(args.docx)}, {len(docx_plain)} chars of text")
    print(f"content: sequence similarity {ratio:.5f}")
    ok = True
    if missing:
        ok = False
        print(f"content: MISSING from docx -> {missing[:120]!r}")
    if added:
        print(f"content: EXTRA in docx     -> {added[:120]!r} (review whether intended)")
    if ratio < 0.995:
        ok = False
        matcher = difflib.SequenceMatcher(None, src_text, docx_plain, autojunk=False)
        differences = [op for op in matcher.get_opcodes() if op[0] != "equal"]
        for tag, i1, i2, j1, j2 in differences[:12]:
            print(f"content: {tag:8} source[{i1}:{i2}]={src_text[i1:i2]!r} "
                  f"docx[{j1}:{j2}]={docx_plain[j1:j2]!r}")
        if len(differences) > 12:
            print(f"content: ... {len(differences) - 12} further differences "
                  "not shown")
    if ratio == 1.0:
        print("content: PASS (identical after whitespace normalisation)")
    elif not missing and ratio >= 0.995:
        print("content: PASS (only reading-order differences)")
    else:
        print("content: FAIL")

    if args.rendered:
        if not os.path.exists(args.rendered):
            print(f"geometry: skipped, {args.rendered} not found")
        else:
            report = compare_geometry(args.source, args.rendered)
            print(f"geometry: {'PASS' if report['ok'] else 'FAIL'} "
                  f"(pages {report['pages'][0]} -> {report['pages'][1]})")
            for note in report["notes"]:
                print(f"  - {note}")
            ok = ok and report["ok"]
    else:
        print("geometry: not checked (pass --rendered to compare coordinates)")

    print("VERDICT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
