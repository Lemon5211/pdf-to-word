"""Runnable skeleton for a per-document PDF -> DOCX build script.

Copy this file into the task's working directory, replace the CONTENT block
with the text measured by ``inspect_pdf.py``, and adjust the geometry constants
to that PDF.  All layout primitives come from ``ooxml.py``.

The recipe below covers the constructs that show up in most designed PDFs
(resume templates, one-pagers, reports):

    1 header block   -> 1x2 borderless table: big title | stacked strapline
    2 accent band    -> 1x2 shaded table row (two-tone rule under the header)
    3 section bar    -> run shading + full-width paragraph bottom border
    4 form grid      -> 4-column table, 2-character labels spread over 3 cells
    5 picture slot   -> floating picture at absolute page coordinates
    6 body/bullets   -> exact line spacing, Wingdings bullet with hanging indent

Running it produces a demo document, which also smoke-tests the helpers:

    python build_template.py --out demo.docx [--photo portrait.png]
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ooxml import (  # noqa: E402
    add_floating_picture, add_hyperlink, add_run, fixed_table, new_document,
    page_border, para_border, row_height, set_para, shade_cell, spacer,
    symbol_run,
)

# --------------------------------------------------------------- CONTENT
PRIMARY = "4E7282"
GOLD = "C19F67"
INK = "414141"
CN = "微软雅黑"
LATIN = "Microsoft YaHei"

TITLE = "个人简历"
STRAPLINE = "细心从每一个小细节开始。"
SUBTITLE = "Personal resume"

FIELDS = [
    [("姓名", "张小明", False), ("政治面貌", "中共党员", True)],
    [("出生年月", "2000.01.01", False), ("身高", "170cm", False)],
    [("电话", "13800000000", False), ("毕业院校", "某某大学", True)],
    [("邮箱", ("zhangxm@example.com", "mailto:zhangxm@example.com"), False),
     ("学历", "硕士研究生", True)],
]

EDUCATION = [
    ("研究生 ", False), ("某某大学 计算机科学与技术（硕士）", True),
    ("  GPA：3.9/4.0", True),
]
EDUCATION_COURSES = "主修课程：机器学习、分布式系统、计算机视觉、编译原理。"

SELF_INTRO = "本人性格开朗，做事细致，具备扎实的工程能力与良好的团队协作意识。"
SELF_BULLETS = [
    ("有扎实的专业能力", "：熟悉从需求分析到上线运维的完整流程。"),
    ("有优秀的学习能力", "：能快速吸收新知识并落地到实际项目中。"),
]

BULLET_MARK = "\uf0b2"  # Wingdings: exactly how Word stores this bullet glyph

# geometry read off the source PDF, in points from the page's top-left corner
PHOTO_BOX = (460, 115, 540, 228)
PAGE = (21.0, 29.7)                # A4, cm
MARGINS = (1.1, 0.7, 1.45, 1.75)   # top, right, bottom, left (cm)
CONTENT_CM = PAGE[0] - MARGINS[1] - MARGINS[3]


def placeholder_photo(path, width=340, height=480, fill=(214, 222, 228)):
    """Stand-in for a portrait/logo cropped out of the source PDF."""
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (width, height), fill)
    draw = ImageDraw.Draw(image)
    draw.ellipse([width * 0.3, height * 0.12, width * 0.7, height * 0.46],
                 fill=(150, 168, 178))
    draw.ellipse([width * 0.16, height * 0.5, width * 0.84, height * 1.05],
                 fill=(150, 168, 178))
    image.save(path)
    return path


def section_bar(doc, text):
    """White heading on a coloured bar, with the rule that runs under it."""
    paragraph = doc.add_paragraph()
    set_para(paragraph, line=18, before=6, after=7, left=-0.4,
             keep_with_next=True)
    for chunk in ("\u3000", text, "\u3000"):
        add_run(paragraph, chunk, size=12, color="FFFFFF", shading=PRIMARY,
                font=LATIN, east_asian=CN)
    para_border(paragraph, side="bottom", color=PRIMARY, sz=6)
    return paragraph


def body(doc, runs):
    paragraph = doc.add_paragraph()
    set_para(paragraph, line=18)
    for text, bold in runs:
        add_run(paragraph, text, size=10.5, bold=bold, font=LATIN, east_asian=CN)
    return paragraph


def bullets(doc, items):
    for head, tail in items:
        paragraph = doc.add_paragraph()
        set_para(paragraph, line=18, left=0.33, hanging=0.33)
        symbol_run(paragraph, BULLET_MARK, size=10.5, color=INK)
        add_run(paragraph, head, size=10.5, bold=True, font=LATIN, east_asian=CN)
        add_run(paragraph, tail, size=10.5, font=LATIN, east_asian=CN)


def field_grid(doc, rows, label_cm=1.9, value_cm=4.82):
    """Form grid of label/value pairs; 2-character labels are spread to 3 cells."""
    widths = [label_cm, value_cm, label_cm, CONTENT_CM - label_cm * 2 - value_cm]
    table = fixed_table(doc.add_table(rows=len(rows), cols=4), widths)
    for index, pair in enumerate(rows):
        row = table.rows[index]
        row_height(row, 18)
        for column, (label, value, bold) in enumerate(pair):
            cell = row.cells[column * 2]
            set_para(cell.paragraphs[0], line=18)
            spacing = 10.55 if len(label) == 2 else None
            add_run(cell.paragraphs[0], label, size=10.5, spacing_pt=spacing,
                    font=LATIN, east_asian=CN)
            add_run(cell.paragraphs[0], "：", size=10.5, font=LATIN, east_asian=CN)
            value_cell = row.cells[column * 2 + 1]
            set_para(value_cell.paragraphs[0], line=18)
            if isinstance(value, tuple):  # (text, url)
                add_hyperlink(value_cell.paragraphs[0], value[1], value[0],
                              font=LATIN, east_asian=CN)
            else:
                add_run(value_cell.paragraphs[0], value, size=10.5, bold=bold,
                        font=LATIN, east_asian=CN)
    return table


def build(out_path, photo_path=None):
    doc = new_document(page_cm=PAGE, margins_cm=MARGINS, font=LATIN,
                       east_asian=CN, size=10.5, color=INK, line_pt=18.0)
    page_border(doc.sections[0], sides=("left",), color=PRIMARY, sz=6, space=9)

    # 1 - header: the row height pins the block to the source PDF's band
    header = fixed_table(doc.add_table(rows=1, cols=2), [5.74, 12.81])
    row_height(header.rows[0], 46, rule="atLeast")
    paragraph = header.cell(0, 0).paragraphs[0]
    set_para(paragraph, line=44)
    add_run(paragraph, TITLE, size=36, color=PRIMARY, font=LATIN, east_asian=CN)
    cell = header.cell(0, 1)
    set_para(cell.paragraphs[0], line=16)
    add_run(cell.paragraphs[0], STRAPLINE, size=10.5, color=PRIMARY, font=LATIN,
            east_asian=CN)
    strapline2 = cell.add_paragraph()
    set_para(strapline2, line=24)
    add_run(strapline2, SUBTITLE, size=18, color=PRIMARY, font=LATIN, east_asian=CN)

    spacer(doc)  # without an intervening paragraph Word merges the two tables

    # 2 - two-tone accent band
    divider = fixed_table(doc.add_table(rows=1, cols=2), [10.39, 8.16])
    row_height(divider.rows[0], 14)
    for index, fill in enumerate((PRIMARY, GOLD)):
        cell = divider.cell(0, index)
        shade_cell(cell, fill)
        set_para(cell.paragraphs[0], line=1)
        add_run(cell.paragraphs[0], "", size=1)

    # 3 + 4 - section bar and the form grid beneath it
    anchor = section_bar(doc, "基本信息")
    field_grid(doc, FIELDS)

    # 5 - floating picture, positioned straight from the PDF's image bbox.
    #     The anchor must be a body-level paragraph (not inside a table cell).
    if photo_path and os.path.exists(photo_path):
        x0, top, x1, bottom = PHOTO_BOX
        add_floating_picture(anchor, photo_path, x0, top, x1 - x0, bottom - top)

    # 6 - body copy and bullets
    section_bar(doc, "教育背景")
    body(doc, EDUCATION)
    body(doc, [(EDUCATION_COURSES, False)])
    section_bar(doc, "自我评价")
    body(doc, [(SELF_INTRO, False)])
    bullets(doc, SELF_BULLETS)

    doc.save(out_path)
    return out_path


def main(argv=None):
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--out", default="demo.docx")
    parser.add_argument("--photo", help="image for the floating picture slot")
    args = parser.parse_args(argv)

    photo = args.photo
    if not photo:
        photo = os.path.join(os.path.dirname(os.path.abspath(args.out)) or ".",
                             "demo_portrait.png")
        placeholder_photo(photo)
    build(args.out, photo)
    print(f"[build] wrote {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
