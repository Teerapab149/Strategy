from __future__ import annotations

import re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(r"E:\strategy\research")
SOURCE = ROOT / "report-source.md"
OUTPUT = ROOT / "PG_Strategic_Research_Dossier_2026.docx"
CHART = ROOT / "pg_financial_trend.png"

# standard_business_brief preset with a named Thai-font override.
FONT = "Leelawadee UI"
BLUE = "2E74B5"
DARK_BLUE = "1F4D78"
NAVY = "0B2545"
MUTED = "667085"
LIGHT = "F2F4F7"
PALE_BLUE = "EAF2F8"
PALE_GOLD = "FFF7E0"
RED = "9B1C1C"
GREEN = "1E6B52"


def set_font(run, size=None, bold=None, color=None, italic=None, name=FONT):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_table_borders(table, color="D0D5DD", size="4"):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = borders.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            borders.append(tag)
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), size)
        tag.set(qn("w:space"), "0")
        tag.set(qn("w:color"), color)


def set_table_geometry(table, widths_dxa):
    total = sum(widths_dxa)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl_pr = table._tbl.tblPr

    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")

    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), "120")
    tbl_ind.set(qn("w:type"), "dxa")

    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths_dxa:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)

    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            width = widths_dxa[min(idx, len(widths_dxa) - 1)]
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(width))
            tc_w.set(qn("w:type"), "dxa")
            cell.width = Inches(width / 1440)
            set_cell_margins(cell)


def choose_widths(rows):
    cols = len(rows[0])
    # Content-aware but deterministic; totals exactly 9360 DXA.
    if cols == 2:
        return [2700, 6660]
    if cols == 3:
        return [3100, 1700, 4560]
    if cols == 4:
        return [3300, 1400, 1400, 3260]
    if cols == 5:
        return [3300, 1050, 1050, 1350, 2610]
    if cols == 6:
        return [2700, 900, 900, 900, 900, 3060]
    return [9360 // cols] * (cols - 1) + [9360 - (9360 // cols) * (cols - 1)]


def parse_inline(paragraph, text, base_size=11, base_color=None):
    # Markdown bold + URLs. Source tags stay human-readable.
    token_re = re.compile(r"(\*\*.*?\*\*|https?://\S+)")
    cursor = 0
    for match in token_re.finditer(text):
        if match.start() > cursor:
            run = paragraph.add_run(text[cursor:match.start()])
            set_font(run, size=base_size, color=base_color)
        token = match.group(0)
        if token.startswith("**"):
            run = paragraph.add_run(token[2:-2])
            set_font(run, size=base_size, bold=True, color=base_color)
        else:
            add_hyperlink(paragraph, token.rstrip(".,)"), token.rstrip(".,)"), base_size)
            tail = token[len(token.rstrip(".,)")):]
            if tail:
                run = paragraph.add_run(tail)
                set_font(run, size=base_size, color=base_color)
        cursor = match.end()
    if cursor < len(text):
        run = paragraph.add_run(text[cursor:])
        set_font(run, size=base_size, color=base_color)


def add_hyperlink(paragraph, text, url, size=9):
    part = paragraph.part
    rid = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rid)
    run_el = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), BLUE)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    size_el = OxmlElement("w:sz")
    size_el.set(qn("w:val"), str(int(size * 2)))
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), FONT)
    fonts.set(qn("w:hAnsi"), FONT)
    fonts.set(qn("w:eastAsia"), FONT)
    rpr.extend([fonts, color, underline, size_el])
    run_el.append(rpr)
    text_el = OxmlElement("w:t")
    text_el.text = text
    run_el.append(text_el)
    hyperlink.append(run_el)
    paragraph._p.append(hyperlink)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("หน้า ")
    set_font(run, size=8.5, color=MUTED)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    run_el = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), MUTED)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "17")
    rpr.extend([color, size])
    run_el.append(rpr)
    text_el = OxmlElement("w:t")
    text_el.text = "1"
    run_el.append(text_el)
    fld.append(run_el)
    paragraph._p.append(fld)


def add_toc(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run("สารบัญ")
    set_font(run, size=16, bold=True, color=BLUE)
    toc_p = doc.add_paragraph()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), 'TOC \\o "1-3" \\h \\z \\u')
    run_el = OxmlElement("w:r")
    text_el = OxmlElement("w:t")
    text_el.text = "คลิกขวาและเลือก Update Field หากสารบัญไม่อัปเดตอัตโนมัติ"
    run_el.append(text_el)
    fld.append(run_el)
    toc_p._p.append(fld)


def make_chart():
    years = ["2023", "2024", "2025"]
    sales = [703.52, 773.93, 605.20]
    net = [25.99, 1.96, -5.56]
    core = [-27.98, -24.11, -48.56]
    width, height = 1500, 700
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    font_path = Path(r"C:\Windows\Fonts\calibri.ttf")
    bold_path = Path(r"C:\Windows\Fonts\calibrib.ttf")
    font = ImageFont.truetype(str(font_path), 28) if font_path.exists() else ImageFont.load_default()
    small = ImageFont.truetype(str(font_path), 24) if font_path.exists() else ImageFont.load_default()
    bold = ImageFont.truetype(str(bold_path), 34) if bold_path.exists() else font
    draw.text((width // 2, 35), "PG Sales and Profit Trend (THB million)", fill="#0B2545", font=bold, anchor="ma")
    left, right, top, bottom = 125, 1380, 115, 560
    draw.line((left, bottom, right, bottom), fill="#98A2B3", width=2)
    draw.line((left, top, left, bottom), fill="#98A2B3", width=2)
    for tick in range(0, 901, 150):
        y = bottom - (tick / 900) * (bottom - top)
        draw.line((left, y, right, y), fill="#EAECF0", width=1)
        draw.text((left - 18, y), str(tick), fill="#667085", font=small, anchor="rm")
    xs = [350, 750, 1150]
    for x, year, value in zip(xs, years, sales):
        bar_h = (value / 900) * (bottom - top)
        draw.rounded_rectangle((x - 78, bottom - bar_h, x + 78, bottom), radius=8, fill="#2E74B5")
        draw.text((x, bottom - bar_h - 12), f"{value:.0f}", fill="#0B2545", font=small, anchor="ms")
        draw.text((x, bottom + 24), year, fill="#344054", font=font, anchor="ma")
    # Profit/loss points are drawn against a secondary -70..50 scale.
    def py(v):
        return bottom - ((v + 70) / 120) * (bottom - top)
    for values, color, shape in ((net, "#1E6B52", "circle"), (core, "#9B1C1C", "square")):
        pts = [(x, py(v)) for x, v in zip(xs, values)]
        draw.line(pts, fill=color, width=5)
        for x, y in pts:
            if shape == "circle":
                draw.ellipse((x - 10, y - 10, x + 10, y + 10), fill=color)
            else:
                draw.rectangle((x - 10, y - 10, x + 10, y + 10), fill=color)
    draw.line((920, 635, 980, 635), fill="#1E6B52", width=5)
    draw.text((995, 635), "Net profit", fill="#344054", font=small, anchor="lm")
    draw.line((1130, 635, 1190, 635), fill="#9B1C1C", width=5)
    draw.text((1205, 635), "Approx. core result", fill="#344054", font=small, anchor="lm")
    img.save(CHART, quality=95)


def configure_styles(doc):
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1)
    sec.right_margin = Inches(1)
    sec.header_distance = Inches(0.492)
    sec.footer_distance = Inches(0.492)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT
    normal._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal.font.size = Pt(11)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.10

    heading_specs = {
        "Heading 1": (16, BLUE, 16, 8),
        "Heading 2": (13, BLUE, 12, 6),
        "Heading 3": (12, DARK_BLUE, 8, 4),
    }
    for name, (size, color, before, after) in heading_specs.items():
        style = styles[name]
        style.font.name = FONT
        style._element.rPr.rFonts.set(qn("w:ascii"), FONT)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    for name in ("List Bullet", "List Number"):
        style = styles[name]
        style.font.name = FONT
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style.font.size = Pt(11)
        style.paragraph_format.left_indent = Inches(0.5)
        style.paragraph_format.first_line_indent = Inches(-0.25)
        style.paragraph_format.space_after = Pt(8)
        style.paragraph_format.line_spacing = 1.167

    if "Table Source" not in styles:
        style = styles.add_style("Table Source", WD_STYLE_TYPE.PARAGRAPH)
    else:
        style = styles["Table Source"]
    style.font.name = FONT
    style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    style.font.size = Pt(8.5)
    style.font.color.rgb = RGBColor.from_string(MUTED)
    style.paragraph_format.space_before = Pt(4)
    style.paragraph_format.space_after = Pt(4)

    # Quiet running header/footer, no decorative rule.
    header = sec.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hr = hp.add_run("PG Strategic Research Dossier | ข้อมูล ณ 31 สิงหาคม 2569")
    set_font(hr, size=8.5, color=MUTED)
    footer = sec.footer
    fp = footer.paragraphs[0]
    add_page_number(fp)


def add_cover(doc):
    for _ in range(5):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(14)
    r = p.add_run("STRATEGIC RESEARCH DOSSIER")
    set_font(r, size=11, bold=True, color=BLUE)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("บริษัท ประชาอาภรณ์ จำกัด (มหาชน)")
    set_font(r, size=28, bold=True, color=NAVY)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("People’s Garment Public Company Limited (PG)")
    set_font(r, size=15, color=DARK_BLUE)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(40)
    r = p.add_run("Focused Turnaround + Functional Uniform & Workwear Solutions")
    set_font(r, size=12, italic=True, color=MUTED)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ฐานข้อมูลเพื่อปรับ EFAS • IFAS • SFAS • TOWS • Strategic Alternatives • Action Plan")
    set_font(r, size=10.5, color=NAVY)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(54)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("ข้อมูล ณ 31 สิงหาคม 2569")
    set_font(r, size=11, bold=True, color=NAVY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("จัดทำจากเอกสารบริษัท หน่วยงานรัฐ กฎระเบียบ EU และข้อมูลคู่แข่งสาธารณะ")
    set_font(r, size=9, color=MUTED)
    doc.add_page_break()


def add_callout(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.18)
    p.paragraph_format.right_indent = Inches(0.18)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.line_spacing = 1.15
    p_pr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), PALE_BLUE)
    p_pr.append(shd)
    parse_inline(p, text, base_size=11, base_color=NAVY)


def add_table(doc, rows):
    is_claim_ledger = bool(rows and rows[0] and rows[0][0].startswith("Claim"))
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    widths = choose_widths(rows)
    set_table_geometry(table, widths)
    set_table_borders(table)
    set_repeat_table_header(table.rows[0])
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            cell = table.cell(i, j)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 0.90 if is_claim_ledger else 1.05
            if j > 0 and re.fullmatch(r"[-+≥≤<>0-9.,%–—/() ]+", value.strip()):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if is_claim_ledger:
                table_font_size = 7.2
            else:
                table_font_size = 8.3 if len(rows[0]) >= 5 else 8.8
            parse_inline(p, value, base_size=table_font_size)
            if i == 0:
                set_cell_shading(cell, LIGHT)
                for run in p.runs:
                    run.bold = True
                    run.font.color.rgb = RGBColor.from_string(NAVY)
    if not is_claim_ledger:
        after = doc.add_paragraph(style="Table Source")
        after.add_run("ที่มา: การสังเคราะห์จากแหล่งข้อมูลที่ระบุในเนื้อหาและรายการอ้างอิง")
    return table


def parse_markdown(doc, text):
    lines = text.splitlines()
    i = 0
    seen_title = False
    inserted_chart = False
    while i < len(lines):
        raw = lines[i].rstrip()
        stripped = raw.strip()
        if not stripped:
            i += 1
            continue

        if stripped.startswith("# ") and not seen_title:
            # Cover already carries the source title.
            seen_title = True
            i += 1
            continue

        if stripped.startswith("### "):
            # Markdown level 3 is the first subsection level in this report
            # because level 2 is mapped to the report's chapter Heading 1.
            p = doc.add_paragraph(style="Heading 2")
            parse_inline(p, stripped[4:])
            i += 1
            continue
        if stripped.startswith("## "):
            title = stripped[3:]
            p = doc.add_paragraph(style="Heading 1")
            parse_inline(p, title, base_size=16, base_color=BLUE)
            if title.startswith("2.5") and not inserted_chart:
                # Chart follows its introductory text later, not immediately.
                inserted_chart = True
            i += 1
            continue
        if stripped.startswith("# "):
            p = doc.add_paragraph(style="Heading 1")
            parse_inline(p, stripped[2:], base_size=16, base_color=BLUE)
            i += 1
            continue

        if stripped.startswith("> "):
            add_callout(doc, stripped[2:])
            i += 1
            continue

        # Markdown table
        if stripped.startswith("|") and i + 1 < len(lines) and re.match(r"^\s*\|?\s*:?-+", lines[i + 1]):
            rows = []
            header = [x.strip() for x in stripped.strip("|").split("|")]
            rows.append(header)
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append([x.strip() for x in lines[i].strip().strip("|").split("|")])
                i += 1
            add_table(doc, rows)
            # Insert the financial chart after the first 9-row diagnostic table.
            if header and header[0].startswith("รายการ") and any("สินทรัพย์รวม" in r[0] for r in rows[1:]):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                shape = run.add_picture(str(CHART), width=Inches(6.15))
                shape._inline.docPr.set(
                    "descr",
                    "กราฟยอดขาย กำไรสุทธิ และผลจากธุรกิจหลักโดยประมาณของ PG ปี 2566 ถึง 2568",
                )
                shape._inline.docPr.set("title", "PG Sales and Profit Trend")
                cap = doc.add_paragraph(style="Table Source")
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cap.add_run("ภาพ: ยอดขาย กำไรสุทธิ และผลจากธุรกิจหลักโดยประมาณของ PG ปี 2566–2568")
            continue

        if re.match(r"^- ", stripped):
            p = doc.add_paragraph(style="List Bullet")
            parse_inline(p, stripped[2:])
            i += 1
            continue
        if re.match(r"^\d+\. ", stripped):
            # Preserve the number written in Markdown. Word's built-in List Number
            # style otherwise continues numbering across unrelated sections after
            # field updates (for example 1-4 can render as 6-9).
            match = re.match(r"^(\d+)\.\s+(.*)$", stripped)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.38)
            p.paragraph_format.first_line_indent = Inches(-0.22)
            p.add_run(f"{match.group(1)}.  ")
            parse_inline(p, match.group(2))
            i += 1
            continue

        # Merge wrapped Markdown prose until a structural boundary.
        parts = [stripped]
        j = i + 1
        while j < len(lines):
            nxt = lines[j].strip()
            if not nxt:
                break
            if nxt.startswith(("#", ">", "|", "- ")) or re.match(r"^\d+\. ", nxt):
                break
            parts.append(nxt)
            j += 1
        para_text = " ".join(parts).replace("  ", " ")
        p = doc.add_paragraph()
        parse_inline(p, para_text)
        i = j if j > i + 1 else i + 1


def add_final_metadata(doc):
    settings = doc.settings._element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")
    doc.core_properties.title = "PG Strategic Research Dossier 2026"
    doc.core_properties.subject = "Strategy Management research dossier for People's Garment PCL"
    doc.core_properties.author = "Strategy Management Research Team"
    doc.core_properties.keywords = "PG, People's Garment, EFAS, IFAS, SFAS, TOWS, strategy"


def main():
    make_chart()
    doc = Document()
    configure_styles(doc)
    add_cover(doc)
    add_toc(doc)
    doc.add_page_break()
    parse_markdown(doc, SOURCE.read_text(encoding="utf-8"))
    add_final_metadata(doc)
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
