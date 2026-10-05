from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.document import Document as _Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


ROOT = Path(r"E:\strategy")
OUT = ROOT / "research" / "comparison_extract"
OUT.mkdir(parents=True, exist_ok=True)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

FILES = [
    ROOT / "บริษัท ประชาอาภรณ์ จำกัด (มหาชน) ทั้งเล่ม.docx",
    ROOT / "บริษัท ประชาอาภรณ์ จำกัด (มหาชน) ทั้งเล่มใหม่.docx",
    ROOT / "บริษัท ประชาอาภรณ์ จำกัด เมกุแบบแก้.docx",
]


def iter_blocks(parent: _Document):
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def paragraph_record(paragraph: Paragraph, index: int) -> dict:
    text = clean(paragraph.text)
    style = paragraph.style.name if paragraph.style else ""
    xml = paragraph._p.xml
    return {
        "kind": "paragraph",
        "index": index,
        "style": style,
        "text": text,
        "page_break": "w:type=\"page\"" in xml or "w:lastRenderedPageBreak" in xml,
        "numbered": paragraph._p.pPr is not None and paragraph._p.pPr.numPr is not None,
    }


def table_record(table: Table, index: int) -> dict:
    rows = []
    for row in table.rows:
        rows.append([clean(cell.text.replace("\n", " / ")) for cell in row.cells])
    return {
        "kind": "table",
        "index": index,
        "rows": len(rows),
        "columns": max((len(r) for r in rows), default=0),
        "data": rows,
    }


def analyze(path: Path) -> dict:
    doc = Document(str(path))
    blocks = []
    p_index = 0
    t_index = 0
    for block in iter_blocks(doc):
        if isinstance(block, Paragraph):
            rec = paragraph_record(block, p_index)
            p_index += 1
            if rec["text"] or rec["page_break"]:
                blocks.append(rec)
        else:
            t_index += 1
            blocks.append(table_record(block, t_index))

    all_text = "\n".join(
        b["text"] if b["kind"] == "paragraph" else "\n".join(" | ".join(r) for r in b["data"])
        for b in blocks
    )
    headings = [
        b for b in blocks
        if b["kind"] == "paragraph"
        and b["text"]
        and (
            b["style"].lower().startswith("heading")
            or b["style"].lower() in {"title", "subtitle", "toc 1", "toc 2"}
            or re.match(r"^(บทที่|บท\s*ที่|ส่วนที่|ภาคผนวก|บทนำ|คำนำ|สารบัญ|บรรณานุกรม|เอกสารอ้างอิง)", b["text"], re.I)
        )
    ]
    rels = doc.part.rels.values()
    hyperlinks = sorted({rel.target_ref for rel in rels if "hyperlink" in rel.reltype})
    images = [rel.target_ref for rel in rels if "image" in rel.reltype]
    sections = []
    for i, section in enumerate(doc.sections, 1):
        sections.append({
            "index": i,
            "width_inches": round(section.page_width.inches, 2),
            "height_inches": round(section.page_height.inches, 2),
            "top_margin_inches": round(section.top_margin.inches, 2),
            "bottom_margin_inches": round(section.bottom_margin.inches, 2),
            "left_margin_inches": round(section.left_margin.inches, 2),
            "right_margin_inches": round(section.right_margin.inches, 2),
        })
    urls = sorted(set(re.findall(r"https?://[^\s)\]>]+", all_text)))
    years = sorted(set(re.findall(r"\b(?:20\d{2}|25\d{2})\b", all_text)))
    result = {
        "file": str(path),
        "paragraphs": p_index,
        "tables": t_index,
        "blocks": blocks,
        "headings": headings,
        "characters": len(all_text),
        "words_rough": len(re.findall(r"\S+", all_text)),
        "images": len(images),
        "hyperlinks": hyperlinks,
        "visible_urls": urls,
        "years": years,
        "sections": sections,
    }
    return result


def write_text(record: dict) -> Path:
    path = OUT / (Path(record["file"]).stem + ".txt")
    lines = [f"FILE: {record['file']}"]
    for block in record["blocks"]:
        if block["kind"] == "paragraph":
            flags = []
            if block["numbered"]:
                flags.append("numbered")
            if block["page_break"]:
                flags.append("page-break")
            suffix = f" [{' '.join(flags)}]" if flags else ""
            lines.append(f"\n[P{block['index']} | {block['style']}]{suffix}\n{block['text']}")
        else:
            lines.append(f"\n===== TABLE {block['index']} ({block['rows']}x{block['columns']}) =====")
            lines.extend(" | ".join(row) for row in block["data"])
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def main() -> None:
    summaries = []
    for path in FILES:
        record = analyze(path)
        target = write_text(record)
        summary = {k: v for k, v in record.items() if k not in {"blocks"}}
        summary["text_file"] = str(target)
        summaries.append(summary)
        print(
            f"{path.name}: paragraphs={record['paragraphs']} tables={record['tables']} "
            f"chars={record['characters']} images={record['images']} headings={len(record['headings'])}"
        )
    (OUT / "summary.json").write_text(
        json.dumps(summaries, ensure_ascii=False, indent=2), encoding="utf-8"
    )


if __name__ == "__main__":
    main()
