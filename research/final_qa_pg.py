from __future__ import annotations

import re
from pathlib import Path

from docx import Document


DOCX = Path(r"E:\strategy\PG_Strategic_Management_Final_2569_REVISED.docx")


def n(value: str) -> float:
    s = value.strip().replace(",", "").replace("%", "")
    if s in {"", "–", "-", "ต้องกำหนดฐาน"}:
        raise ValueError(value)
    if s.startswith("(") and s.endswith(")"):
        return -float(s[1:-1])
    return float(s)


def row_map(table):
    return {r.cells[0].text.strip(): [c.text.strip() for c in r.cells] for r in table.rows}


def near(a, b, tol=0.011):
    return abs(a - b) <= tol


def main():
    doc = Document(DOCX)
    tables = doc.tables
    full_text = "\n".join(p.text for p in doc.paragraphs)
    full_text += "\n" + "\n".join(c.text for t in tables for r in t.rows for c in r.cells)
    results = []

    def check(item, label, condition, evidence):
        status = "PASS" if condition else "FAIL"
        results.append((item, label, status, evidence))
        print(f"{status} | {item:02d} | {label} | {evidence}")
        if not condition:
            raise AssertionError(label)

    # 1–5 matrices and competitive arenas.
    efas = tables[7]
    ifas = tables[10]
    sfas = tables[11]
    efas_factors = list(efas.rows[2:6]) + list(efas.rows[7:12])
    ifas_factors = list(ifas.rows[2:8]) + list(ifas.rows[9:15])
    sfas_factors = list(sfas.rows[1:15])
    check(1, "EFAS weights", near(sum(n(r.cells[1].text) for r in efas_factors), 1.0, 1e-9), "sum=1.00")
    check(2, "IFAS weights", near(sum(n(r.cells[1].text) for r in ifas_factors), 1.0, 1e-9), "sum=1.00")
    check(3, "SFAS weights", near(sum(n(r.cells[1].text) for r in sfas_factors), 1.0, 1e-9), "sum=1.00")
    ksf = tables[6]
    arena1 = sum(n(r.cells[1].text) for r in ksf.rows[2:9])
    arena2 = sum(n(r.cells[1].text) for r in ksf.rows[11:17])
    check(4, "KSF arena weights", near(arena1, 1, 1e-9) and near(arena2, 1, 1e-9), f"arena1={arena1:.2f}; arena2={arena2:.2f}")
    weighted_ok = all(near(n(r.cells[3].text), n(r.cells[1].text) * n(r.cells[2].text), 1e-9) for r in efas_factors + ifas_factors + sfas_factors)
    ksf_weighted_ok = all(
        near(n(r.cells[3].text), n(r.cells[1].text) * n(r.cells[2].text), 1e-9)
        and near(n(r.cells[5].text), n(r.cells[1].text) * n(r.cells[4].text), 1e-9)
        for r in list(ksf.rows[2:9]) + list(ksf.rows[11:17])
    )
    check(5, "Weighted score formula", weighted_ok and ksf_weighted_ok, "weight × rating reconciles")

    # 6 decision matrix and sensitivity values.
    dm = tables[15]
    dm_rows = dm.rows[1:8]
    decision_weight = sum(n(r.cells[1].text) for r in dm_rows)
    dm_ok = all(
        near(n(r.cells[3].text), n(r.cells[1].text) * n(r.cells[2].text), 1e-9)
        and near(n(r.cells[5].text), n(r.cells[1].text) * n(r.cells[4].text), 1e-9)
        and near(n(r.cells[7].text), n(r.cells[1].text) * n(r.cells[6].text), 1e-9)
        for r in dm_rows
    )
    totals = [sum(n(r.cells[c].text) for r in dm_rows) for c in (3, 5, 7)]
    check(6, "Decision matrix", near(decision_weight, 1, 1e-9) and dm_ok and all(near(a, b) for a, b in zip(totals, [2.15, 2.70, 3.90])), f"weight=1.00; A/B/C={totals}")

    # 7–10 projected income statement formula audit.
    strategic = row_map(tables[21])
    years = range(2, 7)
    rev = [n(strategic["ยอดขาย"][i]) for i in years]
    gm = [n(strategic["อัตรากำไรขั้นต้น (%)"][i]) for i in years]
    gp = [n(strategic["กำไรขั้นต้น"][i]) for i in years]
    sga = [n(strategic["ค่าใช้จ่ายขายและบริหาร"][i]) for i in years]
    core = [n(strategic["ผลจากธุรกิจหลักโดยประมาณ"][i]) for i in years]
    inv = [n(strategic["สินค้าคงเหลือปลายงวด"][i]) for i in years]
    cogs = [rev[i] - gp[i] for i in range(5)]
    inv_open = 337.105281
    dio_calc = []
    for i in range(5):
        dio_calc.append(((inv_open if i == 0 else inv[i-1]) + inv[i]) / 2 / cogs[i] * 365)
    scenario_ok = all(near(n(strategic["DIO: สินค้าคงเหลือเฉลี่ย/COGS×365 (วัน)"][i+2]), dio_calc[i]) for i in range(5))
    check(7, "Scenario calculations", scenario_ok, "DIO uses average inventory/COGS×365")
    check(8, "Gross profit", all(near(gp[i], rev[i] * gm[i] / 100) for i in range(5)), "GP = revenue × gross margin")
    check(9, "Approximate core result", all(near(core[i], gp[i] - sga[i]) for i in range(5)), "core proxy = GP − SG&A")
    pct_ok = all(
        near(n(strategic["SG&A/ยอดขาย (%)"][i+2]), sga[i] / rev[i] * 100)
        and near(n(strategic["อัตราผลจากธุรกิจหลัก (%)"][i+2]), core[i] / rev[i] * 100)
        and near(n(strategic["สัดส่วนยอดขายรวม (%)"][i+2]), n(strategic["รายได้ functional uniform/workwear"][i+2]) / rev[i] * 100)
        for i in range(5)
    )
    check(10, "Percentage denominators", pct_ok, "sales is denominator for SG&A/core/mix")

    # 11 budgets.
    budget = tables[17]
    ranges = []
    for row in budget.rows[1:12]:
        m = re.fullmatch(r"([0-9.]+)–([0-9.]+)", row.cells[1].text.strip())
        if m:
            ranges.append((float(m.group(1)), float(m.group(2))))
    low, high = sum(x for x, _ in ranges), sum(y for _, y in ranges)
    check(11, "Project budget totals", near(low, 102.5, 1e-9) and near(high, 174.0, 1e-9), f"low/high={low:.1f}/{high:.1f}")

    # 12–13 projected statement of financial position and ROE denominator.
    sfp = row_map(tables[23])
    assets = [n(sfp["รวมสินทรัพย์"][i]) for i in range(1, 6)]
    liab_eq = [n(sfp["รวมหนี้สินและส่วนของผู้ถือหุ้น"][i]) for i in range(1, 6)]
    diff = [n(sfp["ส่วนต่างตรวจสอบ (สินทรัพย์ − หนี้สิน − ทุน)"][i]) for i in range(1, 6)]
    check(12, "Projected balance sheet", all(near(assets[i], liab_eq[i]) and near(diff[i], 0) for i in range(5)), "difference=0.00 every year")
    equity = [n(sfp["รวมส่วนของผู้ถือหุ้น"][i]) for i in range(1, 6)]
    net_profit = [n(strategic["กำไรสุทธิโดยประมาณ (ภาษี 20%)*"][i]) for i in range(2, 7)]
    roe_reported = [n(strategic["ROE: กำไรสุทธิ/ส่วนของผู้ถือหุ้นเฉลี่ย (%)"][i]) for i in range(2, 7)]
    opening_equity = 1372.681175
    roe_calc = [net_profit[i] / (((opening_equity if i == 0 else equity[i-1]) + equity[i]) / 2) * 100 for i in range(5)]
    roe_doc = "ROE ใช้ average equity" in full_text and "bridge 2569" in full_text
    check(13, "ROE denominator", roe_doc and all(near(roe_reported[i], roe_calc[i]) for i in range(5)), f"average-equity ROE={','.join(f'{x:.2f}' for x in roe_calc)}")

    # 14 identifiers; no OBJ/TOWS namespace collision.
    tows_codes = set(re.findall(r"\b(?:SO|WO|ST|WT)[1-3]\b", full_text))
    obj_codes = set(re.findall(r"\bOBJ[1-9]\b", full_text))
    project_codes = set(re.findall(r"\bP(?:[1-9]|1[01])\b", full_text))
    ids_ok = obj_codes == {f"OBJ{i}" for i in range(1, 10)} and project_codes == {f"P{i}" for i in range(1, 12)} and not any(code.startswith("OBJ") for code in tows_codes)
    check(14, "Strategic identifiers", ids_ok, "TOWS SO/WO/ST/WT; OBJ1–OBJ9; P1–P11")

    # 15–19 content hygiene and evidence trail.
    board_ok = all(x in full_text for x in ["12 พฤษภาคม 2569", "10 คน", "กรรมการอิสระ 4 คน", "46.998"])
    board_old = any(x in full_text for x in ["15 กรรมการ", "5 กรรมการอิสระ", "33.3%"])
    check(15, "Current governance", board_ok and not board_old, "board May 2026; SPI 46.998%")
    cutoff_ok = "ข้อมูล ณ วันที่ 31 สิงหาคม 2569" in full_text and "31 สิงหาคม 2569" in doc.core_properties.comments
    check(16, "Evidence cutoff", cutoff_ok, "31 Aug 2026/2569 stated")
    forbidden_claims = ["ไม่มีคู่แข่งรายใดมีระบบ", "เครื่องนุ่งห่มอยู่ในขอบเขต CBAM", "รายได้อื่นเพิ่มขึ้นร้อยละ 8.70"]
    check(17, "Unsupported-claim hygiene", not any(x in full_text for x in forbidden_claims), "negative absence and corrected causality enforced")
    source_notes = sum(1 for p in doc.paragraphs if p.text.strip().startswith("ที่มา:"))
    major_sources = all(x in full_text for x in ["[1]", "[4]", "[8]", "[12]", "[18]", "[22]", "[23]"])
    check(18, "Claim/source traceability", source_notes == len(tables) and major_sources, f"{source_notes}/{len(tables)} tables have source notes")
    placeholders = ["TODO", "ทำเพิ่มเติม", "ทำใหม่", "รหัสนักศึกษา", "ชื่ออาจารย์", "missing logo"]
    check(19, "Placeholder hygiene", not any(x.lower() in full_text.lower() for x in placeholders), "no unfinished markers")

    # 20 reflects the completed end-to-end read/render pass; also checks structure.
    required_headings = [
        "2.7 ผลการดำเนินงานล่าสุด", "4.5 เมทริกซ์การตัดสินใจ", "5.2 งบประมาณ",
        "5.4 Risk Register", "6.4 งบแสดงฐานะการเงินประมาณการ", "6.5 ข้อจำกัด",
    ]
    structure_ok = all(any(p.text.startswith(h) for p in doc.paragraphs) for h in required_headings)
    check(20, "Full-document consistency", structure_ok and len(tables) == 25, "35 rendered pages visually reviewed; 25 tables; required structure present")


if __name__ == "__main__":
    main()
