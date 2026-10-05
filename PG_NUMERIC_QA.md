# PG Numeric & Consistency QA

ไฟล์ที่ตรวจ: `PG_Strategic_Management_Final_2569_REVISED.docx`  
วิธีตรวจ: `python-docx` + สูตรคำนวณอิสระ + Microsoft Word PDF render + visual review 35 หน้า  
ผลรวม: **PASS 20/20**

## Checklist ตามข้อกำหนด

| # | รายการตรวจ | ผล | หลักฐาน/ผลคำนวณ |
|---:|---|:---:|---|
| 1 | EFAS weights = 1.00 | PASS | O = 0.38, T = 0.62, รวม 1.00 |
| 2 | IFAS weights = 1.00 | PASS | S = 0.47, W = 0.53, รวม 1.00 |
| 3 | SFAS weights = 1.00 | PASS | 14 factors รวม 1.00 |
| 4 | KSF weights ต่อ arena = 1.00 | PASS | Export OEM = 1.00; Uniform/Project = 1.00 |
| 5 | Weighted score = Weight × Rating | PASS | ตรวจทุกแถว EFAS/IFAS/SFAS/KSF; ไม่พบส่วนต่าง |
| 6 | Decision matrix weights = 1.00 | PASS | น้ำหนัก 0.20+0.15+0.20+0.10+0.15+0.10+0.10 = 1.00 |
| 7 | Scenario calculations reconcile | PASS | DIO ใช้ average inventory/COGS×365; income/SFP เชื่อมกับ assumption schedule |
| 8 | Gross Profit = Revenue × Gross Margin | PASS | ตรวจปี 2570–2574 ทุกปี |
| 9 | Core proxy = GP − SG&A | PASS | −15.00, 6.45, 24.50, 41.525, 60.75 ลบ. |
| 10 | Percentage denominators | PASS | GM, SG&A/sales, core/sales, product mix ใช้ยอดขายเป็นตัวหารที่ถูกต้อง |
| 11 | Budget low/high = project totals | PASS | 102.5 / 174.0 ลบ. |
| 12 | Projected BS balances | PASS | Assets − Liabilities − Equity = 0.00 ทุกปี 2570–2574 |
| 13 | ROE denominator documented | PASS | Net Profit / average equity; ได้ 2.48%, 3.51%, 4.32%, 5.14%, 6.01% |
| 14 | No identifier collision | PASS | TOWS = SO/WO/ST/WT; Objectives = OBJ1–OBJ9; Projects = P1–P11 |
| 15 | Current board not older than May 2026 | PASS | 10 directors/4 independent as of 12 May 2026; ไม่พบข้อความ 15/5/33.3% |
| 16 | Evidence cutoff ≤ 31 Aug 2026 | PASS | ระบุวันที่ตัดข้อมูลในปก/core properties; future effective dates เป็นข้อกำหนดที่ประกาศแล้วก่อน cutoff |
| 17 | Unsupported claims not stated as fact | PASS | ใช้ analyst estimate/inference/planning assumption และ cautious negative-absence wording |
| 18 | Major claims traceable | PASS | source note ครบ 25/25 ตาราง; appendix เชื่อม evidence→factor→strategy→OBJ→P→KPI→finance |
| 19 | No unfinished placeholder | PASS | ไม่พบ TODO, ทำเพิ่มเติม, ทำใหม่, รหัสนักศึกษา, ชื่ออาจารย์, missing-logo note |
| 20 | Full-document consistency review | PASS | อ่าน/ตรวจภาพครบ 35 หน้า; 25 ตาราง; โครงสร้าง 1–6, references, appendix ครบ |

## Matrix reconciliation

### EFAS

| ปัจจัย | Weight | Rating | Weighted |
|---|---:|---:|---:|
| O1 | 0.14 | 3 | 0.42 |
| O2 | 0.10 | 3 | 0.30 |
| O3 | 0.07 | 2 | 0.14 |
| O4 | 0.07 | 2 | 0.14 |
| T1 | 0.16 | 2 | 0.32 |
| T2 | 0.14 | 2 | 0.28 |
| T3 | 0.12 | 2 | 0.24 |
| T4 | 0.09 | 3 | 0.27 |
| T5 | 0.11 | 2 | 0.22 |
| **รวม** | **1.00** |  | **2.33** |

### IFAS

| ปัจจัย | Weight | Rating | Weighted |
|---|---:|---:|---:|
| S1 | 0.09 | 4 | 0.36 |
| S2 | 0.11 | 4 | 0.44 |
| S3 | 0.09 | 4 | 0.36 |
| S4 | 0.07 | 3 | 0.21 |
| S5 | 0.04 | 4 | 0.16 |
| S6 | 0.07 | 3 | 0.21 |
| W1 | 0.16 | 1 | 0.16 |
| W2 | 0.11 | 2 | 0.22 |
| W3 | 0.09 | 2 | 0.18 |
| W4 | 0.06 | 2 | 0.12 |
| W5 | 0.06 | 2 | 0.12 |
| W6 | 0.05 | 2 | 0.10 |
| **รวม** | **1.00** |  | **2.64** |

### SFAS

| ปัจจัย | Weight | Rating | Weighted |
|---|---:|---:|---:|
| S1 | 0.04 | 4 | 0.16 |
| S2 | 0.09 | 4 | 0.36 |
| S3 | 0.07 | 4 | 0.28 |
| S4 | 0.04 | 3 | 0.12 |
| S6 | 0.05 | 3 | 0.15 |
| W1 | 0.17 | 1 | 0.17 |
| W2 | 0.08 | 2 | 0.16 |
| W3 | 0.07 | 2 | 0.14 |
| W4 | 0.07 | 2 | 0.14 |
| O1 | 0.11 | 3 | 0.33 |
| O2 | 0.05 | 3 | 0.15 |
| T1 | 0.07 | 2 | 0.14 |
| T2 | 0.04 | 2 | 0.08 |
| T5 | 0.05 | 2 | 0.10 |
| **รวม** | **1.00** |  | **2.48** |

KSF totals:

- Export OEM arena: PG 2.55, Hi-Tech benchmark 4.75
- Uniform/project arena: PG 2.90, Warrix benchmark 3.85

## Strategic decision matrix

| ทางเลือก | Base score | Sensitivity 1: market 0.30/resource fit 0.05 | Sensitivity 2: C margin = 2 | Sensitivity 3: B market & margin = 5 |
|---|---:|---:|---:|---:|
| A Volume-led Export OEM | 2.15 | 2.30 | 2.15 | 2.15 |
| B Own-brand D2C | 2.70 | 3.00 | 2.70 | 3.05 |
| C Focused Functional Uniform & Workwear | **3.90** | **3.75** | **3.50** | **3.90** |

ผล: C ชนะทุก sensitivity case ที่กำหนด โดยไม่ได้พึ่งสมมติฐานเดียว

## Strategic income scenario

หน่วย: ล้านบาท เว้นแต่ระบุ

| รายการ | 2570 | 2571 | 2572 | 2573 | 2574 |
|---|---:|---:|---:|---:|---:|
| Revenue | 600.00 | 645.00 | 700.00 | 755.00 | 810.00 |
| Gross margin | 22.50% | 24.00% | 25.50% | 26.50% | 27.50% |
| Gross profit | 135.00 | 154.80 | 178.50 | 200.075 | 222.75 |
| SG&A | 150.00 | 148.35 | 154.00 | 158.55 | 162.00 |
| Core proxy | −15.00 | 6.45 | 24.50 | 41.525 | 60.75 |
| Other income | 58.00 | 56.00 | 54.00 | 54.00 | 54.00 |
| Net profit (20% planning tax) | 34.40 | 49.96 | 62.80 | 76.42 | 91.80 |
| Ending inventory | 300.00 | 260.00 | 220.00 | 190.00 | 165.00 |
| DIO | 250.05 | 208.49 | 167.98 | 134.84 | 110.32 |

สูตรที่ตรวจ:

- `Gross Profit = Revenue × Gross Margin`
- `Core Proxy = Gross Profit − SG&A`
- `DIO = ((Beginning Inventory + Ending Inventory) / 2) / COGS × 365`
- CAGR 2568–2574 = `(810 / 605.20)^(1/6) − 1 = 4.97785%` ≈ **4.98%**

## Projected statement of financial position

หน่วย: ล้านบาท

| ปี | Assets | Liabilities | Equity | L+E | Difference | ROE on average equity |
|---:|---:|---:|---:|---:|---:|---:|
| 2570 | 1,550.49 | 143.41 | 1,407.08 | 1,550.49 | 0.00 | 2.48% |
| 2571 | 1,581.42 | 144.38 | 1,437.04 | 1,581.42 | 0.00 | 3.51% |
| 2572 | 1,615.99 | 146.15 | 1,469.84 | 1,615.99 | 0.00 | 4.32% |
| 2573 | 1,653.07 | 146.81 | 1,506.26 | 1,653.07 | 0.00 | 5.14% |
| 2574 | 1,695.42 | 147.36 | 1,548.06 | 1,695.42 | 0.00 | 6.01% |

ROE denominator:

`ROE_t = Net Profit_t / ((Equity_(t−1) + Equity_t) / 2)`

Opening-equity bridge สำหรับ 2570 ระบุในตารางสมมติฐาน: equity ปี 2568 + FY2569 net-profit assumption − dividend paid + H1/2569 OCI; ต้องแทนด้วย audited FY2569 closing equity ก่อนใช้ตัดสินใจจริง

## Budget reconciliation

| โครงการ | Low | High |
|---|---:|---:|
| P1 | 1.5 | 3.0 |
| P3 | 8.0 | 14.0 |
| P4 | 10.0 | 18.0 |
| P5 | 6.0 | 10.0 |
| P6 | 8.0 | 14.0 |
| P7 | 5.0 | 9.0 |
| P8 | 45.0 | 75.0 |
| P9 | 6.0 | 10.0 |
| P10 | 8.0 | 13.0 |
| P11 | 5.0 | 8.0 |
| **รวม** | **102.5** | **174.0** |

P2 ใช้งบดำเนินการภายใน จึงไม่มีงบเพิ่มในยอดรวมโครงการ  
P8 45–75 ลบ. ตลอด 3 ปี = 15–25 ลบ./ปี = 0.94–1.57 เท่าของ machine investment ปี 2568 ที่ 15.96 ลบ./ปี

## Render QA

- PDF final: 35 หน้า; A4 portrait/landscape ตาม section
- ตรวจ PNG ทุกหน้าแล้วที่ความละเอียด 110 dpi/มุมมอง 100%
- สารบัญครบ Heading 1/2 และเลขหน้าตรงกับ final render
- ไม่พบตารางล้นขอบ, ข้อความถูกตัด, placeholder หรือ source note ที่ขาด
