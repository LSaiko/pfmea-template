# pFMEA Generator — AIAG-VDA 2019 | Medical Device Quality Engineering

[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![Standard: AIAG-VDA 2019](https://img.shields.io/badge/standard-AIAG--VDA%202019-darkblue)](https://www.aiag.org/)
[![Regulatory: 21 CFR 820 | ISO 13485](https://img.shields.io/badge/regulatory-21%20CFR%20820%20%7C%20ISO%2013485-green)](https://www.fda.gov/medical-devices/quality-system-qs-regulationmedical-device-good-manufacturing-practices)

A collection of production-ready Python scripts that auto-generate **complete, professionally formatted pFMEA Excel workbooks** following the **AIAG-VDA FMEA 4th Edition (2019)** 7-step methodology across three distinct medical device manufacturing scenarios.

Built as a portfolio piece for a **Manufacturing Transfer Project Engineer** application, demonstrating:
- Deep knowledge of AIAG-VDA 2019 vs. legacy FMEA methodology
- Medical device regulatory fluency (21 CFR 820, ISO 13485, ISO 11607, PMA Class III)
- Python automation for quality engineering documentation
- GMP-compliant document formatting across diverse process types

---

## Samples at a Glance

| # | Script | Device / Process | Classification | Key Standards |
|---|--------|-----------------|----------------|---------------|
| 1 | `pfmea_generator.py` | Ti-6Al-4V Femoral Stem — CNC Machining | Class III PMA | 21 CFR 820, ISO 13485 |
| 2 | `pfmea_generator_spinal.py` | PEEK Lumbar Interbody Cage — Injection Molding | Class II 510(k) | 21 CFR 820, ASTM F2026 |
| 3 | `pfmea_generator_packaging.py` | Sterile Barrier Packaging — Tyvek-Film Heat Sealing | Class III support | ISO 11607-1, 21 CFR 820 |

---

## Workbook Structure — AIAG-VDA 7-Step Approach

Each generated workbook contains the same 7 sheets, populated with scenario-specific content:

| Sheet | Step | Content |
|-------|------|---------|
| **1. Scope & Planning** | Step 1 | Process name, scope, team, device classification, standards |
| **2. Structure Analysis** | Step 2 | System hierarchy: Cell / Line → Station → Components |
| **3. Function Analysis** | Step 3 | Intended functions and requirements satisfied |
| **4. Failure Analysis** | Step 4 | 6 failure modes with effects, severity, causes, controls |
| **5. Risk Rating** | Step 5 | AIAG-VDA Action Priority (AP) — NOT legacy RPN |
| **6. Optimization** | Step 6 | 4 corrective actions with owners, dates, revised AP |
| **7. Results Summary** | Step 7 | Metrics, linked documents, archive location |

---

## AIAG-VDA 2019 Key Differentiator: Action Priority vs. RPN

All three scripts implement the **correct AP lookup table**, not RPN multiplication (`S x O x D`).

```
AP Logic (patient-safety weighting):
  S 9-10 + D 7-10              -> H  (regardless of O)
  S 9-10 + D 4-6 + O >= 4     -> H
  S 9-10 + D 4-6 + O < 4      -> M
  S 7-8  + D 7-10 + O >= 4    -> H
  All others                   -> M or L
```

**Why it matters for medical devices:** An undetected failure reaching a patient is catastrophic regardless of occurrence rate. AP weights Detection gaps more heavily than RPN arithmetic does.

---

## Quick Start

### 1. Clone the repo
```bash
git clone https://github.com/LSaiko/pfmea-template.git
cd pfmea-template
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Generate a workbook

```bash
# Sample 1 — CNC Femoral Stem (Class III PMA)
python pfmea_generator.py
# Output: pfmea_aiag_vda_2019.xlsx

# Sample 2 — PEEK Spinal Cage (Class II 510k)
python pfmea_generator_spinal.py
# Output: pfmea_spinal_peek_cage.xlsx

# Sample 3 — Sterile Packaging (ISO 11607)
python pfmea_generator_packaging.py
# Output: pfmea_sterile_packaging.xlsx

# Custom output path (any script):
python pfmea_generator.py /path/to/my_pfmea.xlsx
```

---

## Requirements

- Python 3.9+
- `openpyxl >= 3.1.2`

No other dependencies.

---

## Sample 1 — CNC Machining: Ti-6Al-4V Femoral Stem

**Device:** Cementless Ti-6Al-4V Femoral Stem — Implantable Hip Prosthesis
**Classification:** Class III PMA (21 CFR 814)

| # | Failure Mode | S | O | D | Initial AP | After Action |
|---|---|---|---|---|---|---|
| 1 | OD out of tolerance (oversize) | 8 | 4 | 3 | M | M |
| 2 | Surface roughness Ra > 0.8 um | 9 | 3 | 2 | M | L |
| 3 | Fixture slip / part shift | 7 | 3 | 4 | M | M |
| 4 | Coolant pressure drop / flow loss | 9 | 3 | 5 | **H** | L |
| 5 | Wrong program revision loaded | 8 | 2 | 3 | M | L |
| 6 | CMM inspection step skipped | 7 | 4 | 7 | **H** | L |

**Corrective Actions:**
1. Coolant flow sensor interlock — feed-hold alarm if PSI < 45
2. eDHR barcode-locked program selection — auto-archives obsolete revisions
3. Insert change interval: 150 → 100 parts; profilometer every 50 parts
4. Mandatory CMM gate-check in eDHR — part cannot advance without CMM pass record

---

## Sample 2 — Injection Molding: PEEK Lumbar Interbody Fusion Cage

**Device:** PEEK Lumbar Interbody Fusion Cage — Spinal Implant
**Classification:** Class II 510(k) (21 CFR 888.3070)

| # | Failure Mode | S | O | D | Initial AP | After Action |
|---|---|---|---|---|---|---|
| 1 | Flash / fin at parting line | 6 | 4 | 3 | L | L |
| 2 | Outer footprint or height OOS | 8 | 4 | 5 | M | L |
| 3 | Wrong PEEK grade loaded | 9 | 3 | 7 | **H** | M |
| 4 | Short shot / incomplete fill | 7 | 3 | 2 | L | L |
| 5 | Gate vestige height > 0.3 mm | 7 | 4 | 4 | M | L |
| 6 | Endplate surface texture OOS | 8 | 5 | 7 | **H** | L |

**Corrective Actions:**
1. Material barcode scan at hopper — eDHR locks start until CoA matches traveler
2. Inline profilometer at ejection — 100% Ra measurement with SPC alert
3. 100% vision system check for footprint and height — auto-reject to quarantine
4. Gate vestige height added to first-article CMM program; increased sampling to Cpk >= 1.67

---

## Sample 3 — Sterile Barrier Packaging: Tyvek-Film Pouch Heat Sealing

**Device:** Sterile Packaging Process supporting Class III Implants
**Standards:** ISO 11607-1:2019 | 21 CFR 820 | ISO 13485:2016

| # | Failure Mode | S | O | D | Initial AP | After Action |
|---|---|---|---|---|---|---|
| 1 | Incomplete / channel seal defect | 9 | 4 | 6 | **H** | M |
| 2 | Pinhole / micro-leak in film layer | 10 | 2 | 8 | **H** | M |
| 3 | Seal temperature out of range | 8 | 4 | 5 | M | L |
| 4 | Incorrect dwell time | 8 | 3 | 3 | L | L |
| 5 | UDI label / content mismatch | 8 | 2 | 4 | L | L |
| 6 | Package breach during EtO / transit | 9 | 3 | 5 | M | M |

**Corrective Actions:**
1. 100% bubble emission test per ASTM F2096 — replaces sampling; SPC peel strength
2. Dye penetration test every batch per ISO 11607-1 Annex A2.2; incoming pin-hole detection
3. Independent thermocouple with dual-channel alarm (Delta T > 3 C triggers feed-hold)
4. Post-EtO 100% visual inspection; ASTM D4169 distribution simulation semi-annually

---

## Repository Structure

```
pfmea-template/
├── pfmea_generator.py              # Sample 1: CNC Ti-6Al-4V Femoral Stem (Class III)
├── pfmea_generator_spinal.py       # Sample 2: Injection Molding PEEK Spinal Cage (Class II)
├── pfmea_generator_packaging.py    # Sample 3: Sterile Barrier Packaging (ISO 11607)
├── pfmea_aiag_vda_2019.xlsx        # Pre-generated output — Sample 1
├── pfmea_spinal_peek_cage.xlsx     # Pre-generated output — Sample 2
├── pfmea_sterile_packaging.xlsx    # Pre-generated output — Sample 3
├── requirements.txt                # openpyxl dependency
└── README.md                       # This file
```

---

## Portfolio Context

This project demonstrates competencies directly relevant to a **Manufacturing Transfer Project Engineer** role at a medical device company:

- **FMEA methodology:** AIAG-VDA 2019 AP table implementation across three process types (machining, molding, packaging)
- **Regulatory breadth:** 21 CFR 820.70(i) software control, ISO 13485 §7.5.3 traceability, 21 CFR 820.184 DHR, ISO 11607-1 sterile barrier, 21 CFR 830 UDI
- **Process knowledge:** Ti-6Al-4V CNC machining, PEEK injection molding, Tyvek-film heat sealing, EtO sterilization, CMM, SPC
- **Documentation:** GMP-compliant formatting, linked document control, archive conventions
- **Python automation:** Clean, self-contained scripts deployable in a quality engineering context

---

## License

MIT License — free to use, modify, and adapt for your own quality engineering projects.

---

*Built with Python + openpyxl | AIAG-VDA FMEA 4th Edition 2019 | 21 CFR 820 | ISO 13485 | ISO 11607*
