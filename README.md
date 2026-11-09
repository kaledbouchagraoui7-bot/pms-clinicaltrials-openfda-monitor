# Automated ClinicalTrials.gov + OpenFDA MAUDE Post-Market Surveillance (PMS) Monitor

**Automated EU MDR 2017/745 Article 83–86 (PMS / PSUR), MDCG 2022-21, and FDA 21 CFR Part 803 Competitive Safety & Clinical Trial Landscape Surveillance Engine.**

[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![Regulatory Scope](https://img.shields.io/badge/Scope-EU_MDR_PSUR_%7C_OpenFDA_MAUDE_%7C_ClinicalTrials.gov-0D9488.svg)](#overview)

---

## Overview

Under **EU MDR 2017/745 Articles 83–86** and **MDCG 2022-21**, medical device manufacturers and Clinical Research Organizations (CROs) must continuously monitor similar/competitor devices across **FDA MAUDE (adverse events, recalls, 510(k) clearances)** and **`ClinicalTrials.gov` (ongoing PMCF and competitor clinical investigations)** to compile their **Periodic Safety Update Report (PSUR)** and **Clinical Evaluation Report (CER)** updates.

`pms_monitor.py` automates this workflow and produces a complete, audit-ready **PMS / PSUR Dossier**:

1. **Landscape Word PSUR Safety & Clinical Trial Report (`.docx`)** — Pre-formatted tables covering MAUDE adverse event severity (`Malfunction`, `Injury`, `Death`), FDA recalls & 510(k) clearances, and tracked `ClinicalTrials.gov` studies.
2. **4-Sheet Filterable Excel Surveillance Workbook (`.xlsx`)** — `MAUDE_Adverse_Events`, `Recalls_and_510k`, `ClinicalTrials_Landscape` (with clickable `https://clinicaltrials.gov/study/<NCTId>` hyperlinks), and `README`.
3. **300-DPI Visual PMS & Clinical Landscape Dashboard (`.png` & `.pdf`)** — High-resolution executive chart ready to embed in PSUR decks or client briefings.
4. **Zipped Deliverable Package (`<slug>_dossier.zip`)**.

---

## Quickstart

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Generate Dossier from Sample Snapshot (Deterministic Demo)

```bash
python pms_monitor.py --input sample_pms_input.json --output-dir examples
```

### 3. Query Live OpenFDA + ClinicalTrials.gov APIs

```bash
python pms_monitor.py --live-device "insulin pump" --live-condition "type 1 diabetes" --output-dir output
```

### 4. Run Automated Unit Tests

```bash
python -m unittest discover -p "test_*.py" -v
```

---

## Data Source Terms & Legal Compliance

- **ClinicalTrials.gov API v2**: Subject to NLM/NIH terms at https://clinicaltrials.gov/
- **OpenFDA API**: Public domain (CC0 1.0) subject to terms at https://open.fda.gov/apis/ and https://open.fda.gov/license
- See [`LEGAL_COMPLIANCE_AND_ETHICS.md`](LEGAL_COMPLIANCE_AND_ETHICS.md) for complete Swiss (`StGB`, `UWG`, `nFADP`) and EU (`MDR`, `GDPR`) operational, data-integrity, and outreach compliance rules.
