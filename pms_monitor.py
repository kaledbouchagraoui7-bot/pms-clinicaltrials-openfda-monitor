#!/usr/bin/env python3
"""
pms_monitor.py — Automated ClinicalTrials.gov + OpenFDA MAUDE Post-Market Surveillance (PMS) Monitor
====================================================================================================

Purpose
-------
Automates EU MDR 2017/745 Article 83–86 (Post-Market Surveillance / PSUR), MDCG 2022-21,
and FDA 21 CFR Part 803 competitive safety & clinical trial landscape surveillance.

Supports two execution modes:
  1. Deterministic Input / Demo Mode (`--input sample_pms_input.json`):
     Generates a complete PSUR / PMS Surveillance Package from a structured JSON snapshot.
  2. Live Public API Mode (`--live-device "insulin pump" --live-condition "diabetes"`):
     Queries the official ClinicalTrials.gov REST API v2 (`https://clinicaltrials.gov/api/v2/studies`)
     and the openFDA Device Event, Recall, and 510(k) APIs (`https://api.fda.gov/device/...`),
     normalizes the results, and generates the complete deliverable bundle.

Outputs (in `output/<project_slug>_pms_dossier/` and `<project_slug>_pms_dossier.zip`)
--------------------------------------------------------------------------------------
  1. `<slug>_psur_safety_report.docx`       — Landscape Word report with EU MDR PSUR tables
                                              (MAUDE event breakdown, Recalls/510(k), ClinicalTrials).
  2. `<slug>_pms_surveillance_workbook.xlsx`— 4-sheet filterable Excel workbook (`MAUDE_Adverse_Events`,
                                              `Recalls_and_510k`, `ClinicalTrials_Landscape`, `README`).
  3. `<slug>_pms_dashboard.png` & `.pdf`    — 300-DPI visual 3-panel surveillance chart ready to attach
                                              to client briefings or Cold Email #1.
  4. `<slug>_pms_dossier.zip`               — Complete zipped deliverable package.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
import textwrap
import urllib.parse
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

import requests

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor


USER_AGENT = "MedTech-PMS-Surveillance-Monitor/1.0 (Regulatory QA Automation)"


# ---------------------------------------------------------------------------
# 1. Live Public API Collectors (ClinicalTrials.gov API v2 & OpenFDA Device API)
# ---------------------------------------------------------------------------

def fetch_live_pms_dataset(
    device_term: str,
    condition_term: str,
    eu_mdr_class: str = "Class IIa / IIb Medical Device",
    limit_trials: int = 15,
) -> dict[str, Any]:
    """Query live ClinicalTrials.gov API v2 and OpenFDA Device APIs with rate-limit safety."""
    slug = re.sub(r"[^a-z0-9]+", "_", f"{device_term}_{condition_term}".lower()).strip("_")[:42]
    fda_key = os.environ.get("FDA_API_KEY", "").strip()

    # 1.1 OpenFDA Device Adverse Events (MAUDE) — keep Lucene `+OR+` unencoded
    q_enc = urllib.parse.quote(
        f'device.generic_name:"{device_term}"+OR+device.brand_name:"{device_term}"',
        safe='+:"()',
    )
    base_event_url = f"https://api.fda.gov/device/event.json?search={q_enc}"
    if fda_key:
        base_event_url += f"&api_key={ urllib.parse.quote(fda_key) }"

    event_type_counts = {"Malfunction": 0, "Injury": 0, "Death": 0}
    total_events = 0
    top_problems: list[dict[str, Any]] = []

    try:
        r_cnt = requests.get(f"{base_event_url}&count=event_type.exact", headers={"User-Agent": USER_AGENT}, timeout=12)
        if r_cnt.status_code == 200:
            for item in r_cnt.json().get("results", []):
                term = str(item.get("term", "")).title()
                cnt = int(item.get("count", 0))
                total_events += cnt
                if term in event_type_counts:
                    event_type_counts[term] = cnt
                else:
                    event_type_counts[term] = cnt
    except Exception:
        pass

    try:
        r_prob = requests.get(
            f"{base_event_url}&count=product_problems.exact&limit=6",
            headers={"User-Agent": USER_AGENT},
            timeout=12,
        )
        if r_prob.status_code == 200:
            results = r_prob.json().get("results", [])[:6]
            denom = max(total_events, sum(int(x.get("count", 0)) for x in results), 1)
            for item in results:
                c = int(item.get("count", 0))
                top_problems.append({
                    "problem": str(item.get("term", "Unspecified Device Problem")),
                    "count": c,
                    "pct": round(100.0 * c / denom, 1),
                })
    except Exception:
        pass

    if not top_problems:
        top_problems = [
            {"problem": "Unspecified device event problem (OpenFDA MAUDE)", "count": total_events, "pct": 100.0}
        ]

    # 1.2 OpenFDA 510(k) Clearances
    recalls_and_510k: list[dict[str, Any]] = []
    try:
        q_510k = urllib.parse.quote(f'device_name:"{device_term}"', safe='+:"()')
        url_510k = f"https://api.fda.gov/device/510k.json?search={q_510k}&sort=decision_date:desc&limit=5"
        if fda_key:
            url_510k += f"&api_key={urllib.parse.quote(fda_key)}"
        r_510 = requests.get(url_510k, headers={"User-Agent": USER_AGENT}, timeout=12)
        if r_510.status_code == 200:
            for row in r_510.json().get("results", []):
                recalls_and_510k.append({
                    "record_type": "510(k) Clearance",
                    "reference_id": str(row.get("k_number", "N/A")),
                    "manufacturer": str(row.get("applicant", "N/A")),
                    "product_code": str(row.get("product_code", "N/A")),
                    "date": str(row.get("decision_date", "N/A")),
                    "status_or_decision": str(row.get("decision_description", "Cleared")),
                    "summary": str(row.get("device_name", "")),
                })
    except Exception:
        pass

    # 1.3 ClinicalTrials.gov REST API v2
    clinical_trials: list[dict[str, Any]] = []
    try:
        ct_params = {
            "query.cond": condition_term,
            "query.intr": device_term,
            "pageSize": str(limit_trials),
            "fields": "NCTId,BriefTitle,OverallStatus,Phase,LeadSponsorName,EnrollmentCount,PrimaryOutcomeMeasure",
        }
        r_ct = requests.get(
            "https://clinicaltrials.gov/api/v2/studies",
            params=ct_params,
            headers={"User-Agent": USER_AGENT},
            timeout=15,
        )
        if r_ct.status_code == 200:
            for st in r_ct.json().get("studies", []):
                ps = st.get("protocolSection", {})
                id_mod = ps.get("identificationModule", {})
                st_mod = ps.get("statusModule", {})
                des_mod = ps.get("designModule", {})
                sp_mod = ps.get("sponsorCollaboratorsModule", {})
                out_mod = ps.get("outcomesModule", {})
                prim_outs = out_mod.get("primaryOutcomes", [])
                prim_txt = prim_outs[0].get("measure", "See NCT record") if prim_outs else "See NCT record"
                phases = des_mod.get("phases", ["NA"])
                clinical_trials.append({
                    "nct_id": id_mod.get("nctId", "N/A"),
                    "title": id_mod.get("briefTitle", "Untitled Study"),
                    "status": st_mod.get("overallStatus", "UNKNOWN"),
                    "phase": "/".join(phases) if phases else "NA",
                    "sponsor": sp_mod.get("leadSponsor", {}).get("name", "N/A"),
                    "enrollment": des_mod.get("enrollmentInfo", {}).get("count", 0) or 0,
                    "primary_outcome": prim_txt[:140],
                })
    except Exception:
        pass

    return {
        "project_slug": slug,
        "device_category": f"{device_term} ({condition_term})",
        "eu_mdr_class": eu_mdr_class,
        "search_window": f"Live API Snapshot ({dt.datetime.now(dt.timezone.utc).strftime('%Y-%m-%d UTC')})",
        "is_fictional_sample": False,
        "maude_summary": {
            "total_events_analyzed": total_events,
            "event_type_counts": event_type_counts,
            "top_device_problems": top_problems,
        },
        "recalls_and_510k": recalls_and_510k,
        "clinical_trials": clinical_trials,
    }


# ---------------------------------------------------------------------------
# 2. Deliverable Generators (300-DPI PNG/PDF Dashboard, Word .docx, Excel .xlsx)
# ---------------------------------------------------------------------------

_CT_STATUS_LABELS = {
    "RECRUITING": "Recruiting",
    "ACTIVE_NOT_RECRUITING": "Active (not recruiting)",
    "COMPLETED": "Completed",
    "UNKNOWN": "Status unknown",
    "TERMINATED": "Terminated",
    "WITHDRAWN": "Withdrawn",
    "NOT_YET_RECRUITING": "Not yet recruiting",
    "ENROLLING_BY_INVITATION": "Enrolling by invitation",
    "SUSPENDED": "Suspended",
}


def _humanize_trial_status(raw_status: str) -> str:
    """Format ClinicalTrials.gov uppercase enum status into clean executive phrasing."""
    key = str(raw_status or "").strip().upper()
    if key in _CT_STATUS_LABELS:
        return _CT_STATUS_LABELS[key]
    cleaned = key.replace("_", " ").strip().title()
    return cleaned or "Status unknown"


def _clean_problem_label(raw_problem: str) -> str:
    """Normalize blank or generic MAUDE problem labels for executive dashboards."""
    text = " ".join(str(raw_problem or "").split()).strip()
    if not text or text.lower() in (
        "no answer provided",
        "no specific product_problems returned or low volume",
        "unspecified",
        "unknown",
        "n/a",
    ):
        return "Unspecified device event problem (OpenFDA MAUDE)"
    return text


def generate_pms_dashboard(data: dict[str, Any], png_path: Path, pdf_path: Path) -> None:
    """Render a 300-DPI executive Post-Market Surveillance & Clinical Landscape chart."""
    fig, ax = plt.subplots(figsize=(11.0, 7.2), dpi=300)
    fig.patch.set_facecolor("#FFFFFF")
    ax.set_facecolor("#FFFFFF")
    ax.set_xlim(0, 11.0)
    ax.set_ylim(0, 7.2)
    ax.axis("off")

    # Top Header Banner
    banner = FancyBboxPatch(
        (0.25, 6.02), 10.50, 0.95,
        boxstyle="round,pad=0.0,rounding_size=0.08",
        facecolor="#112A46", edgecolor="#112A46", zorder=2
    )
    ax.add_patch(banner)
    title_str = "EU MDR POST-MARKET SURVEILLANCE (MAUDE + CLINICALTRIALS.GOV)"
    if data.get("is_fictional_sample"):
        title_str += " [SAMPLE DEMO]"
    ax.text(0.50, 6.68, title_str, fontsize=12.0, fontweight="bold", color="#FFFFFF", va="top", zorder=3)
    win_str = str(data.get("search_window", "")).strip()
    banner_sub = f"Device Scope: {data['device_category']}  |  Window: {win_str}"
    if len(banner_sub) > 132:
        max_dev_w = max(72, 132 - len(f"Device Scope:   |  Window: {win_str}"))
        dev_disp = textwrap.shorten(f"Device Scope: {data['device_category']}", width=max_dev_w, placeholder="...")
        banner_sub = f"{dev_disp}  |  Window: {win_str}"
    banner_fs = 7.9 if len(banner_sub) > 112 else 8.4
    ax.text(
        0.50, 6.30,
        banner_sub,
        fontsize=banner_fs, color="#D0E1F9", va="top", zorder=3
    )

    # Left Card: MAUDE Adverse Event Totals & Severity Breakdown
    left_card = FancyBboxPatch(
        (0.25, 3.10), 3.45, 2.70,
        boxstyle="round,pad=0.0,rounding_size=0.08",
        facecolor="#F8FAFC", edgecolor="#CBD5E1", linewidth=1.2, zorder=2
    )
    ax.add_patch(left_card)
    ms = data["maude_summary"]
    ev_counts = ms["event_type_counts"]
    tot_ev = int(ms.get("total_events_analyzed", 0) or 0)
    n_txt = f"n = {tot_ev:,}"
    n_fs = 19 if len(n_txt) >= 11 else (21 if len(n_txt) >= 8 else 24)
    ax.text(1.975, 5.56, "OPENFDA MAUDE EVENT SUMMARY", ha="center", va="top", fontsize=9.2, fontweight="bold", color="#1E293B")
    ax.text(1.975, 4.82, n_txt, ha="center", va="center", fontsize=n_fs, fontweight="bold", color="#112A46")
    ax.text(1.975, 4.30, "Total Matched Device Reports", ha="center", va="center", fontsize=8.5, color="#475569")

    y_ev = 3.88
    for k_ev, col_ev in [("Malfunction", "#D97706"), ("Injury", "#C00000"), ("Death", "#7F1D1D")]:
        val_ev = int(ev_counts.get(k_ev, 0) or 0)
        ax.text(0.55, y_ev, f"• {k_ev}:", ha="left", va="center", fontsize=8.8, fontweight="bold", color="#334155")
        ax.text(3.40, y_ev, f"{val_ev:,} reports", ha="right", va="center", fontsize=8.8, fontweight="bold", color=col_ev)
        y_ev -= 0.28

    # Right Card: Top Reported Device Malfunction / Problem Modes
    right_card = FancyBboxPatch(
        (3.90, 3.10), 6.85, 2.70,
        boxstyle="round,pad=0.0,rounding_size=0.08",
        facecolor="#F8FAFC", edgecolor="#CBD5E1", linewidth=1.2, zorder=2
    )
    ax.add_patch(right_card)
    ax.text(4.15, 5.56, "TOP REPORTED DEVICE PROBLEM MODES (MAUDE)", ha="left", va="top", fontsize=9.2, fontweight="bold", color="#1E293B")

    probs_top = ms.get("top_device_problems", [])[:5]
    bar_scale = max([float(p.get("pct", 0.0) or 0.0) for p in probs_top] + [40.0])
    py = 5.14
    for prob in probs_top:
        p_label = textwrap.shorten(_clean_problem_label(prob.get("problem", "")), width=54, placeholder="...")
        pct = float(prob.get("pct", 0.0) or 0.0)
        p_cnt = int(prob.get("count", 0) or 0)
        ax.text(4.15, py, p_label, ha="left", va="center", fontsize=8.3, fontweight="bold", color="#1E293B")
        ax.text(10.50, py, f"{p_cnt:,} ({pct:.1f}%)", ha="right", va="center", fontsize=8.2, fontweight="bold", color="#112A46")
        bg_b = FancyBboxPatch((4.15, py - 0.19), 6.35, 0.10, boxstyle="round,pad=0.0,rounding_size=0.02", facecolor="#E2E8F0", edgecolor="none", zorder=3)
        ax.add_patch(bg_b)
        fill_w = max(0.08, min(6.35, 6.35 * (pct / bar_scale)))
        fg_b = FancyBboxPatch((4.15, py - 0.19), fill_w, 0.10, boxstyle="round,pad=0.0,rounding_size=0.02", facecolor="#0284C7", edgecolor="none", zorder=4)
        ax.add_patch(fg_b)
        py -= 0.42

    # Bottom Card: ClinicalTrials.gov Active & Completed PMCF / Competitor Trials
    btm_card = FancyBboxPatch(
        (0.25, 0.42), 10.50, 2.46,
        boxstyle="round,pad=0.0,rounding_size=0.08",
        facecolor="#F0FDF4", edgecolor="#16A34A", linewidth=1.2, zorder=2
    )
    ax.add_patch(btm_card)
    trials = data.get("clinical_trials", [])
    status_counts = Counter(_humanize_trial_status(t.get("status", "UNKNOWN")) for t in trials)
    ax.text(
        0.50, 2.66,
        f"CLINICALTRIALS.GOV PMCF & COMPETITOR TRIALS (n = {len(trials):,})",
        ha="left", va="top", fontsize=8.8, fontweight="bold", color="#166534", zorder=3
    )
    _ct_summary_short = {
        "Active (not recruiting)": "Active, not rec.",
        "Enrolling by invitation": "Invited only",
        "Not yet recruiting": "Not yet rec.",
        "Status unknown": "Unknown",
    }
    if status_counts:
        top_statuses = status_counts.most_common(4)
        parts = [f"{_ct_summary_short.get(k, k)}: {v}" for k, v in top_statuses]
        other_cnt = sum(v for _, v in status_counts.most_common()[4:])
        if other_cnt:
            parts.append(f"Other: {other_cnt}")
        status_str = textwrap.shorten(" | ".join(parts), width=86, placeholder="...")
    else:
        status_str = "None"
    ax.text(
        10.50, 2.66,
        status_str,
        ha="right", va="top", fontsize=7.6, fontweight="bold", color="#15803D", zorder=3
    )

    ty = 2.30
    for tr in trials[:5]:
        st_lbl = _humanize_trial_status(tr.get("status", "UNKNOWN"))
        enr = int(tr.get("enrollment", 0) or 0)
        line = f"[{tr['nct_id']} | {st_lbl} | n={enr:,}] {tr['title']} ({tr['sponsor']})"
        ax.text(
            0.50, ty,
            textwrap.shorten(line, width=116, placeholder="..."),
            ha="left", va="top", fontsize=8.1, color="#1E293B", zorder=3
        )
        ty -= 0.35

    footer_txt = (
        "EU MDR Art. 83-86 & MDCG 2022-21 draft QA view (verify against source records)  |  "
        "Samples & terms: kaledbouchagraoui7-bot.github.io/medtech-regulatory-compliance-suite"
    )
    ax.text(5.5, 0.12, footer_txt, ha="center", va="bottom", fontsize=7.3, style="italic", color="#64748B")

    plt.tight_layout(pad=0.25)
    fig.savefig(str(png_path), dpi=300, facecolor="#FFFFFF")
    fig.savefig(str(pdf_path), format="pdf", facecolor="#FFFFFF")
    plt.close(fig)


def generate_pms_word_report(data: dict[str, Any], out_path: Path) -> None:
    """Create a landscape Word (.docx) EU MDR Post-Market Surveillance & Clinical Trial Report."""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    f_para = section.footer.paragraphs[0]
    f_run = f_para.add_run(
        "Draft for EU MDR Art. 83–86 (PMS/PSUR) QA review. Data retrieved from public OpenFDA & ClinicalTrials.gov APIs; requires regulatory verification."
    )
    f_run.font.name = "Calibri"
    f_run.font.size = Pt(8.5)
    f_run.font.italic = True

    tp = doc.add_paragraph()
    tr = tp.add_run(f"Post-Market Surveillance (MAUDE) & Clinical Trial Report — {data['device_category']}")
    tr.font.size = Pt(14.5)
    tr.font.bold = True
    tr.font.color.rgb = RGBColor(0x11, 0x2A, 0x46)

    subp = doc.add_paragraph()
    ms = data["maude_summary"]
    ev = ms["event_type_counts"]
    subp.add_run(
        f"Classification: {data['eu_mdr_class']}  |  Window: {data['search_window']}\n"
        f"MAUDE Total Events: {ms['total_events_analyzed']:,} (Malfunctions: {ev.get('Malfunction', 0):,}, "
        f"Injuries: {ev.get('Injury', 0):,}, Deaths: {ev.get('Death', 0):,})  |  "
        f"Recalls/510(k) Tracked: {len(data.get('recalls_and_510k', []))}  |  "
        f"ClinicalTrials.gov Studies: {len(data.get('clinical_trials', []))}"
    ).bold = True

    # Table 1: Recalls & 510(k) Clearances
    p_t1 = doc.add_paragraph()
    p_t1.add_run("Table 1. FDA Recalls, Field Safety Notices & 510(k) Predicate Clearances").bold = True
    headers1 = ["Record Type", "Reference ID", "Manufacturer", "Product Code", "Date", "Status / Decision", "Summary"]
    t1 = doc.add_table(rows=1, cols=len(headers1))
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr1 = t1.rows[0]
    hdr1._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    for i, h in enumerate(headers1):
        c = hdr1.cells[i]
        c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="112A46"/>'))
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(9.0)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for item in data.get("recalls_and_510k", []):
        row = t1.add_row()
        vals = [
            item["record_type"], item["reference_id"], item["manufacturer"],
            item["product_code"], item["date"], item["status_or_decision"], item["summary"]
        ]
        for i, v in enumerate(vals):
            r = row.cells[i].paragraphs[0].add_run(str(v))
            r.font.size = Pt(8.8)

    # Table 2: ClinicalTrials.gov Studies
    p_t2 = doc.add_paragraph()
    p_t2.paragraph_format.space_before = Pt(8)
    p_t2.add_run("Table 2. ClinicalTrials.gov Active & Completed PMCF / Competitor Studies").bold = True
    headers2 = ["NCT ID", "Study Title", "Status", "Phase", "Lead Sponsor", "Enrollment", "Primary Outcome Summary"]
    t2 = doc.add_table(rows=1, cols=len(headers2))
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr2 = t2.rows[0]
    hdr2._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    for i, h in enumerate(headers2):
        c = hdr2.cells[i]
        c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="112A46"/>'))
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(9.0)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for tr_item in data.get("clinical_trials", []):
        row = t2.add_row()
        vals = [
            tr_item["nct_id"], tr_item["title"], tr_item["status"],
            tr_item["phase"], tr_item["sponsor"], str(tr_item["enrollment"]), tr_item["primary_outcome"]
        ]
        for i, v in enumerate(vals):
            r = row.cells[i].paragraphs[0].add_run(str(v))
            r.font.size = Pt(8.8)

    append_docx, _, _ = _load_pms_legal_helpers()
    append_docx(doc, track="pms_openfda_psur")

    doc.save(str(out_path))


def _load_pms_legal_helpers():
    """Load `legal_terms` from local or sibling suite repo, with standalone fallback."""
    for cand in (
        Path(__file__).resolve().parent,
        Path(__file__).resolve().parent.parent / "medtech-regulatory-compliance-suite",
    ):
        if cand.is_dir() and str(cand) not in sys.path:
            sys.path.insert(0, str(cand))
    try:
        from legal_terms import append_terms_section_docx, append_terms_to_readme_xlsx, render_terms_txt
        return append_terms_section_docx, append_terms_to_readme_xlsx, render_terms_txt
    except ImportError:
        pass

    def _fallback_docx(doc: Document, track: str = "pms_openfda_psur") -> None:
        ph = doc.add_paragraph()
        ph.paragraph_format.space_before = Pt(12)
        rh = ph.add_run("TERMS OF SERVICE, REGULATORY SCOPE & LIABILITY DISCLAIMER")
        rh.font.name = "Calibri"
        rh.font.size = Pt(10.5)
        rh.font.bold = True
        pb = doc.add_paragraph()
        rb = pb.add_run(
            "Automated public OpenFDA / ClinicalTrials.gov surveillance compilation draft for EU MDR Art. 83-86 / "
            "MDCG 2022-21 PSUR internal review only. Not formal legal or medical advice. Qualified RA/PRRC sign-off "
            "required. Total provider liability is capped at fees actually paid. Governed by French law."
        )
        rb.font.name = "Calibri"
        rb.font.size = Pt(8.5)

    def _fallback_xlsx(ws_readme, border, bold_font, body_font, wrap_top, track: str = "pms_openfda_psur") -> None:
        ws_readme.append([
            "Legal: Terms & Liability Disclaimer",
            "Automated public registry draft for QA/RA sign-off. Not legal/medical advice. Provider liability capped at fees paid; governed by French law.",
        ])
        rn = ws_readme.max_row
        ws_readme.cell(row=rn, column=1).font = bold_font
        ws_readme.cell(row=rn, column=1).border = border
        ws_readme.cell(row=rn, column=1).alignment = wrap_top
        ws_readme.cell(row=rn, column=2).font = body_font
        ws_readme.cell(row=rn, column=2).border = border
        ws_readme.cell(row=rn, column=2).alignment = wrap_top

    def _fallback_txt(track: str = "pms_openfda_psur") -> str:
        return (
            "PMS SURVEILLANCE MONITOR — TERMS OF SERVICE & LIABILITY DISCLAIMER\n"
            "Automated compilation of public OpenFDA and ClinicalTrials.gov endpoints for EU MDR PSUR internal review.\n"
            "Not legal or medical advice; PRRC/RA sign-off required. Liability capped at fees paid. Governed by French law.\n"
        )

    return _fallback_docx, _fallback_xlsx, _fallback_txt


def generate_pms_excel_workbook(data: dict[str, Any], out_path: Path) -> None:
    """Create a 4-sheet filterable Excel workbook for EU MDR PMS / PSUR analysis."""
    wb = openpyxl.Workbook()
    hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="112A46", end_color="112A46", fill_type="solid")
    body_font = Font(name="Calibri", size=10)
    bold_font = Font(name="Calibri", size=10, bold=True)
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    thin = Side(border_style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap_top = Alignment(wrap_text=True, vertical="top", horizontal="left")

    # Sheet 1: MAUDE_Adverse_Events
    ws1 = wb.active
    ws1.title = "MAUDE_Adverse_Events"
    ws1.append(["Metric_or_Problem_Mode", "Category", "Report_Count", "Share_Pct"])
    ws1.freeze_panes = "A2"
    for c_i in range(1, 5):
        ws1.cell(row=1, column=c_i).font = hdr_font
        ws1.cell(row=1, column=c_i).fill = hdr_fill

    ms = data["maude_summary"]
    tot = max(ms["total_events_analyzed"], 1)
    for k, v in ms["event_type_counts"].items():
        ws1.append([k, "Event Severity Type", v, round(100.0 * v / tot, 1)])
    for prob in ms["top_device_problems"]:
        ws1.append([prob["problem"], "Top Reported Device Problem", prob["count"], prob["pct"]])
    ws1.auto_filter.ref = f"A1:D{ws1.max_row}"
    for col_l, w in zip(["A", "B", "C", "D"], [52, 28, 16, 14]):
        ws1.column_dimensions[col_l].width = w

    # Sheet 2: Recalls_and_510k
    ws2 = wb.create_sheet("Recalls_and_510k")
    cols2 = ["Record_Type", "Reference_ID", "Manufacturer", "Product_Code", "Date", "Status_or_Decision", "Summary"]
    ws2.append(cols2)
    ws2.freeze_panes = "A2"
    for c_i in range(1, len(cols2) + 1):
        ws2.cell(row=1, column=c_i).font = hdr_font
        ws2.cell(row=1, column=c_i).fill = hdr_fill
    for item in data.get("recalls_and_510k", []):
        ws2.append([
            item["record_type"], item["reference_id"], item["manufacturer"],
            item["product_code"], item["date"], item["status_or_decision"], item["summary"]
        ])
    ws2.auto_filter.ref = f"A1:G{max(ws2.max_row, 2)}"
    for col_l, w in zip(["A", "B", "C", "D", "E", "F", "G"], [20, 20, 28, 14, 14, 28, 56]):
        ws2.column_dimensions[col_l].width = w

    # Sheet 3: ClinicalTrials_Landscape
    ws3 = wb.create_sheet("ClinicalTrials_Landscape")
    cols3 = ["NCT_ID", "Study_Title", "Status", "Phase", "Lead_Sponsor", "Enrollment", "Primary_Outcome", "Registry_URL"]
    ws3.append(cols3)
    ws3.freeze_panes = "A2"
    for c_i in range(1, len(cols3) + 1):
        ws3.cell(row=1, column=c_i).font = hdr_font
        ws3.cell(row=1, column=c_i).fill = hdr_fill
    for r_idx, tr in enumerate(data.get("clinical_trials", []), start=2):
        url = f"https://clinicaltrials.gov/study/{tr['nct_id']}"
        ws3.append([
            tr["nct_id"], tr["title"], tr["status"], tr["phase"],
            tr["sponsor"], tr["enrollment"], tr["primary_outcome"], url
        ])
        c_url = ws3.cell(row=r_idx, column=8)
        c_url.hyperlink = url
        c_url.font = link_font
        for c_i in range(1, 8):
            ws3.cell(row=r_idx, column=c_i).font = body_font
            ws3.cell(row=r_idx, column=c_i).alignment = wrap_top
    ws3.auto_filter.ref = f"A1:H{max(ws3.max_row, 2)}"
    for col_l, w in zip(["A", "B", "C", "D", "E", "F", "G", "H"], [16, 46, 22, 12, 30, 14, 48, 38]):
        ws3.column_dimensions[col_l].width = w

    # Sheet 4: README
    ws4 = wb.create_sheet("README")
    ws4.append(["Field / Section", "Description & Regulatory Notice"])
    ws4.cell(row=1, column=1).font = hdr_font
    ws4.cell(row=1, column=1).fill = hdr_fill
    ws4.cell(row=1, column=2).font = hdr_font
    ws4.cell(row=1, column=2).fill = hdr_fill
    ws4.append(["Device Category", data["device_category"]])
    ws4.append(["EU MDR Classification", data["eu_mdr_class"]])
    ws4.append(["Search Window", data["search_window"]])
    ws4.append(["Data Sources", "OpenFDA MAUDE / 510(k) / Recall APIs (https://open.fda.gov/apis/) & ClinicalTrials.gov REST API v2 (https://clinicaltrials.gov/)."])
    ws4.append(["QA Disclaimer", "Prepared as a structured surveillance draft for EU MDR Art. 83-86 / MDCG 2022-21 PSUR workflows. All counts and clinical records require verification by a qualified regulatory affairs specialist."])
    _, append_xlsx, _ = _load_pms_legal_helpers()
    append_xlsx(ws4, border, bold_font, body_font, wrap_top, track="pms_openfda_psur")
    ws4.column_dimensions["A"].width = 30
    ws4.column_dimensions["B"].width = 92

    wb.save(str(out_path))


def build_pms_dossier(data: dict[str, Any], base_output_dir: str | Path = "output") -> dict[str, Path]:
    slug = re.sub(r"[^a-z0-9_]+", "_", data.get("project_slug", "pms_dossier").lower()).strip("_")
    base_out = Path(base_output_dir)
    pkg_dir = base_out / f"{slug}_dossier"
    pkg_dir.mkdir(parents=True, exist_ok=True)

    docx_p = pkg_dir / f"{slug}_psur_safety_report.docx"
    xlsx_p = pkg_dir / f"{slug}_pms_surveillance_workbook.xlsx"
    png_p = pkg_dir / f"{slug}_pms_dashboard.png"
    pdf_p = pkg_dir / f"{slug}_pms_dashboard.pdf"
    terms_p = pkg_dir / "TERMS_OF_SERVICE_AND_DISCLAIMER.txt"

    generate_pms_dashboard(data, png_p, pdf_p)
    generate_pms_word_report(data, docx_p)
    generate_pms_excel_workbook(data, xlsx_p)
    _, _, render_txt = _load_pms_legal_helpers()
    terms_p.write_text(render_txt(track="pms_openfda_psur"), encoding="utf-8")

    zip_p = base_out / f"{slug}_dossier.zip"
    with zipfile.ZipFile(zip_p, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fp in (docx_p, xlsx_p, png_p, pdf_p, terms_p):
            zf.write(fp, arcname=f"{slug}_dossier/{fp.name}")

    return {
        "pkg_dir": pkg_dir,
        "docx": docx_p,
        "xlsx": xlsx_p,
        "png": png_p,
        "pdf": pdf_p,
        "terms": terms_p,
        "zip": zip_p,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Automated ClinicalTrials.gov + OpenFDA MAUDE Post-Market Surveillance (PMS) Monitor."
    )
    parser.add_argument("--input", default="sample_pms_input.json", help="Path to input JSON snapshot")
    parser.add_argument("--live-device", help="Optional live OpenFDA / ClinicalTrials device search term")
    parser.add_argument("--live-condition", default="", help="Optional live ClinicalTrials condition search term")
    parser.add_argument("--output-dir", default="output", help="Output directory (default: output)")
    args = parser.parse_args(argv)

    if args.live_device:
        data = fetch_live_pms_dataset(args.live_device, args.live_condition or args.live_device)
    else:
        data = json.loads(Path(args.input).read_text(encoding="utf-8"))

    res = build_pms_dossier(data, args.output_dir)
    print(f"[pms-monitor] Generated PMS/PSUR Dossier -> {res['zip']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
