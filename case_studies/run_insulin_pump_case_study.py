#!/usr/bin/env python3
"""
Reproducible OpenFDA MAUDE + ClinicalTrials.gov v2 Research Study:
Ambulatory Insulin Infusion Pumps (EU MDR Class IIb / MDCG 2022-21).
Outputs verified JSON metrics + 24-month retrospective Z-score sensitivity check + 300-DPI figure.
"""

from __future__ import annotations

import json
import sys
import urllib.parse
from pathlib import Path

import requests

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import pms_monitor


def fetch_monthly_maude_counts(device_term: str = "insulin pump", months_limit: int = 36) -> list[dict]:
    """Query daily OpenFDA receivedate counts and aggregate by YYYY-MM over trailing 36 months."""
    q_enc = urllib.parse.quote(
        f'device.generic_name:"{device_term}"+OR+device.brand_name:"{device_term}"',
        safe='+:"()',
    )
    url = f"https://api.fda.gov/device/event.json?search={q_enc}&count=date_received"
    resp = requests.get(url, headers={"User-Agent": pms_monitor.USER_AGENT}, timeout=15)
    resp.raise_for_status()
    monthly: dict[str, int] = {}
    for row in resp.json().get("results", []):
        t = str(row.get("time", ""))
        if len(t) >= 6:
            ym = f"{t[:4]}-{t[4:6]}"
            monthly[ym] = monthly.get(ym, 0) + int(row.get("count", 0))
    ordered = [{"month": k, "count": monthly[k]} for k in sorted(monthly)]
    # Drop final potentially partial current month if incomplete? Keep complete slice up to -1 for analysis clarity
    return ordered[-months_limit:]


def main() -> int:
    out_dir = Path(__file__).resolve().parent / "artifacts"
    out_dir.mkdir(parents=True, exist_ok=True)

    payload = pms_monitor.fetch_live_pms_dataset(
        device_term="insulin pump",
        condition_term="type 1 diabetes",
        eu_mdr_class="Class IIb / Active Infusion System (EU MDR)",
        limit_trials=15,
    )
    payload["target_company"] = "Public Class IIb Cohort"
    payload["device_category"] = "Ambulatory Insulin Infusion Pumps — OpenFDA MAUDE & ClinicalTrials.gov v2"

    # Extended 36-month series to evaluate Z-score baseline burn-in sensitivity
    monthly_36 = fetch_monthly_maude_counts("insulin pump", months_limit=36)
    complete_35 = monthly_36[:-1] if len(monthly_36) > 12 else monthly_36
    spikes_36 = pms_monitor.compute_maude_signal_spikes(complete_35, z_threshold=2.0, min_baseline=6)

    payload["extended_analysis_36m"] = {
        "months_analyzed": len(complete_35),
        "monthly_counts": complete_35,
        "z_threshold_2_0_spikes": spikes_36,
    }

    json_path = out_dir / "insulin_pump_live_dataset.json"
    png_path = out_dir / "insulin_pump_psur_dashboard.png"
    pdf_path = out_dir / "insulin_pump_psur_dashboard.pdf"

    json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    pms_monitor.generate_pms_dashboard(payload, png_path, pdf_path)

    ms = payload.get("maude_summary", {})
    print("TOTAL_EVENTS:", ms.get("total_events_analyzed"))
    print("EVENT_COUNTS:", ms.get("event_type_counts"))
    print("TOP_PROBLEMS:", ms.get("top_device_problems"))
    print("DEFAULT_12M_SPIKES:", ms.get("signal_spikes"))
    print("EXTENDED_35M_SPIKES (Z>=2.0, min_base=6):", spikes_36)
    print("TRIALS_COUNT:", len(payload.get("clinical_trials", [])))
    print("RECALLS_510K:", len(payload.get("recalls_and_510k", [])))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
