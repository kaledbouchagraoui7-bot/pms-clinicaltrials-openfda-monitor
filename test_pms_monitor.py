#!/usr/bin/env python3
"""Unit & integration tests for pms_monitor.py (OpenFDA MAUDE + ClinicalTrials.gov PMS Dossier Generator)."""

import json
import shutil
import tempfile
import unittest
from pathlib import Path

from unittest import mock

import openpyxl
from docx import Document

from pms_monitor import (
    _clean_problem_label,
    _humanize_trial_status,
    build_pms_dossier,
    compute_maude_signal_spikes,
    fetch_live_pms_dataset,
)


class TestPMSMonitor(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = Path(tempfile.mkdtemp(prefix="pms_test_"))
        self.sample_path = Path(__file__).resolve().parent / "sample_pms_input.json"
        self.data = json.loads(self.sample_path.read_text(encoding="utf-8"))

    def tearDown(self) -> None:
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_fetch_live_pms_preserves_openfda_or_operator(self) -> None:
        seen_urls: list[str] = []

        def _fake_get(url: str, **_kwargs):
            seen_urls.append(url)
            m = mock.MagicMock()
            m.status_code = 200
            m.json.return_value = {"results": []}
            return m

        with mock.patch("pms_monitor.requests.get", side_effect=_fake_get):
            fetch_live_pms_dataset("ventilator", "respiratory failure")

        event_urls = [u for u in seen_urls if "/device/event.json" in u]
        self.assertTrue(event_urls)
        self.assertTrue(any("+OR+" in u and "%2BOR%2B" not in u for u in event_urls), f"Bad URL encoding: {event_urls}")

    def test_build_pms_dossier_creates_all_deliverables(self) -> None:
        res = build_pms_dossier(self.data, self.temp_dir)
        self.assertTrue(res["docx"].exists(), "Word PSUR report must exist")
        self.assertTrue(res["xlsx"].exists(), "Excel PMS workbook must exist")
        self.assertTrue(res["png"].exists() and res["png"].stat().st_size > 15000, "300-DPI PNG dashboard must exist")
        self.assertTrue(res["pdf"].exists(), "Vector PDF dashboard must exist")
        self.assertTrue(res["zip"].exists(), "Zipped PMS dossier must exist")

        doc = Document(str(res["docx"]))
        self.assertEqual(len(doc.tables), 2)
        self.assertGreater(doc.sections[0].page_width, doc.sections[0].page_height)

        wb = openpyxl.load_workbook(str(res["xlsx"]))
        self.assertEqual(
            wb.sheetnames,
            ["MAUDE_Adverse_Events", "Recalls_and_510k", "ClinicalTrials_Landscape", "README"],
        )

    def test_large_counts_humanized_statuses_and_clean_labels(self) -> None:
        self.assertEqual(_humanize_trial_status("ACTIVE_NOT_RECRUITING"), "Active (not recruiting)")
        self.assertEqual(_humanize_trial_status("UNKNOWN"), "Status unknown")
        self.assertEqual(_humanize_trial_status("COMPLETED"), "Completed")
        self.assertEqual(_clean_problem_label("No Answer Provided"), "Unspecified device event problem (OpenFDA MAUDE)")
        self.assertEqual(_clean_problem_label("High Test Results"), "High Test Results")

        big_data = json.loads(json.dumps(self.data))
        big_data["device_category"] = (
            "Extremely Long Ventilator Manufacturer Name Inc. (example-medtech-device.com) — ventilator"
        )
        big_data["maude_summary"]["total_events_analyzed"] = 559335
        big_data["maude_summary"]["event_type_counts"]["Malfunction"] = 481253
        big_data["maude_summary"]["top_device_problems"] = [
            {"problem": "Degraded", "count": 304380, "pct": 54.4},
            {"problem": "No Answer Provided", "count": 12000, "pct": 2.1},
        ]
        big_data["clinical_trials"][0]["status"] = "ACTIVE_NOT_RECRUITING"
        big_data["clinical_trials"][0]["enrollment"] = 12500

        res = build_pms_dossier(big_data, self.temp_dir)
        self.assertTrue(res["png"].exists() and res["png"].stat().st_size > 15000)
        doc = Document(str(res["docx"]))
        full_txt = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("559,335", full_txt)
        self.assertIn("481,253", full_txt)

    def test_maude_zscore_signal_spike_detection_and_outputs(self) -> None:
        series = [
            {"month": "2025-03", "count": 28},
            {"month": "2025-04", "count": 31},
            {"month": "2025-05", "count": 29},
            {"month": "2025-06", "count": 30},
            {"month": "2025-07", "count": 32},
            {"month": "2025-08", "count": 30},
            {"month": "2025-09", "count": 29},
            {"month": "2025-10", "count": 31},
            {"month": "2025-11", "count": 28},
            {"month": "2025-12", "count": 30},
            {"month": "2026-01", "count": 31},
            {"month": "2026-02", "count": 61},
        ]
        spikes = compute_maude_signal_spikes(series, z_threshold=2.0)
        self.assertEqual(len(spikes), 1)
        self.assertEqual(spikes[0]["month"], "2026-02")
        self.assertGreater(spikes[0]["z_score"], 10.0)

        res = build_pms_dossier(self.data, self.temp_dir)
        doc = Document(str(res["docx"]))
        self.assertIn("SIGNAL SPIKE ALERT in 2026-02", "\n".join(p.text for p in doc.paragraphs))
        wb = openpyxl.load_workbook(str(res["xlsx"]))
        ws1 = wb["MAUDE_Adverse_Events"]
        categories = [ws1.cell(row=r, column=2).value for r in range(2, ws1.max_row + 1)]
        self.assertIn("MDCG 2022-21 Signal Spike Alert (Z >= 2.0)", categories)
        self.assertTrue(res["png"].exists() and res["png"].stat().st_size > 25000)

        # Also verify clean PNG/PDF rendering when baseline has zero signal spikes
        no_spike_data = json.loads(json.dumps(self.data))
        no_spike_data["maude_summary"]["monthly_trend"] = [
            {"month": "2026-01", "count": 30},
            {"month": "2026-02", "count": 31},
        ]
        res_clean = build_pms_dossier(no_spike_data, self.temp_dir / "no_spike")
        self.assertTrue(res_clean["png"].exists() and res_clean["png"].stat().st_size > 25000)
        self.assertTrue(res_clean["pdf"].exists() and res_clean["pdf"].stat().st_size > 5000)


if __name__ == "__main__":
    unittest.main()
