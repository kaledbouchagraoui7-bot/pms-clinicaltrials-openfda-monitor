#!/usr/bin/env python3
"""Unit & integration tests for pms_monitor.py (OpenFDA MAUDE + ClinicalTrials.gov PMS Dossier Generator)."""

import json
import shutil
import tempfile
import unittest
from pathlib import Path

import openpyxl
from docx import Document

from pms_monitor import build_pms_dossier


class TestPMSMonitor(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = Path(tempfile.mkdtemp(prefix="pms_test_"))
        self.sample_path = Path(__file__).resolve().parent / "sample_pms_input.json"
        self.data = json.loads(self.sample_path.read_text(encoding="utf-8"))

    def tearDown(self) -> None:
        shutil.rmtree(self.temp_dir, ignore_errors=True)

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


if __name__ == "__main__":
    unittest.main()
