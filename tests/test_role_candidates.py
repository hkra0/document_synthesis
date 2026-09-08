# -*- coding: utf-8 -*-
"""S5 numbering, conflict, and abstention contracts."""

import unittest
import hashlib
import json
from pathlib import Path
from types import SimpleNamespace

from lib.format_analysis import analyze_format_sample
from lib.role_candidates import rank_unstructured_candidates


ROOT = Path(__file__).resolve().parents[1]


def _block(path, text, size=12, *, bold=False, alignment=None, before=None):
    return SimpleNamespace(
        node=SimpleNamespace(element_path=path),
        structure_type="paragraph",
        visible_text=text,
        effective_run=SimpleNamespace(size_pt=size, bold=bold),
        effective_paragraph=SimpleNamespace(
            alignment=alignment,
            space_before_pt=before,
        ),
    )


class RoleCandidateRemediationTest(unittest.TestCase):
    def test_S5_date_like_decimal_is_not_numbering_evidence_or_auto_heading(self):
        blocks = [
            _block(f"/w:document/w:body/w:p[{index}]", "普通正文。" * 30)
            for index in range(1, 4)
        ]
        blocks.insert(
            0,
            _block(
                "/w:document/w:body/w:p[4]",
                "2026.09.07 版本说明",
                18,
                bold=True,
            ),
        )
        evidence = rank_unstructured_candidates(blocks)["/w:document/w:body/w:p[4]"]
        self.assertEqual(evidence["numbering_evidence"]["kind"], "date_or_version")
        self.assertEqual(evidence["selected_role"], "body")
        self.assertEqual(evidence["review_status"], "pending_review")
        self.assertNotEqual(evidence["score_kind"], "calibrated_confidence_band")

    def test_S5_numbering_depth_conflict_keeps_heading3_candidate_pending(self):
        blocks = [
            _block("/w:document/w:body/w:p[1]", "1. Introduction", 18, bold=True),
            _block("/w:document/w:body/w:p[2]", "正文内容。" * 30),
            _block("/w:document/w:body/w:p[3]", "1.1.1. Detailed protocol", 14, bold=True),
            _block("/w:document/w:body/w:p[4]", "正文内容。" * 30),
        ]
        evidence = rank_unstructured_candidates(blocks)["/w:document/w:body/w:p[3]"]
        self.assertEqual(evidence["selected_role"], "heading.3")
        self.assertEqual(evidence["review_status"], "pending_review")
        self.assertTrue(any(
            item["kind"] == "numbering_vs_size_level"
            for item in evidence["conflicts"]
        ))
        self.assertNotEqual(evidence["score_kind"], "calibrated_confidence_band")

    def test_S5_frozen_hard_cases_bind_real_fixtures_and_fail_closed(self):
        root = ROOT / "docs" / "acceptance" / "r7-evaluation"
        manifest = json.loads((root / "unstructured-hard-cases.json").read_text(encoding="utf-8"))
        self.assertTrue(manifest["frozen"])
        self.assertEqual(manifest["document_count"], 4)
        self.assertEqual(set(manifest["categories"]), {"公文", "报告", "论文", "手工格式"})
        observed_cases = set()
        for item in manifest["documents"]:
            path = root / "unstructured-hard-samples" / item["filename"]
            self.assertTrue(path.is_file())
            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), item["source_sha256"])
            report = analyze_format_sample(path)
            evidence = report["unstructured_candidate_evidence"]
            for check in item["checks"]:
                observed_cases.add(check["case"])
                node_path = f"/w:document/w:body/w:p[{check['paragraph_index'] + 1}]"
                candidate = evidence[node_path]
                self.assertEqual(candidate["selected_role"], check["expected_role"], (item["filename"], check))
                if "expected_review_status" in check:
                    self.assertEqual(candidate["review_status"], check["expected_review_status"])
                if "expected_numbering_kind" in check:
                    self.assertEqual(candidate["numbering_evidence"]["kind"], check["expected_numbering_kind"])
                if "required_conflict" in check:
                    self.assertIn(check["required_conflict"], {item["kind"] for item in candidate["conflicts"]})
        self.assertEqual(observed_cases, set(manifest["case_types"]))


if __name__ == "__main__":
    unittest.main()
