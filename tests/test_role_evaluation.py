# -*- coding: utf-8 -*-
"""S0/S4 pure metrics contracts for automatic role decisions."""

import unittest

from lib.role_evaluation import evaluate_role_metrics


class RoleEvaluationMetricsTest(unittest.TestCase):
    def test_S0_D05_one_wrong_exact_role_is_not_perfect(self):
        labels = {"h1": "heading.1", "h2": "heading.2"}
        predictions = {"h1": "heading.1", "h2": "heading.1"}
        result = evaluate_role_metrics(
            labels,
            predictions,
            accepted={"h1": "heading.1", "h2": "heading.1"},
            candidate_roles={"h1": ["heading.1"], "h2": ["heading.1", "heading.2"]},
        )
        self.assertEqual(result["heading_detection_precision"], 1.0)
        self.assertEqual(result["exact_role_auto_precision"], 0.5)
        self.assertEqual(result["auto_acceptance_coverage"], 1.0)
        self.assertEqual(result["correct_auto_coverage"], 0.5)

    def test_S0_D05_body_false_positive_stays_in_detection_denominator(self):
        labels = {"h1": "heading.1", "body": "body"}
        predictions = {"h1": "heading.1", "body": "heading.1"}
        result = evaluate_role_metrics(
            labels,
            predictions,
            accepted={"h1": "heading.1", "body": "heading.1"},
        )
        self.assertEqual(result["heading_detection_precision"], 0.5)
        self.assertEqual(result["exact_role_auto_precision"], 0.5)
        self.assertEqual(result["body_as_heading_false_positives"], 1)

    def test_S0_D05_empty_acceptance_has_null_precision(self):
        result = evaluate_role_metrics(
            {"h1": "heading.1"},
            {"h1": "heading.1"},
            accepted={},
        )
        self.assertIsNone(result["heading_detection_precision"])
        self.assertIsNone(result["exact_role_auto_precision"])
        self.assertEqual(result["auto_acceptance_coverage"], 0.0)
        self.assertEqual(result["correct_auto_coverage"], 0.0)

    def test_S0_D05_no_heading_document_does_not_report_one(self):
        result = evaluate_role_metrics(
            {"body": "body"},
            {"body": "body"},
            accepted={},
        )
        self.assertIsNone(result["auto_acceptance_coverage"])
        self.assertIsNone(result["correct_auto_coverage"])

    def test_S0_D05_missing_labels_are_not_silently_dropped(self):
        with self.assertRaisesRegex(ValueError, "缺少人工标签"):
            evaluate_role_metrics(
                {"h1": "heading.1"},
                {"h1": "heading.1", "unknown": "heading.1"},
                accepted={"h1": "heading.1", "unknown": "heading.1"},
            )

    def test_S0_D05_candidate_exact_role_coverage_is_separate(self):
        result = evaluate_role_metrics(
            {"h1": "heading.1", "h2": "heading.3"},
            {"h1": "heading.1", "h2": "heading.2"},
            accepted={"h1": "heading.1"},
            candidate_roles={"h1": ["heading.1"], "h2": ["heading.2"]},
        )
        self.assertEqual(result["candidate_exact_role_coverage"], 0.5)


if __name__ == "__main__":
    unittest.main()
