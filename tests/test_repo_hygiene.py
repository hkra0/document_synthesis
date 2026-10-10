#!/usr/bin/env python3
"""Repository hygiene and anti-leak regression tests.

Verifies:
1. No developer-specific absolute paths (user home directory) or email
   addresses exist in any tracked repository text files (excluding the LICENSE copyright line).
2. Proper .gitattributes line ending (eol=lf) and binary definitions are configured.
3. Required exclusion rules in .gitignore exist.
4. Evidence-hashes.json records match: tracked files strictly; ignored raw
   evidence (logs, DOCX/PDF exports kept only locally) when present.
"""

from __future__ import annotations

import hashlib
import json
import re
import subprocess
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]

# Construct patterns dynamically to prevent self-matching in git grep
USER_DIR_PATTERN = re.compile(r"/" + r"Users" + r"/|/" + r"home" + r"/")
WIN_USER_PATTERN = re.compile(r"[A-Za-z]:\\\\" + r"Users" + r"\\\\|[A-Za-z]:/" + r"Users" + r"/")
EMAIL_PATTERN = re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b")

BINARY_EXTENSIONS = {
    ".docx", ".pdf", ".png", ".jpg", ".jpeg", ".pptx", ".pyc", ".ico", ".bin"
}


def is_binary_file(path: Path) -> bool:
    if path.suffix.lower() in BINARY_EXTENSIONS:
        return True
    try:
        chunk = path.read_bytes()[:2048]
        return b"\x00" in chunk
    except Exception:
        return True


class RepoHygieneTest(unittest.TestCase):
    def test_no_forbidden_user_paths_or_emails(self):
        try:
            out = subprocess.check_output(
                ["git", "ls-files"],
                cwd=REPO_ROOT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            tracked_files = [line.strip() for line in out.splitlines() if line.strip()]
        except Exception as exc:
            self.skipTest(f"git ls-files unavailable: {exc}")

        violations: list[str] = []

        for rel_path in tracked_files:
            file_path = REPO_ROOT / rel_path
            if not file_path.is_file():
                continue
            if is_binary_file(file_path):
                continue

            try:
                content = file_path.read_text(encoding="utf-8", errors="replace")
            except Exception:
                continue

            lines = content.splitlines()
            for line_no, line in enumerate(lines, start=1):
                # LICENSE copyright line is exempt
                if file_path.name == "LICENSE" and "Copyright" in line:
                    continue

                if USER_DIR_PATTERN.search(line):
                    violations.append(f"{rel_path}:{line_no} [user-path] {line.strip()[:100]}")
                elif WIN_USER_PATTERN.search(line):
                    violations.append(f"{rel_path}:{line_no} [win-user-path] {line.strip()[:100]}")
                elif EMAIL_PATTERN.search(line):
                    violations.append(f"{rel_path}:{line_no} [email] {line.strip()[:100]}")

        self.assertEqual(
            violations,
            [],
            "Found hygiene violations in tracked files:\n" + "\n".join(violations),
        )

    def test_gitattributes_configured(self):
        gitattributes = REPO_ROOT / ".gitattributes"
        self.assertTrue(gitattributes.is_file(), ".gitattributes file must exist")
        text = gitattributes.read_text(encoding="utf-8")

        self.assertIn("* text=auto eol=lf", text)
        self.assertIn("*.json", text)
        self.assertIn("*.md", text)
        self.assertIn("*.py", text)
        self.assertIn("*.toml", text)
        self.assertIn("eol=lf", text)
        self.assertIn("*.docx binary", text)
        self.assertIn("*.png binary", text)
        self.assertIn("*.pdf binary", text)

    def test_gitignore_contains_p0_rules(self):
        gitignore = REPO_ROOT / ".gitignore"
        self.assertTrue(gitignore.is_file(), ".gitignore file must exist")
        text = gitignore.read_text(encoding="utf-8")

        self.assertIn("修改/", text)
        self.assertIn("tmp_*.pdf", text)
        self.assertIn(".test_*", text)

    def test_evidence_hashes_integrity(self):
        evidence_hash_files = [
            REPO_ROOT / "docs" / "acceptance" / "n0-n10-review" / "evidence-hashes.json",
            REPO_ROOT / "docs" / "acceptance" / "r0-r11-review" / "evidence-hashes.json",
            REPO_ROOT / "docs" / "acceptance" / "word-followup" / "evidence-hashes.json",
        ]

        try:
            out = subprocess.check_output(
                ["git", "ls-files"], cwd=REPO_ROOT, text=True,
                encoding="utf-8", errors="replace",
            )
        except Exception as exc:
            self.skipTest(f"git ls-files unavailable: {exc}")
        tracked = {line.strip() for line in out.splitlines() if line.strip()}

        mismatches: list[str] = []
        checked_tracked = 0
        for hash_file in evidence_hash_files:
            self.assertTrue(hash_file.is_file(), f"Missing hash file: {hash_file}")
            records = json.loads(hash_file.read_text(encoding="utf-8"))

            for file_key, expected_hash in records.items():
                candidates = [REPO_ROOT / file_key, hash_file.parent / file_key]
                rel_candidates = {
                    c.resolve().relative_to(REPO_ROOT.resolve()).as_posix()
                    for c in candidates
                }
                is_tracked = bool(rel_candidates & tracked)
                target = next((c for c in candidates if c.is_file()), None)

                if target is None:
                    # Raw evidence is gitignored and exists only where it was
                    # produced; a tracked file must always be present.
                    if is_tracked:
                        mismatches.append(f"File not found: {file_key} (in {hash_file.name})")
                    continue
                if is_tracked:
                    checked_tracked += 1

                actual_hash = hashlib.sha256(target.read_bytes()).hexdigest()
                if actual_hash != expected_hash:
                    mismatches.append(
                        f"Hash mismatch for {file_key} in {hash_file.name}: "
                        f"expected {expected_hash}, got {actual_hash}"
                    )

        self.assertEqual(
            mismatches,
            [],
            "Evidence hashes do not match current disk files:\n" + "\n".join(mismatches),
        )
        self.assertGreater(checked_tracked, 0, "No tracked evidence file was verified.")


if __name__ == "__main__":
    unittest.main()
