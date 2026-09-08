#!/usr/bin/env python3
"""Path sanitization and evidence hash synchronization utility.

Used to replace machine-specific absolute paths in tracked evidence files with
portable placeholders such as <repo>, <word-access-dir>, and <downloads>, and
synchronize corresponding evidence-hashes.json files.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path


def find_repo_root(start: Path | None = None) -> Path:
    current = (start or Path(__file__)).resolve()
    for parent in [current] + list(current.parents):
        if (parent / ".git").is_dir() or (parent / "synthesize.py").is_file():
            return parent
    return current.parents[3]


REPO_ROOT = find_repo_root()


def relativize(path: Path | str, repo_root: Path | None = None) -> str:
    """Convert an absolute or environment path into a standardized portable representation."""
    root = (repo_root or REPO_ROOT).resolve()
    raw_str = str(path).replace("\\", "/")
    root_str = str(root).replace("\\", "/")

    if raw_str == root_str or raw_str.startswith(root_str + "/"):
        rel = raw_str[len(root_str):].lstrip("/")
        return f"<repo>/{rel}" if rel else "<repo>"

    # Check for known word access directories and temp paths
    if raw_str == "/private/tmp/document-synthesis-word-access" or raw_str == "/tmp/document-synthesis-word-access":
        return "<word-access-dir>"
    if raw_str.startswith("/private/tmp/document-synthesis-word-access/"):
        return f"<word-access-dir>/{raw_str[len('/private/tmp/document-synthesis-word-access/'):]}"
    if raw_str.startswith("/tmp/document-synthesis-word-access/"):
        return f"<word-access-dir>/{raw_str[len('/tmp/document-synthesis-word-access/'):]}"

    if raw_str.startswith("/private/tmp/"):
        return f"<word-access-dir>/{raw_str[len('/private/tmp/'):]}"
    if raw_str.startswith("/tmp/"):
        return f"<word-access-dir>/{raw_str[len('/tmp/'):]}"

    # Downloads directory
    downloads_match = re.search(r"/(?:Users|home)/[^/]+/Downloads(?:/(.*))?", raw_str)
    if downloads_match:
        sub = downloads_match.group(1)
        return f"<downloads>/{sub}" if sub else "<downloads>"

    # Generic user directory
    user_match = re.search(r"/(?:Users|home)/[^/]+(?:/(.*))?", raw_str)
    if user_match:
        sub = user_match.group(1)
        return f"<user>/{sub}" if sub else "<user>"

    return raw_str


def sanitize_text(text: str, repo_root: Path | None = None) -> str:
    """Replace personal, repo and system temporary paths in arbitrary text."""
    root = (repo_root or REPO_ROOT).resolve()
    root_str = str(root)

    # 1. Repo root
    result = text.replace(root_str, "<repo>")
    # In case path had symlink resolution differences (e.g. /private/var vs /var)
    if root_str.startswith("/private"):
        alt_root = root_str[len("/private"):]
        result = result.replace(alt_root, "<repo>")

    # 2. Known specific locations
    result = re.sub(
        r"/(?:Users|home)/[^/\s\"'\)]+/Downloads",
        "<downloads>",
        result,
    )
    result = result.replace(
        "/private/tmp/document-synthesis-word-access",
        "<word-access-dir>",
    )
    result = result.replace(
        "/tmp/document-synthesis-word-access",
        "<word-access-dir>",
    )

    # 3. Generic temp directories inside /private/tmp/
    result = result.replace("/private/tmp/", "<word-access-dir>/")

    # 4. Any remaining user paths
    result = re.sub(
        r"/(?:Users|home)/[^/\s\"'\)]+",
        "<user>",
        result,
    )

    return result


def compute_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def update_evidence_hashes(repo_root: Path | None = None) -> list[tuple[Path, int]]:
    root = (repo_root or REPO_ROOT).resolve()
    hash_files = [
        root / "docs" / "acceptance" / "n0-n10-review" / "evidence-hashes.json",
        root / "docs" / "acceptance" / "r0-r11-review" / "evidence-hashes.json",
        root / "docs" / "acceptance" / "word-followup" / "evidence-hashes.json",
    ]

    results: list[tuple[Path, int]] = []
    for hash_file in hash_files:
        if not hash_file.is_file():
            continue
        data = json.loads(hash_file.read_text(encoding="utf-8"))
        updated_count = 0
        new_data = {}
        for file_key, old_hash in data.items():
            # Try repo relative first, then directory relative
            target_path = root / file_key
            if not target_path.is_file():
                target_path = hash_file.parent / file_key

            if target_path.is_file():
                current_hash = compute_sha256(target_path)
                new_data[file_key] = current_hash
                if current_hash != old_hash:
                    updated_count += 1
            else:
                new_data[file_key] = old_hash

        hash_file.write_text(
            json.dumps(new_data, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        results.append((hash_file, updated_count))

    return results


def get_tracked_files(repo_root: Path) -> list[str]:
    try:
        out = subprocess.check_output(
            ["git", "ls-files"],
            cwd=repo_root,
            text=True,
        )
        return [line.strip() for line in out.splitlines() if line.strip()]
    except Exception as exc:
        print(f"Warning: git ls-files failed ({exc}), falling back to disk traversal", file=sys.stderr)
        return []


def run_check(repo_root: Path) -> int:
    """Check whether tracked files contain forbidden absolute user paths."""
    forbidden_pattern = re.compile(
        r"/" + r"(?:Users|home)/|[A-Za-z]:\\" + r"Users\\|[A-Za-z]:/" + r"Users/"
    )
    email_pattern = re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b")

    violations = []
    tracked = get_tracked_files(repo_root)

    for rel_path in tracked:
        path = repo_root / rel_path
        if not path.is_file():
            continue
        # Skip binary files by extension
        if path.suffix.lower() in {".docx", ".pdf", ".png", ".jpg", ".jpeg", ".pptx", ".pyc"}:
            continue
        try:
            content = path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            continue

        lines = content.splitlines()
        for idx, line in enumerate(lines, start=1):
            if path.name == "LICENSE" and "Copyright" in line:
                continue
            if forbidden_pattern.search(line):
                violations.append((rel_path, idx, "user_path", line.strip()[:100]))
            elif email_pattern.search(line):
                violations.append((rel_path, idx, "email", line.strip()[:100]))

    if violations:
        print(f"Found {len(violations)} hygiene violation(s):", file=sys.stderr)
        for rel_path, line_no, vtype, snippet in violations:
            print(f"  {rel_path}:{line_no} [{vtype}] {snippet}", file=sys.stderr)
        return 1

    print("Hygiene check passed: no user paths or emails found in tracked files.")
    return 0


def run_sanitize(repo_root: Path) -> int:
    """Perform sanitization on target evidence files and update evidence hashes."""
    # Target file patterns in docs/acceptance
    acceptance_dir = repo_root / "docs" / "acceptance"
    if not acceptance_dir.is_dir():
        print("docs/acceptance directory not found", file=sys.stderr)
        return 1

    target_extensions = {".json", ".txt", ".md", ".log"}
    modified_files = []

    for file_path in sorted(acceptance_dir.rglob("*")):
        if not file_path.is_file():
            continue
        if file_path.suffix.lower() not in target_extensions:
            continue
        # Don't sanitize source-hashes.json as it records code hashes
        if file_path.name == "source-hashes.json":
            continue

        try:
            orig = file_path.read_text(encoding="utf-8")
        except Exception:
            continue

        sanitized = sanitize_text(orig, repo_root)
        if sanitized != orig:
            file_path.write_text(sanitized, encoding="utf-8")
            modified_files.append(file_path.relative_to(repo_root))

    print(f"Sanitized {len(modified_files)} evidence file(s):")
    for f in modified_files:
        print(f"  {f}")

    print("Updating evidence hashes...")
    hash_results = update_evidence_hashes(repo_root)
    for hf, count in hash_results:
        print(f"  {hf.relative_to(repo_root)}: updated {count} entries")

    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Sanitize paths in acceptance evidence and synchronize hashes.")
    parser.add_argument("--check", action="store_true", help="Check for absolute user paths without modifying")
    parser.add_argument("--write", action="store_true", help="Sanitize evidence files and update hashes")
    parser.add_argument("--update-hashes-only", action="store_true", help="Only recompute evidence-hashes.json")
    args = parser.parse_args()

    repo = REPO_ROOT

    if args.update_hashes_only:
        results = update_evidence_hashes(repo)
        for hf, count in results:
            print(f"Updated {hf.relative_to(repo)}: {count} hash(es)")
        return 0

    if args.check:
        return run_check(repo)

    # Default or --write
    return run_sanitize(repo)


if __name__ == "__main__":
    raise SystemExit(main())
