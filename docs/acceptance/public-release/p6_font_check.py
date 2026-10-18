#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""P6 F5 字体可移植性验收：按 P6 计划第 1 节退出条件，串行运行生产 CLI 并记录结果。

为 Windows + Word 编写，也能在 macOS 上运行（Word 字体列表的交叉核对只在 Windows 上检查）。
“缺字体”用保证不存在的哨兵名 ``P6缺失测试字体`` 构造，覆写文件在 ``p6-fonts/``。

    python docs/acceptance/public-release/p6_font_check.py
    python docs/acceptance/public-release/p6_font_check.py --skip-matrix --skip-word-tests
    python docs/acceptance/public-release/p6_font_check.py --only F06,F08

结果写入 ``docs/acceptance/public-release/<label>/p6-font-check.json``（``--label`` 默认按平台为
``windows-word`` 或 ``macos-word``）；仓库路径、临时目录与用户目录替换为占位符。
构建产物放在临时目录，结束后删除（``--keep-work`` 指定目录时复制保留，不入库）。

状态口径：``passed`` 仅表示步骤中的命令退出码与全部检查都符合预期；``skipped`` 不算通过。
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from typing import Any, Callable, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[2]
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(HERE))

from word_matrix import Scrubber, compare_results  # noqa: E402

FIXTURES = HERE / "p6-fonts"
SENTINEL = "P6缺失测试字体"
FALLBACK_ENV = "DOCUMENT_SYNTHESIS_FONT_FALLBACKS"
STEP_TIMEOUT = 3600

CHINESE = (REPO / "examples" / "chinese-demo" / "source", REPO / "examples" / "chinese-demo" / "manifest.json")
MINIMAL = (REPO / "examples" / "minimal-demo" / "source", REPO / "examples" / "minimal-demo" / "manifest.json")
CUSTOM = (REPO / "examples" / "custom-format-demo", REPO / "examples" / "custom-format-demo" / "manifest.json")


class Step:
    def __init__(self, step_id: str, title: str, condition: str) -> None:
        self.id = step_id
        self.title = title
        self.condition = condition
        self.commands: List[Dict[str, Any]] = []
        self.checks: List[Dict[str, Any]] = []
        self.status = "passed"
        self.note: Optional[str] = None

    def check(self, name: str, ok: bool, detail: str = "") -> bool:
        self.checks.append({"check": name, "ok": bool(ok), "detail": detail})
        if not ok:
            self.status = "failed"
        return ok

    def to_dict(self, scrub: Scrubber) -> Dict[str, Any]:
        return scrub.data({
            "id": self.id, "title": self.title, "condition": self.condition, "status": self.status,
            "note": self.note, "commands": self.commands, "checks": self.checks,
        })


class Runner:
    def __init__(self, work: Path) -> None:
        self.work = work

    def run(self, step: Step, args: List[str], *, fallbacks: bool = False,
            env: Optional[Dict[str, str]] = None, label: str = "") -> Tuple[int, str]:
        full_env = dict(os.environ, PYTHONIOENCODING="utf-8")
        full_env.pop(FALLBACK_ENV, None)
        full_env.update(env or {})
        command = [sys.executable] + args + (["--font-fallbacks"] if fallbacks else [])
        started = time.monotonic()
        try:
            result = subprocess.run(command, cwd=REPO, env=full_env, capture_output=True,
                                    encoding="utf-8", errors="replace", timeout=STEP_TIMEOUT)
            code, output = result.returncode, result.stdout + result.stderr
        except subprocess.TimeoutExpired as exc:
            code, output = -1, f"超时 {STEP_TIMEOUT} 秒: {exc}"
        step.commands.append({
            "label": label or args[0],
            "argv": [_short(item) for item in args] + (["--font-fallbacks"] if fallbacks else []),
            "env": {key: value for key, value in (env or {}).items()},
            "exit_code": code,
            "seconds": round(time.monotonic() - started, 1),
            "output_tail": [line for line in output.splitlines() if line.strip()][-12:],
        })
        return code, output

    def build_and_smoke(self, step: Step, project: Tuple[Path, Path], out: Path, *,
                        overrides: List[Path] = (), fallbacks: bool = False,
                        smoke_fallbacks: Optional[bool] = None) -> Tuple[int, str, int, str]:
        source, manifest = project
        common = ["--source", str(source), "--manifest", str(manifest)]
        for path in overrides:
            common += ["--override", str(path)]
        build = self.run(step, ["synthesize.py"] + common + ["--output-dir", str(out)],
                         fallbacks=fallbacks, label="build")
        smoke = self.run(step, ["smoke_test.py"] + common + ["--output-dir", str(out)],
                         fallbacks=fallbacks if smoke_fallbacks is None else smoke_fallbacks, label="smoke")
        return build[0], build[1], smoke[0], smoke[1]


def _short(value: str) -> str:
    return value.replace(str(REPO) + os.sep, "").replace(str(REPO), "<repo>")


def _project_args(project: Tuple[Path, Path], *extra: Path) -> List[str]:
    args = ["--source", str(project[0]), "--manifest", str(project[1])]
    for path in extra:
        args += ["--override", str(path)]
    return args


def _font_lines(output: str) -> Dict[str, str]:
    """``--plan``/``--doctor`` 字体一节：字体名 → 状态标签。"""
    result = {}
    for match in re.finditer(r"^\s*\[(已安装|缺失|无法判断|仅 Word 可用)\] (.+?)  ← ", output, re.M):
        result[match.group(2)] = match.group(1)
    return result


def _docx_xml(path: Path) -> Dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist() if name.endswith(".xml")}


def _normalized_parts(path: Path) -> Dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {
            name: re.sub(rb'w:rsid\w*="[^"]*"', b"", archive.read(name))
            for name in archive.namelist() if name != "docProps/core.xml"
        }


def _single_docx(directory: Path) -> Optional[Path]:
    files = sorted(directory.glob("*.docx"))
    return files[0] if len(files) == 1 else None


# ---------------------------------------------------------------------------
# 步骤
# ---------------------------------------------------------------------------

def step_offline_tests(runner: Runner) -> Step:
    step = Step("F01", "离线全量测试", "6")
    code, output = runner.run(step, ["-m", "unittest", "discover", "-s", "tests"], label="unittest")
    summary = [line for line in output.splitlines() if line.startswith(("Ran ", "OK", "FAILED"))]
    step.check("退出码为 0", code == 0, " ".join(summary))
    return step


def step_doctor_default(runner: Runner) -> Step:
    step = Step("F02", "--doctor（默认配置）：公文字体识别为已安装，Windows 读取 Word 字体列表", "2")
    code, output = runner.run(step, ["synthesize.py", "--doctor"], label="doctor")
    step.check("退出码为 0（精确构建就绪）", code == 0, f"exit={code}")
    fonts = _font_lines(output)
    step.check("默认配置字体全部已安装", bool(fonts) and set(fonts.values()) == {"已安装"},
               json.dumps(fonts, ensure_ascii=False))
    if sys.platform == "win32":
        match = re.search(r"\[信息\] Word 字体列表: (\d+) 个", output)
        step.check("读取到 Word 字体列表", bool(match), match.group(0) if match else "未找到 Word 字体列表一行")
    return step


def step_plan(runner: Runner) -> Step:
    step = Step("F03", "--plan 中文示例：示例使用的英文字体名全部识别为已安装", "2")
    code, output = runner.run(step, ["synthesize.py", "--plan"] + _project_args(CHINESE), label="plan")
    fonts = _font_lines(output)
    step.check("退出码为 0", code == 0, f"exit={code}")
    step.check("全部已安装", bool(fonts) and set(fonts.values()) == {"已安装"}, json.dumps(fonts, ensure_ascii=False))
    return step


def step_doctor_missing(runner: Runner) -> Step:
    step = Step("F04", "--doctor 中文示例 + 哨兵字体：报告缺失，退出码不变", "3")
    missing = FIXTURES / "v2-missing.override.json"
    code, output = runner.run(step, ["synthesize.py", "--doctor"] + _project_args(CHINESE, missing), label="doctor")
    fonts = _font_lines(output)
    step.check("退出码为 0", code == 0, f"exit={code}")
    step.check("哨兵字体报告为缺失", fonts.get(SENTINEL) == "缺失", f"{SENTINEL}: {fonts.get(SENTINEL)}")
    step.check("输出替换提示", "Word 会用其他字体替换缺失的字体" in output)
    return step


def step_build_missing(runner: Runner, work: Path) -> Step:
    step = Step("F05", "中文示例 + 哨兵字体，不启用回退：精确构建与 smoke 通过，终端记录缺失", "3")
    out = work / "F05"
    build, build_out, smoke, _ = runner.build_and_smoke(step, CHINESE, out,
                                                        overrides=[FIXTURES / "v2-missing.override.json"])
    step.check("构建退出码为 0", build == 0, f"exit={build}")
    step.check("smoke 退出码为 0", smoke == 0, f"exit={smoke}")
    step.check("终端记录缺失且未替换", f"[字体] 本机缺少且未替换: {SENTINEL}" in build_out)
    docx = _single_docx(out)
    if step.check("生成一份交付物", docx is not None):
        styles = _docx_xml(docx).get("word/styles.xml", b"")
        step.check("正文样式保留请求的字体名", f'w:eastAsia="{SENTINEL}"'.encode() in styles)
    return step


def step_build_fallback(runner: Runner, work: Path) -> Step:
    step = Step("F06", "中文示例 + 哨兵字体，--font-fallbacks：写入仿宋，构建与 smoke 通过", "4")
    out = work / "F06"
    build, build_out, smoke, _ = runner.build_and_smoke(step, CHINESE, out, fallbacks=True,
                                                        overrides=[FIXTURES / "v2-fallback.override.json"])
    step.check("构建退出码为 0", build == 0, f"exit={build}")
    step.check("smoke --font-fallbacks 退出码为 0", smoke == 0, f"exit={smoke}")
    step.check("终端记录回退", f"[字体] 回退: {SENTINEL} → 仿宋" in build_out)
    docx = _single_docx(out)
    if step.check("生成一份交付物", docx is not None):
        parts = _docx_xml(docx)
        step.check("正文样式写入仿宋", 'w:eastAsia="仿宋"'.encode() in parts.get("word/styles.xml", b""))
        leftover = [name for name, data in parts.items() if SENTINEL.encode() in data]
        step.check("文档中没有哨兵字体名", not leftover, ", ".join(leftover))
    return step


def step_fallback_unavailable(runner: Runner) -> Step:
    step = Step("F07", "回退候选全部缺失：启用回退时失败", "5")
    bad = FIXTURES / "v2-fallback-unavailable.override.json"
    code, output = runner.run(step, ["synthesize.py", "--plan"] + _project_args(CHINESE, bad),
                              fallbacks=True, label="plan")
    step.check("退出码为 1", code == 1, f"exit={code}")
    step.check("说明回退字体都没有安装", "声明的回退字体也都没有安装" in output)
    return step


def step_v3_fallback(runner: Runner, work: Path) -> Step:
    step = Step("F08", "custom-format-demo（v3）+ 哨兵字体，--font-fallbacks：元数据记录替换", "4")
    out = work / "F08"
    override = FIXTURES / "v3-fallback.override.json"
    build, _, smoke, _ = runner.build_and_smoke(step, CUSTOM, out, overrides=[override], fallbacks=True)
    step.check("构建退出码为 0", build == 0, f"exit={build}")
    step.check("smoke --font-fallbacks 退出码为 0", smoke == 0, f"exit={smoke}")
    metadata_path = out / "build-metadata.json"
    if step.check("生成 build-metadata.json", metadata_path.is_file()):
        fonts = json.loads(metadata_path.read_text(encoding="utf-8")).get("fonts", {})
        step.check("fallbacks_enabled 为 true", fonts.get("fallbacks_enabled") is True)
        step.check("记录 哨兵 → 宋体", fonts.get("substitutions") == [
            {"requested": SENTINEL, "used": "宋体", "reason": "missing"}], json.dumps(fonts.get("substitutions"), ensure_ascii=False))
    code, output = runner.run(step, ["smoke_test.py"] + _project_args(CUSTOM, override) + ["--output-dir", str(out)],
                              label="smoke（不加 --font-fallbacks）")
    step.check("smoke 不加 --font-fallbacks 时失败", code != 0, f"exit={code}")
    step.check("失败原因为配置契约不一致", "配置契约 SHA-256" in output)
    return step


def step_regression(runner: Runner, work: Path) -> Step:
    step = Step("F09", "三个示例在不加与加 --font-fallbacks 时都通过，且交付物相同", "6")
    for name, project in (("minimal", MINIMAL), ("chinese", CHINESE), ("custom", CUSTOM)):
        outputs = {}
        for fallbacks in (False, True):
            out = work / f"F09-{name}-{'on' if fallbacks else 'off'}"
            build, build_out, smoke, _ = runner.build_and_smoke(step, project, out, fallbacks=fallbacks)
            mode = "加" if fallbacks else "不加"
            step.check(f"{name} {mode} --font-fallbacks 构建与 smoke 通过", build == 0 and smoke == 0,
                       f"build={build} smoke={smoke}")
            step.check(f"{name} {mode} --font-fallbacks 没有字体提示", "[字体]" not in build_out)
            outputs[fallbacks] = _single_docx(out)
        if outputs[False] and outputs[True]:
            a, b = _normalized_parts(outputs[False]), _normalized_parts(outputs[True])
            differing = sorted(name_ for name_ in set(a) | set(b) if a.get(name_) != b.get(name_))
            step.check(f"{name} 两种方式的 DOCX 部件相同（不计 core.xml 与 rsid）", not differing, ", ".join(differing))
    return step


def step_matrix(runner: Runner, label: str) -> Step:
    step = Step("F10", "WW01–WW12：不加与加 DOCUMENT_SYNTHESIS_FONT_FALLBACKS=1，与入库结果页码对照", "6")
    baseline_path = HERE / label / "ww-results.json"
    if not step.check("有入库的基线结果", baseline_path.is_file(), _short(str(baseline_path))):
        return step
    baseline = json.loads(baseline_path.read_text(encoding="utf-8"))
    for suffix, env in (("", {}), ("-fallbacks", {FALLBACK_ENV: "1"})):
        output_name = f"ww-results-p6{suffix}.json"
        code, _ = runner.run(step, ["docs/acceptance/public-release/word_matrix.py", "run", "--label", label,
                                    "--output", output_name, "--note", f"P6 F5 {env or '未启用回退'}"],
                             env=env, label=f"word_matrix {output_name}")
        result_path = HERE / label / output_name
        if not step.check(f"{output_name} 已生成", code == 0 and result_path.is_file(), f"exit={code}"):
            continue
        result = json.loads(result_path.read_text(encoding="utf-8"))
        step.check(f"{output_name} 12/12 verified", result.get("summary", {}).get("verified") == 12,
                   json.dumps(result.get("summary"), ensure_ascii=False))
        comparison = compare_results(baseline, result)
        step.check(f"{output_name} 与基线页码差异为 0",
                   comparison["summary"].get("identical") == 12 and comparison["summary"].get("with_differences") == 0,
                   json.dumps(comparison["summary"], ensure_ascii=False))
    return step


def step_word_tests(runner: Runner) -> Step:
    step = Step("F11", "含真实 Word 用例的全量测试", "6")
    code, output = runner.run(step, ["-m", "unittest", "discover", "-s", "tests"],
                              env={"DOCUMENT_SYNTHESIS_WORD_TEST": "1"}, label="unittest（真实 Word）")
    summary = [line for line in output.splitlines() if line.startswith(("Ran ", "OK", "FAILED"))]
    step.check("退出码为 0", code == 0, " ".join(summary))
    return step


# ---------------------------------------------------------------------------

def _environment() -> Dict[str, Any]:
    commit = subprocess.run(["git", "rev-parse", "HEAD"], cwd=REPO, capture_output=True, text=True).stdout.strip()
    dirty = bool(subprocess.run(["git", "status", "--porcelain", "--untracked-files=no"], cwd=REPO,
                                capture_output=True, text=True).stdout.strip())
    info: Dict[str, Any] = {
        "git_commit": commit, "worktree_dirty": dirty,
        "platform": platform.platform(), "python": sys.version.split()[0],
    }
    try:
        from lib.office import get_backend

        backend = get_backend()
        status = backend.probe()
        info["office_backend"] = {"name": backend.name, "available": status.available,
                                  "fidelity": status.fidelity, "reason": status.reason}
    except Exception as exc:  # 只是记录
        info["office_backend"] = {"error": f"{type(exc).__name__}: {exc}"}
    return info


def main(argv: Optional[List[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    default_label = "windows-word" if sys.platform == "win32" else "macos-word"
    parser = argparse.ArgumentParser(description="P6 F5 字体可移植性验收")
    parser.add_argument("--label", default=default_label, help=f"证据目录名，默认 {default_label}")
    parser.add_argument("--output", default="p6-font-check.json", help="结果文件名")
    parser.add_argument("--only", help="只运行指定步骤，逗号分隔，例如 F06,F08")
    parser.add_argument("--skip-tests", action="store_true", help="跳过 F01 离线全量")
    parser.add_argument("--skip-matrix", action="store_true", help="跳过 F10 WW 矩阵（耗时最长）")
    parser.add_argument("--skip-word-tests", action="store_true", help="跳过 F11 真实 Word 全量")
    parser.add_argument("--keep-work", type=Path, help="把构建产物复制到此处（不入库）")
    args = parser.parse_args(argv)

    work = Path(tempfile.mkdtemp(prefix="p6-font-check-"))
    runner = Runner(work)
    plan: List[Tuple[str, Callable[[], Step], bool]] = [
        ("F01", lambda: step_offline_tests(runner), args.skip_tests),
        ("F02", lambda: step_doctor_default(runner), False),
        ("F03", lambda: step_plan(runner), False),
        ("F04", lambda: step_doctor_missing(runner), False),
        ("F05", lambda: step_build_missing(runner, work), False),
        ("F06", lambda: step_build_fallback(runner, work), False),
        ("F07", lambda: step_fallback_unavailable(runner), False),
        ("F08", lambda: step_v3_fallback(runner, work), False),
        ("F09", lambda: step_regression(runner, work), False),
        ("F10", lambda: step_matrix(runner, args.label), args.skip_matrix),
        ("F11", lambda: step_word_tests(runner), args.skip_word_tests),
    ]
    only = {item.strip().upper() for item in args.only.split(",")} if args.only else None
    scrub = Scrubber(work)
    environment = _environment()
    steps = []
    try:
        for step_id, make, skipped in plan:
            if (only and step_id not in only) or skipped:
                steps.append({"id": step_id, "status": "skipped"})
                print(f"[{step_id}] skipped")
                continue
            print(f"[{step_id}] 运行中 ...", flush=True)
            step = make()
            steps.append(step.to_dict(scrub))
            print(f"[{step_id}] {step.status}  {step.title}")
            for item in step.checks:
                if not item["ok"]:
                    print(f"    未通过: {item['check']}  {scrub.text(item['detail'])}")
    finally:
        if args.keep_work:
            shutil.copytree(work, args.keep_work, dirs_exist_ok=True)
        shutil.rmtree(work, ignore_errors=True)

    counts: Dict[str, int] = {}
    for item in steps:
        counts[item["status"]] = counts.get(item["status"], 0) + 1
    result = scrub.data({
        "schema_version": 1,
        "label": args.label,
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "environment": environment,
        "summary": counts,
        "steps": steps,
    })
    out_path = HERE / args.label / args.output
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"\n结果: {counts}，已写入 {_short(str(out_path))}")
    return 0 if counts.get("failed", 0) == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
