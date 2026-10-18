# AGENTS.md

本仓库把 PDF、Word、图片和演示文稿合并排版为带目录和页码的公文格式 Word 文档。完整操作说明见 Skill：[`skills/document-synthesis/SKILL.md`](skills/document-synthesis/SKILL.md)。下面是不读 Skill 时也必须遵守的最低要求。

## 入口命令

- `synthesize.py` 是唯一入口：`--plan`（只读预览，不调用 Office）、`--doctor`（环境检查）、构建、格式学习都通过它。
- `smoke_test.py` 核验交付物，参数与构建时相同（`--source`、`--manifest`、`--override`、`--font-fallbacks`）。
- 命令写作 `python ...`；macOS/Linux 上没有 `python` 时用 `python3`，Windows 上也可以用 `py`。

处理用户材料时依次运行 `--plan` → `--doctor` → 构建 → `smoke_test.py`。

## 测试命令

```bash
python -m unittest discover -s tests -v
python -m unittest tests.test_repo_hygiene -v
python scripts/sync_skills.py --check
```

含真实 Word 的用例需要设置 `DOCUMENT_SYNTHESIS_WORD_TEST=1`。跳过的测试不算通过。

## 禁止事项

- 不新增入口脚本，不把项目名写死进引擎；差异用 profile 或 manifest 表达。
- 输入材料只读；输出只写到输出目录内。
- `--doctor` 退出码不是 `0` 时，不得声称页码已核验。**不得把草稿（`draft`）说成精确结果（`exact`）**，也不得把 LibreOffice 的版面当作 Word 的等价结果。
- 不分发字体文件，不替用户下载来源不明的字体。
- 不在仓库中写入个人路径、邮箱或真实材料内容。

## Skill 维护

Skill 只在 `skills/document-synthesis/SKILL.md` 中修改，然后运行 `python scripts/sync_skills.py`，把它复制到 `.claude/skills/`（Claude Code）和 `.agents/skills/`（Codex）。不要直接改这两份副本；`tests/test_repo_hygiene.py` 会检查副本与源文件逐字节一致。
