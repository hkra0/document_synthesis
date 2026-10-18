# P7 实施记录：多 agent Skill 分发

> **文档状态**：实施中；分发与内容（第 2–4 节）已完成，agent 试用（第 5 节）未进行  
> **依据**：[Windows 支持与公开发布执行计划](windows-and-public-release-plan.md)第 12 节

## 1. 退出条件

1. `python scripts/sync_skills.py --check` 通过，测试检查副本与源文件逐字节一致。
2. 根目录有 `AGENTS.md`，`CLAUDE.md` 引用它。
3. Skill 写明平台与后端的判断方法、各平台的命令写法和排障方法，`description` 含中英文触发词。
4. 四次 agent 试用（Codex、Claude Code 各在 macOS 与 Windows 上一次）记录写入 `docs/acceptance/public-release/agent-trials.md`；每次都先运行了 plan/doctor，且没有越级声称页码已核验。

## 2. 发现目录核对（2026-10-18）

按各工具官方文档核对仓库级 Skill 目录：

| 工具 | 读取的仓库目录 | 来源 |
| --- | --- | --- |
| Claude Code | `.claude/skills/<名>/SKILL.md`（从启动目录向上到仓库根） | code.claude.com/docs/en/skills |
| Codex | `.agents/skills/<名>/SKILL.md`（从启动目录向上到仓库根）；文档不再提 `.codex/skills`；`name` 与 `description` 必填 | learn.chatgpt.com/docs/build-skills |
| GitHub Copilot | `.github/skills`、`.claude/skills`、`.agents/skills` | docs.github.com/en/copilot/concepts/agents/about-agent-skills |
| Cursor | `.agents/skills`、`.cursor/skills`，兼容读取 `.claude/skills`、`.codex/skills` | cursor.com/docs/context/skills |

结论：分发到 `.claude/skills/` 与 `.agents/skills/` 两处即可覆盖以上四个工具。删除旧的 `.codex/skills/`：Codex 文档已不再列出它，而 Cursor 仍会读取，保留会出现第三份重复。Copilot 与 Cursor 会同时看到两份同名 Skill，内容逐字节相同；是否会在列表中重复显示未验证。

国内的 agent 工具（如 Trae、通义灵码、CodeBuddy）未逐一核对。它们若支持 Agent Skills 约定，通常读取上述目录之一；不支持时可以读 `AGENTS.md`。

## 3. 实现

| 文件 | 内容 |
| --- | --- |
| `skills/document-synthesis/SKILL.md` | 唯一源文件，改写为中文。新增“命令写法”“平台与后端”（`--doctor` 退出码 0/2/1 对应能做的事）“字体”“草稿模式”，排障按 macOS/Windows 拆开；Windows 排障条目对应 `lib/office/win_com.py` 中实际输出的报错 |
| `scripts/sync_skills.py` | 把 `skills/*/SKILL.md` 复制到两个发现目录；`--check` 只检查，不一致时退出码 1。用复制而不是符号链接 |
| `.claude/skills/document-synthesis/SKILL.md`、`.agents/skills/document-synthesis/SKILL.md` | 生成的副本 |
| `AGENTS.md` | 入口命令、测试命令、禁止事项（不新增入口、输入只读、不得把 `draft` 说成 `exact`、不分发字体）、Skill 维护方法 |
| `CLAUDE.md` | 只有一行 `@AGENTS.md`，Claude Code 通过导入读取同一份内容 |
| `tests/test_repo_hygiene.py` | 新增 `SkillDistributionTest`：副本与源文件一致；源文件 frontmatter 含 `name`（与目录名相同）和 `description` |
| `README.md` | 项目结构与 Skill 位置改为新路径 |

## 4. 验证

- `python scripts/sync_skills.py --check` 通过；人为修改一份副本后报告不一致、退出码 1，重新同步后恢复。
- `tests.test_repo_hygiene` 6 项通过。

## 5. 未完成：agent 试用

退出条件第 4 条需要在全新会话中只给出自然语言请求（例如“用示例材料生成带目录的文档”），观察 agent 是否找到 Skill、是否先运行 plan/doctor、是否如实报告支持级别。四次试用都还没有进行，P7 在此之前不算完成。
