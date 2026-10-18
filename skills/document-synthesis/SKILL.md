---
name: document-synthesis
description: 在本仓库中规划、配置、构建并核验多来源 Word 材料汇编：把 PDF、Word、图片、演示文稿合并排版为带目录和页码的公文格式 docx，处理 profile/manifest、来源策略、格式包、字体、构建排障与交付 QA。触发词：Word、docx、目录、页码、公文、材料汇编、合并 docx/PDF、排版。Use for merging docx/PDF into a Word document with a table of contents and page numbers, official document formatting, build diagnosis and output QA in this repository; not for unrelated standalone document editing.
---

# 多源材料合成（document-synthesis）

在仓库根目录工作。本 Skill 覆盖仓库的统一引擎。差异能用 profile 或 manifest
表达时，不要新建项目专用的构建脚本，也不要把项目名写死进引擎。

## 命令写法

文中命令写作 `python synthesize.py ...`。macOS/Linux 上没有 `python` 时用
`python3`；Windows 上也可以用 `py`。

## 平台与后端：先看 `--doctor` 的退出码

| 退出码 | 含义 | 可以做的事 |
| --- | --- | --- |
| `0` | Python 依赖齐全，Word 自动化探针通过，可精确构建 | 构建并声称“页码已核验” |
| `2` | 依赖齐全，但没有可精确分页的 Word（Linux，或没装 Word 的 Windows/macOS） | 只能 `--plan`、格式分析、`smoke_test.py --structure-only`，或 `--allow-draft` 生成草稿 |
| `1` | 缺少核心依赖、输入无效，或后端配置无效 | 先修复，再重跑 `--doctor` |

- 只有退出码 `0` 且构建与 smoke 都通过时，才能说页码已核验。
- **不得把草稿（`draft`）说成精确结果（`exact`）**，也不得把 LibreOffice 的版面当作 Word 的等价结果。
- 页码只对生成它的 Word 环境有效；字体、Word 版本不同可能导致分页不同。
- 后端可用 `--office-backend {auto,word,none}` 或环境变量 `DOCUMENT_SYNTHESIS_OFFICE_BACKEND` 指定。

## 安全起步

1. 只读预览：`python synthesize.py --source path/to/source --plan`。`--plan` 不调用 Office，新来源、改配置或排障时优先用它。
2. 要构建时，先运行 `python synthesize.py --doctor --source path/to/source`，按上表判断能做什么。
3. 构建：`python synthesize.py --source path/to/source`（或 `--project 名称`、`--all`）。引擎在输出目录的 `.work/` 中完成工作，所有 QA 通过后才发布交付物；QA 失败不替换已有成果。
4. 核验：`python smoke_test.py` 加上与构建相同的 `--source`、`--manifest`、`--override` 和 `--font-fallbacks` 参数。

## 按来源形态配置

- `directory_tree`：普通的有序材料文件夹。
- `explicit_tree`：层级或顺序无法自动推断时使用。需要从 DOCX 读取标题、又不想重复正文时，用 `docx_outline` 节点。
- `highlighted_docx`：只用于用指定高亮颜色标出标题的单个 DOCX。

配置放在项目内的 `manifest.json`、可复用的 `profiles/项目名.json`，或显式 `--manifest`。改完用 `--plan` 验证。

## 字体

- `--plan` 和 `--doctor` 列出格式要求的每个字体是否已安装。缺失时 Word 会替换字体，分页可能变化；缺失不改变 `--doctor` 的退出码。
- 需要替代字体时，构建和 smoke 都加 `--font-fallbacks`（或 `DOCUMENT_SYNTHESIS_FONT_FALLBACKS=1`）。默认不替换，只替换本机确实缺少的字体；候选都没装时构建失败。
- 仓库不分发字体文件。不要替用户下载来源不明的字体。

## 草稿模式（没有 Word 时）

`--allow-draft` 生成带 `_draft` 后缀的草稿：目录页码是 `PAGEREF` 域，打开时显示“?”，需要用户在 Word 中等排版完成后全选并按 F9 更新。向用户交付草稿时必须说明这一点，并说明页码未经核验。能精确分页时 `--allow-draft` 不起作用。

## 构建与排障

`synthesize.py` 是唯一构建入口。输入只读。输出文件名必须留在输出目录内，并使用 `.docx` 扩展名。重试构建前先看 `--doctor` 的结果。

**macOS**

- “未获准控制 Microsoft Word”：在“系统设置 → 隐私与安全性 → 自动化”中允许运行构建的宿主应用控制 Word。完全磁盘访问权限只管文件访问，不能代替自动化权限。
- 报告“受限上下文”：改从能访问 macOS 图形自动化服务的本机终端运行。
- 导出时出现 “Grant File Access”：需要用户确认授予输出目录访问权限；导出和最终 QA 完成前，不能声称分页验收通过。

**Windows**

- “无法启动 Word”：先手动打开一次 Word，完成首次运行、登录与激活；终端以管理员身份运行而 Office 不是时也会出现。
- “未激活或处于功能受限模式”：先完成 Office 激活。
- “以受保护视图打开了文件”：在文件属性中“解除锁定”，或在信任中心调整受保护视图设置。
- “正忙或被对话框阻塞”：关闭 Word 中打开的对话框后重试。
- 偶发挂起或超时：第三方 COM 插件（如 iSlide、Acrobat PDFMaker）会加载到自动化实例中；超时后只清理本次启动的 Word。频繁发生时在 Word 的 COM 加载项中禁用相关插件。超时默认 600 秒，可用 `DOCUMENT_SYNTHESIS_OFFICE_TIMEOUT` 调整。

## 格式学习（v3）

格式学习同样只用 `synthesize.py`：`--analyze-format` 生成 `analysis.json`、`review.html` 和草稿 `decisions.example.json`；编译时必须给出明确的最终决定，不能使用仍含 `pending` 的草稿。编译格式包或目标 RoleMap 时保留报告 ID 和源文件 SHA-256。`--render-preview` 生成虚构内容的近似预览，`--migrate-manifest` 把 v1/v2 清单迁移到单独的输出文件，不覆盖输入。

## 验证与发布

改动来源策略、安全或配置代码后运行：

```bash
python -m unittest discover -s tests -v
```

交付代码改动前运行 `git diff --check`。发布前，每项声称的能力都要对应测试 ID、支持级别和最近的证据；在全新输出目录中运行示例的 `--plan`、`--doctor`、构建和 `smoke_test.py`，之后删除 `.work`、Office 锁文件、渲染缓存和 Python 缓存。不要把某次真实 Word 结果推广到它的夹具、环境和格式边界之外。

## 安全演示

`examples/minimal-demo/` 是完全虚构的演示项目，不含人名、地点和单位。它的 README 给出只读命令和重新生成命令。把它当作冒烟路径，不要当作处理敏感材料的模板。

修改本 Skill 时只改 `skills/document-synthesis/SKILL.md`，然后运行 `python scripts/sync_skills.py` 同步到各 agent 目录。
