# Windows 支持与公开发布执行计划

日期：2026-10-09；状态：**计划中，尚未实施**。本轮批次编号为 **P0–P9**，不覆盖历史 R/N/S/NW 批次记录。下文的新模块、函数、参数、测试和证据文件均为实施设计，在对应批次验证之前，不能写成已经具备的能力。

依据：当前 `main`（`c2c1a30`）的生产代码、[README](../README.md)、[DEVELOPMENT.md](../DEVELOPMENT.md)、项目 Skill [SKILL.md](../.codex/skills/document-synthesis/SKILL.md)。2026-10-09 在 macOS、Python 3.14 下运行离线全套，结果为 325 项通过、6 项条件跳过；尚未在 Windows 上实际运行。

## 1. 目标与验收口径

### 1.1 目标

1. **Windows 能使用大部分功能**：在装有 Microsoft Word 的 Windows 10/11 上，完成与 macOS 等价的完整构建，包括精确页码、目录回填和页级 QA。没有 Word 的 Windows 或 Linux 至少能完成全部只读、分析和结构类功能，并能生成明确标为“页码未验证”的草稿交付物。
2. **可以直接公开发布**：陌生用户只读 README，就能在三个平台上安装、跑通示例；仓库不含本机路径或个人信息；有 CI、版本号、变更记录、安全与贡献说明；主流 agent（Codex、Claude Code 及其他支持 `SKILL.md` 的工具）能发现并正确使用项目 Skill。

### 1.2 能力分级

所有平台能力统一用三个支持级别描述。README、`--doctor` 输出和状态记录都使用同一套词汇：

| 级别 | 含义 | 允许的表述 |
| --- | --- | --- |
| `exact` | 通过本机 Microsoft Word 导出，完成书签页码读取、目录回填和页级 QA | “页码已核验” |
| `draft` | 无 Word，或使用 LibreOffice 等非 Word 引擎；结构与文字核验通过，页码和版式未经 Word 核验 | “草稿，页码未验证”，不得写成通过 |
| `unsupported` | 当前环境无法执行 | 明确报错并给出替代路径 |

### 1.3 目标功能矩阵（实施完成后）

| 功能 | macOS + Word | Windows + Word | Windows/Linux 无 Word |
| --- | --- | --- | --- |
| `--plan`、`--doctor` | 支持 | 支持（P2 起） | 支持 |
| 格式分析、编译、RoleMap、迁移、HTML 预览 | 支持 | 支持（P1 验证） | 支持 |
| `.docx`/PDF/图片合并、样式统一 | 支持 | 支持 | 支持 |
| `.doc` → `.docx` 转换 | Word | Word COM（P3） | LibreOffice 可选（P5），否则 `unsupported` |
| `.pptx` → PDF 转换 | PowerPoint | PowerPoint COM（P3） | LibreOffice 可选（P5），否则 `unsupported` |
| 精确页码、目录回填、页级 QA | `exact` | `exact`（P3/P4） | `draft`（P5） |
| `smoke_test.py --structure-only` | 支持 | 支持 | 支持 |

“大部分功能”的验收以此表为准：Windows + Word 的每一格都必须有真实运行证据；无 Word 一列不得出现 `exact`。

## 2. 现状评估

### 2.1 已经跨平台的部分

- 依赖库（python-docx、PyMuPDF、pypdf、Pillow、numpy、lxml、jsonschema）在 Windows 上都有预编译 wheel。
- 文本文件读写已显式使用 UTF-8；JSON 输出使用 `ensure_ascii=False`。
- 配置中的相对路径使用 `as_posix()` 规范化；发布事务使用 `os.replace`。
- 示例 docx 的作者元数据为空或为 `python-docx`，不含个人信息。
- `smoke_test.py --structure-only` 已提供无 Office 的核验路径。

### 2.2 只能在 macOS 运行的部分

全部 Office 调用都直接拼接 AppleScript，并通过 `osascript` 执行，没有抽象层：

| 位置 | 作用 | 平台假设 |
| --- | --- | --- |
| `lib/qa.py:23` `_word_access_directory` | 固定 Word 文件访问目录 | macOS 沙盒授权 |
| `lib/qa.py:44` `word_export_status` | 静态可用性检查 | 非 darwin 直接返回不可用；检查 `/usr/bin/osascript`、`/Applications/Microsoft Word.app` |
| `lib/qa.py:77` `word_automation_status` | `--doctor` 探针 | `osascript` |
| `lib/qa.py:102` `export_docx_to_pdf` | DOCX → PDF | AppleScript `save as ... format PDF` |
| `lib/qa.py:188` `get_exact_printed_heading_pages` | 旧路径：从 PDF 文本反查页码 | 依赖 `export_docx_to_pdf` |
| `lib/pagination.py:92` `_inspection_script`、`:232` `inspect_document` | 书签物理页、打印页码，并导出 PDF | AppleScript `get range information` |
| `lib/engine.py:216-285` `_run_applescript`、`prepare_conversions` | `.doc` 转换、`.pptx` 转 PDF | Word/PowerPoint AppleScript |
| `lib/preview.py:119` | 预览时检查 Word | 同 `word_export_status` |
| `synthesize.py:125` | `--doctor` | 同上 |

结果：在 Windows 上，`--plan`、分析类命令和 `--structure-only` 大概可用；任何需要目录和页码的构建都会在 `inspect_document` 处抛出 `OfficeExportError` 并停止。这是项目的核心路径。

### 2.3 Windows 上的其他风险（需验证，未确认）

| 编号 | 风险 | 说明 |
| --- | --- | --- |
| R1 | 控制台编码 | 输出大量中文。在 cp936 控制台或重定向到文件时，`print` 可能抛出 `UnicodeEncodeError` |
| R2 | 文件锁 | Word 打开目标文件时，`os.replace` 会报 `PermissionError`；`~$` 锁文件、杀毒软件扫描也会造成短暂占用 |
| R3 | 输出文件名 | 需要拒绝 `CON`/`NUL`/`COM1` 等保留名、结尾的点和空格、`<>:"\|?*`，以及仅大小写不同的重名 |
| R4 | 长路径 | 中文深层目录加 `.work/run-*` 可能超过 260 字符 |
| R5 | 扩展名大小写 | 当前 `rglob("*.doc")` 等匹配在不同平台上的大小写敏感性不同，`.DOC`、`.PPTX` 可能被漏掉或重复处理 |
| R6 | 换行符 | 没有 `.gitattributes`；Windows 签出后，JSON 和哈希证据可能因 CRLF 改变 |
| R7 | 字体 | 预设使用 `仿宋_GB2312`、`楷体_GB2312`、`方正小标宋简体`。新版 Windows 默认只有 `仿宋`、`楷体`，没有 `_GB2312` 版和方正字体，Word 会替换字体并改变分页 |
| R8 | 测试假设 | 测试中可能存在 POSIX 路径字面量、`/private/tmp`、符号链接测试（Windows 默认无权创建） |
| R9 | 命令名 | 文档和 Skill 中全部写成 `python3`；Windows 上通常是 `python` 或 `py` |

### 2.4 公开发布的阻碍

| 编号 | 问题 | 位置 |
| --- | --- | --- |
| G1 | 被跟踪的证据文件含本机绝对路径 `<user-home>/...` | `docs/acceptance/n0-n10-review/*.json`、`pdf-diagnostics/*.txt`、`post-review/pdf-diagnostics/*` 等 |
| G2 | 多处文档写死 `/private/tmp/document-synthesis-word-access` | README、DEVELOPMENT.md、`post-review/*` |
| G3 | README 只有中文，开头就是内部批次术语（R/N/S/NW），新用户难以理解 | README.md |
| G4 | 没有 CI、`.gitattributes`、CHANGELOG、SECURITY、CONTRIBUTING、版本号 | 仓库根目录 |
| G5 | Skill 只在 `.codex/skills/`，其他 agent 不会自动发现；没写平台限制；命令写死 `python3` | SKILL.md |
| G6 | PyMuPDF 是 AGPL-3.0 或商业双许可，与本项目 MIT 并存，但未说明 | README、LICENSE |
| G7 | 验收证据约 5 MB，占仓库大部分；`nw09-build-metadata.json` 单个约 1.5 MB | `docs/acceptance/` |
| G8 | README 声明支持 Python 3.8（已于 2024 年停止维护），但没有 CI 验证 | README、requirements.txt |
| G9 | 工作目录中有未跟踪的大文件和目录（`.test_export_clean.pdf`、`tmp_test_full_qa.pdf`、`修改/`、`scratch/`） | 本地 |

## 3. 总体设计：Office 后端抽象

新增 `lib/office/` 包，把“Office 能做什么”与“如何调用 Office”分开。业务代码只依赖接口，不再出现 `osascript`、`sys.platform` 或应用路径。

```
lib/office/
├── __init__.py          # get_backend()，按平台与配置选择后端
├── base.py              # OfficeBackend 协议、数据类型、错误类型
├── mac_applescript.py   # 由现有 AppleScript 迁移而来，行为不变
├── win_com.py           # Windows：pywin32 驱动 Word/PowerPoint COM
├── libreoffice.py       # 可选：soffice --headless，只产出 draft
└── null.py              # 无 Office：明确返回 unsupported
```

### 3.1 接口

```python
class OfficeBackend(Protocol):
    name: str                      # "mac-word" | "win-word" | "libreoffice" | "none"
    fidelity: str                  # "exact" | "draft" | "unsupported"

    def static_status(self) -> BackendStatus: ...        # 不弹授权、不启动应用
    def probe(self) -> BackendStatus: ...                # --doctor 用，只读探针
    def convert_doc_to_docx(self, src: Path, dst: Path) -> None: ...
    def convert_pptx_to_pdf(self, src: Path, dst: Path) -> None: ...
    def export_pdf(self, docx: Path, pdf: Path) -> None: ...
    def inspect_bookmarks(self, docx: Path, pdf: Path,
                          names: Sequence[str]) -> BookmarkReport: ...
```

`BookmarkReport` 包含每个书签的物理页和打印页码、总页数、导出的 PDF 路径，以及 `fidelity`。`lib/pagination.py` 中现有的 `_parse_page_map`、页码标签抽取、`page_records_from_map` 保持不变，只把“取得原始报告”这一步换成调用后端。

### 3.2 后端选择

1. 环境变量 `DOCUMENT_SYNTHESIS_OFFICE_BACKEND` 或命令行参数 `--office-backend {auto,word,libreoffice,none}` 显式指定。
2. `auto`：macOS → `mac-word`；Windows → `win-word`；都不可用时 → `none`。**LibreOffice 永远不会被自动选中**，必须显式选择，避免把草稿误当成精确结果。
3. 构建需要 `exact` 而后端只能提供 `draft` 时，默认失败。只有加 `--allow-draft` 才生成草稿，并在文件名、`build-diagnostics.json` 和终端输出中标注。

### 3.3 Windows COM 实现要点（`win_com.py`）

| 操作 | COM 调用 | 对应现有 AppleScript |
| --- | --- | --- |
| 启动隔离实例 | `win32com.client.DispatchEx("Word.Application")`，`Visible=False`，`DisplayAlerts=0` | `set display alerts to none` |
| 打开 | `Documents.Open(path, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False, Visible=False)` | `open POSIX file` |
| 重新分页 | `doc.Repaginate()` | `repaginate` |
| 书签页码 | `r = doc.Range(bm.Range.Start, bm.Range.Start)`；`r.Information(3)` 取物理页（`wdActiveEndPageNumber`），`r.Information(1)` 取打印页码（`wdActiveEndAdjustedPageNumber`） | `get range information ... active end page number` / `adjusted page number` |
| 导出 PDF | `doc.ExportAsFixedFormat(pdf, 17)`（`wdExportFormatPDF`），并设置 `CreateBookmarks`，与 macOS 输出对齐 | `save as ... format PDF` |
| `.doc` 转换 | `doc.SaveAs2(dst, FileFormat=16)`（`wdFormatDocumentDefault`） | `file format document default` |
| `.pptx` 转 PDF | `PowerPoint.Application` → `Presentations.Open(src, ReadOnly=True, WithWindow=False)` → `SaveAs(dst, 32)` | PowerPoint AppleScript |
| 关闭 | `doc.Close(SaveChanges=0)`；`finally` 中 `app.Quit()` | `close saving no` |

可靠性要求：

- **超时**：COM 调用是同步的，无法从内部中断。每次检查都在独立子进程中运行（`python -m lib.office.win_com_worker`），父进程按 600 秒超时结束子进程，并按 PID 清理遗留的 `WINWORD.EXE`。不得用 `taskkill /IM WINWORD.EXE` 关掉用户自己打开的 Word。
- **线程**：worker 入口调用 `pythoncom.CoInitialize()`，退出时调用 `CoUninitialize()`。
- **常量**：使用数值常量，不依赖 `gencache.EnsureDispatch` 生成的缓存，避免 `gen_py` 缓存损坏导致的常见故障。
- **路径**：传给 COM 的路径一律 `Path.resolve()` 后转为 Windows 绝对路径；超过 259 字符时，先复制到短的工作目录再处理。
- **受保护视图**：来自网络的文件会以受保护视图打开，导致 `Documents.Open` 失败。工作目录中的副本需要移除 Zone.Identifier ADS，并把对应错误转换为可操作的提示。
- **错误翻译**：常见 HRESULT（Word 未安装、未激活、许可证对话框、对象被占用 `0x8001010A`）翻译成中文排障提示，与现有 `_automation_failure_message` 风格一致。
- `pywin32` 只在 Windows 上安装：`requirements.txt` 增加 `pywin32>=306; sys_platform == "win32"`。

### 3.4 LibreOffice 草稿后端（可选，`libreoffice.py`）

- 用 `soffice --headless --convert-to` 完成 `.doc` → `.docx`、`.pptx` → PDF、DOCX → PDF。
- 书签页码：导出 PDF 时保留书签（`ExportBookmarks`），再用 PyMuPDF 读取 outline 或命名目标对应的页码。打印页码从 PDF PageLabels 读取，读不到时标为 `null`，不得推算。
- 每个产物和报告都写 `fidelity: "draft"`；QA 规则与 `exact` 相同，但结论只能写“草稿检查通过”。
- 对 `soffice` 设置独立的 `-env:UserInstallation`，避免与用户正在运行的 LibreOffice 冲突。

## 4. 批次、依赖与交付物

| 批次 | 内容 | 依赖 | 主要交付物 |
| --- | --- | --- | --- |
| P0 | 发布前清理与基线固定 | 无 | 路径脱敏、`.gitattributes`、基线测试数 |
| P1 | 无 Office 跨平台加固与三平台 CI | P0 | 编码、文件名、锁、扩展名修复；GitHub Actions 矩阵 |
| P2 | Office 后端抽象，macOS 行为不变 | P1 | `lib/office/`、mac 后端、`none` 后端、契约测试 |
| P3 | Windows Word/PowerPoint COM 后端 | P2 | `win_com.py`、worker 子进程、doctor |
| P4 | Windows 真实 Word 验收矩阵 | P3 | WW01–WW12 证据、与 macOS 页码对照 |
| P5 | 草稿模式与可选 LibreOffice 后端 | P2 | `--allow-draft`、`libreoffice.py`、标注规则 |
| P6 | 字体可移植性 | P2；验收依赖 P4 | 字体检查、回退映射、doctor 提示 |
| P7 | 多 agent Skill 分发 | P1；内容依赖 P3/P5 | 平台感知 SKILL.md、多目录分发、AGENTS.md |
| P8 | 公开发布文档与仓库治理 | P0–P7 | 英文 README、CHANGELOG、SECURITY、CONTRIBUTING、版本号 |
| P9 | 发布候选全量回归与发布 | 全部 | 三平台证据、`v1.0.0` 标签与 Release |

建议顺序：P0 → P1 → P2 → (P3 → P4) ∥ P5 → P6 → P7 → P8 → P9。P5 只依赖 P2，可以与 P3/P4 并行。

沿用既有原则：每批按“失败断言 → 实现 → 对照回归 → 证据”提交；`skip` 不计为通过；`blocked`/`partial` 不写成 `verified`；不以测试数量或日期代替退出条件。

## 5. P0：发布前清理与基线

### 执行动作

1. **路径脱敏（G1）**：新增 `docs/acceptance/tools/sanitize_paths.py`，把被跟踪证据中的仓库绝对路径改写为 `<repo>/...`，并把 `/private/tmp/...` 改写为 `<word-access-dir>/...`。改写后，凡是记录了内容哈希的文件，都要重新计算并更新对应的 `evidence-hashes.json`；文件内容中记录的源文件 SHA-256 不受影响，不得改动。
2. 以后生成证据的脚本（`write_status.py`、`diagnose_pdfs.py` 等）统一调用同一个 `relativize()`，从源头避免再次写入绝对路径。
3. **防回归检查**：新增 `tests/test_repo_hygiene.py`，扫描 `git ls-files` 的文本文件，发现开发机用户主目录路径、邮箱地址时失败（`LICENSE` 版权行除外）。
4. 新增 `.gitattributes`：`* text=auto eol=lf`；`*.json`、`*.md`、`*.py` 设为 `eol=lf`；`*.docx`、`*.png`、`*.pdf` 设为 `binary`。
5. 清理本地未跟踪文件（G9），在 `.gitignore` 中补上 `修改/`、`tmp_*.pdf`、`.test_*`。
6. 记录基线：离线全套 325 项通过（6 项跳过），以及当前 macOS 真实 Word 结果，写入 `docs/acceptance/public-release/p0-baseline.json`。

### 退出条件

`test_repo_hygiene` 通过；全局搜索用户主目录绝对路径无结果；脱敏后离线全套仍与基线一致；脱敏前后的证据只有路径字段不同（附 diff 摘要）。

## 6. P1：无 Office 部分的跨平台加固与 CI

### 执行动作

1. **控制台编码（R1）**：在 `synthesize.py` 和 `smoke_test.py` 入口，如果 `sys.stdout.encoding` 不是 UTF-8，就调用 `sys.stdout.reconfigure(encoding="utf-8", errors="replace")`，`stderr` 同理。子进程调用统一加 `encoding="utf-8", errors="replace"`，不依赖 `text=True` 的本地默认编码。
2. **输出文件名（R3）**：在 `lib/config.py` 的输出名校验中加入 Windows 保留名、非法字符、结尾点或空格、大小写重名检查，**所有平台都执行**，保证在 macOS 上写的配置拿到 Windows 也能用。补成对的正反测试。
3. **文件锁（R2）**：`engine.py:1109` 的发布事务和 `pagination.py:279` 的 `os.replace` 遇到 `PermissionError` 时，有限次退避重试（例如 5 次，共约 2 秒）；仍失败时提示“目标文件可能正在 Word 中打开”，并保持现有回滚语义。补一个用打开的文件句柄模拟占用的测试（只在 Windows CI 运行）。
4. **扩展名（R5）**：扫描和转换统一用 `path.suffix.lower()` 判断，不用大小写敏感的 glob 模式；补 `.DOC`、`.Pptx` 测试。
5. **长路径（R4）**：工作目录名保持短（`run-<8位>`）；doctor 在 Windows 上检查 `LongPathsEnabled`，并在预计路径超过 240 字符时预警。
6. **测试可移植（R8）**：清点 `tests/` 中的 POSIX 路径字面量和 `/private/tmp`，改用 `tempfile`/`Path`；符号链接测试在无权限时用 `skipUnless(can_symlink())` 跳过，并在状态中单独计数。
7. **Python 版本（G8）**：最低版本定为 3.9 或 3.10（取决于 CI 实测），更新 README 和 `requirements.txt`，同时考虑加 `pyproject.toml` 的 `requires-python`。
8. **CI**：新增 `.github/workflows/ci.yml`，在 `ubuntu-latest`、`macos-latest`、`windows-latest` 上，用最低支持版本和最新稳定版 Python 运行：
   - `python -m unittest discover -s tests -v`
   - 用最小示例跑 `--plan`、`--doctor`（预期报告 Office 不可用，退出码约定见 P2）、`smoke_test.py --structure-only`
   - `test_repo_hygiene`
   - Windows 任务额外设置 `PYTHONIOENCODING` 为空，并使用 `chcp 936`，专门复现 R1

### 退出条件

三平台 CI 全绿；Windows 上的跳过项逐项登记原因，且都不属于“本应可移植”的功能；最小示例和中文示例的 `--plan` 输出在三平台上一致（路径分隔符规范化后比较）。

## 7. P2：Office 后端抽象（macOS 行为不变）

### 执行动作

1. 建立 `lib/office/` 和第 3.1 节的接口。把 `qa.py`、`pagination.py`、`engine.py` 中的 AppleScript **原样迁移**到 `mac_applescript.py`，不顺手修改脚本。
2. `word_export_status`、`word_automation_status`、`export_docx_to_pdf`、`inspect_document`、`prepare_conversions` 改为调用 `get_backend()`。保留旧函数名作为薄包装，避免同时修改大量测试。
3. `_word_access_directory` 移到 mac 后端内部；其他后端使用 `tempfile` 下的短目录。
4. `null.py`：`static_status` 返回 `unsupported`，原因按平台给出（“未检测到 Microsoft Word；Windows 需安装桌面版 Word；Linux 可使用 `--office-backend libreoffice --allow-draft` 生成草稿”）。
5. 新增 `tests/test_office_backends.py`：用假后端驱动完整构建路径，覆盖成功、超时、书签缺失、导出 PDF 缺失、`draft` 被拒绝这五种情况。
6. **`--doctor` 退出码**：`0` 表示可进行 `exact` 构建，`2` 表示只能 `draft` 或分析类功能，`1` 表示缺少依赖。写入 README，CI 用它断言。
7. 废弃 `qa.get_exact_printed_heading_pages`（靠 PDF 文本匹配推页码的旧路径）：先确认它是否仍被调用；如仍被调用，改走书签报告；否则标记弃用，下一个次版本删除。

### 退出条件

除导入路径外，macOS 离线全套与 P0 基线一致；macOS 真实 Word 回归（NW01–NW12 和 `remediation_word`）全部重新通过，页码记录与基线逐项相同；生产代码中 `osascript`、`sys.platform == "darwin"` 只出现在 `lib/office/mac_applescript.py` 和 `get_backend()` 中（用测试断言）。

## 8. P3：Windows COM 后端

### 执行动作

1. 按第 3.3 节实现 `win_com.py` 和 `win_com_worker.py`。父子进程之间用 JSON 通过 stdin/stdout 通信，worker 只输出一行结果 JSON，日志走 stderr。
2. `static_status`：通过注册表 `HKCR\Word.Application\CLSID` 判断是否安装，不启动 Word。`probe`：启动隔离实例，读取 `Application.Version` 后退出；探测 Word 是否被激活，以及是否卡在首次运行或许可证对话框。
3. 在一个 Windows 后端内部实现 `.doc` 和 `.pptx` 转换；PowerPoint 缺失时只让 `.pptx` 转换报 `unsupported`，不影响其他功能。
4. 与 macOS 对齐导出参数：页面范围为全部，`CreateBookmarks=1`（按标题），`DocStructureTags=True`，`BitmapMissingFonts=True`。差异写进 `docs/office-backends.md`。
5. 新增 `--doctor` 的 Windows 专属输出：Word 版本与位数、是否为 Click-to-Run 安装、PowerPoint 是否可用、长路径设置、字体检查（P6）。
6. 本地单元测试使用 COM 假对象；真实 Word 测试按现有约定用环境变量开启：`DOCUMENT_SYNTHESIS_WORD_TEST=1`。

### 退出条件

在至少一台 Windows 11 + Microsoft 365 Word（64 位）上，最小示例、中文示例和自定义格式示例都能用 `exact` 完成构建和 `smoke_test.py`；强制结束 Word 进程、目标文件被占用、书签缺失、超时四种故障都能得到可读的报错，且不残留 `WINWORD.EXE` 和半成品交付物。

## 9. P4：Windows 真实 Word 验收矩阵

### 执行动作

1. 参照 NW01–NW12 建立 WW01–WW12，使用同一组匿名夹具，在 Windows 上按 plan → doctor → build → smoke 串行执行。
2. **跨平台页码对照**：对同一夹具，比较 macOS Word 与 Windows Word 的书签物理页和打印页码。在字体完全一致的前提下，**页码应该相同**；出现差异时逐项记录原因（字体替换、Word 版本、打印机驱动或默认纸张），不能直接接受为“正常差异”。
3. 记录环境：Windows 版本、Word 版本号与通道、默认打印机、已安装的相关字体清单（只记名称），写入 `docs/acceptance/public-release/windows-word/environment.json`。
4. 至少再覆盖一个较旧版本（例如 Office 2019 或 2021），作为兼容下限；无法取得时，在 README 的支持范围中如实写“仅验证过 Microsoft 365”。

### 退出条件

WW01–WW12 全部 `verified`（`skip` 不算）；跨平台页码差异为 0，或每一项差异都有已确认的原因与说明；README 的平台矩阵只列出已经验证的 Word 版本。

## 10. P5：草稿模式与可选 LibreOffice 后端

### 执行动作

1. 新增 `--allow-draft`：在没有 `exact` 后端时，仍生成交付物，但目录页码只来自草稿后端或留空占位，并在以下位置标注：
   - 文件名后缀：`*_draft.docx`（可配置，但不能关闭标注）
   - `build-diagnostics.json` 的 `fidelity: "draft"` 与原因
   - 终端最后一行：“草稿：页码与版式未经 Microsoft Word 核验”
2. `none` 后端加 `--allow-draft`：目录写入 Word 原生 TOC 域并设置 `updateFields`，由用户打开 Word 时更新；不填写推算的页码。
3. 按第 3.4 节实现 `libreoffice.py`；doctor 检查 `soffice` 的路径和版本（Windows 上查找 `Program Files\LibreOffice\program\soffice.exe`）。
4. 测试：草稿产物永远不会被标成 `exact`；没有 `--allow-draft` 时，构建必须失败，且已有交付物不被替换。

### 退出条件

在 Linux CI 上安装 LibreOffice，用最小示例跑通 `--office-backend libreoffice --allow-draft`；所有草稿标注都能被测试检查到；README 中关于草稿的表述与第 1.2 节一致。

## 11. P6：字体可移植性

### 执行动作

1. 新增 `lib/fonts.py`：枚举系统已安装字体。Windows 读取注册表 `HKLM`/`HKCU` 下的 `...\Windows NT\CurrentVersion\Fonts`；macOS 和 Linux 用 `fc-list`，没有时扫描常见字体目录。
2. `--plan` 和 `--doctor` 列出格式包需要但系统缺少的字体，并提示：Word 会替换字体，分页可能与预期不同，`exact` 结果仅对当前字体环境有效。
3. 格式包增加可选的 `font_fallbacks`（例如 `仿宋_GB2312 → 仿宋`，`楷体_GB2312 → 楷体`）。默认**不自动替换**，需要在格式包或命令行中显式启用，避免悄悄改变公文格式。
4. 在 README 中说明公文常用字体的合法获取方式，不在仓库中分发字体文件。

### 退出条件

在缺少 `仿宋_GB2312` 的 Windows 上，doctor 明确报告缺失；启用回退后构建通过，并在 `build-diagnostics.json` 中记录实际使用的回退映射。

## 12. P7：多 agent Skill 分发

### 执行动作

1. **只维护一份内容**：把 `SKILL.md` 的正文放在 `skills/document-synthesis/SKILL.md`，作为唯一源文件。
2. **分发到各 agent 的发现目录**：新增 `scripts/sync_skills.py`，把源文件**复制**（不用符号链接，Windows 签出时符号链接不可靠）到：
   - `.codex/skills/document-synthesis/SKILL.md`（Codex）
   - `.claude/skills/document-synthesis/SKILL.md`（Claude Code）
   - `.agents/skills/document-synthesis/SKILL.md`（其他采用通用约定的工具）

   `tests/test_repo_hygiene.py` 同时检查这几份副本与源文件逐字节一致，防止内容漂移。实施时再按各工具的最新官方文档核对发现路径，不在本计划中写死未经核实的路径。
3. 新增根目录 `AGENTS.md`，并让 `CLAUDE.md` 引用它：说明入口命令、测试命令、禁止事项（不新增入口脚本、输入只读、不得把 `draft` 说成 `exact`），给不读 Skill 的 agent 提供最低限度的指引。
4. **Skill 内容修改**：
   - 增加“平台与后端”一节：先看 `--doctor` 的退出码，`0` 才能声称页码已核验，`2` 只能做分析或 `--allow-draft`。
   - 命令写成 `python synthesize.py ...`，并注明：macOS/Linux 上如果 `python` 不可用，就用 `python3`；Windows 上可用 `py`。
   - 把排障内容按平台拆开：macOS 的 Automation 与文件访问授权；Windows 的 Word 激活、受保护视图、文件占用。
   - `description` 加上触发词：Word、目录、页码、公文、合并 docx/PDF，以及对应英文，方便自动触发。
5. 用 Codex 和 Claude Code 各在 macOS 与 Windows 上做一次端到端试用：只给出“用示例材料生成带目录的文档”这样的自然语言请求，记录 agent 是否找到 Skill、是否先运行 plan/doctor、是否如实报告支持级别。

### 退出条件

`sync_skills.py --check` 通过；四次 agent 试用记录写入 `docs/acceptance/public-release/agent-trials.md`；每次都先运行了 plan/doctor，且没有越级声称页码已核验。

## 13. P8：公开发布文档与仓库治理

### 执行动作

1. **README 重构（G3）**：
   - 新增 `README.en.md`，并在两份 README 顶部互相链接。
   - 开头依次写：一句话介绍、效果截图（来自中文示例）、三平台支持矩阵（第 1.3 节）、五分钟快速开始（每个平台一组命令）。
   - 把内部批次与验收描述（R/N/S/NW/WW 与测试数量）移到 `docs/acceptance/README.md`，主 README 只保留一行链接。
   - 删除写死的 `/private/tmp/...`（G2），改为“macOS 默认访问目录，可通过 `DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR` 修改”。
2. **许可说明（G6）**：新增 `THIRD_PARTY_NOTICES.md`，列出依赖及其许可，重点说明 PyMuPDF 的 AGPL-3.0 或商业双许可对再分发和网络服务场景的影响。本项目代码仍为 MIT。
3. **治理文件**：`CHANGELOG.md`（Keep a Changelog 格式，从 `1.0.0` 开始）、`SECURITY.md`（漏洞报告方式；说明工具会处理不可信 docx，以及 zip 炸弹和 XML 实体防护的现状）、`CONTRIBUTING.md`（开发环境、测试、证据规则、Skill 同步）、`.github/ISSUE_TEMPLATE/`（附 `--doctor` 输出的 bug 模板）、`.github/pull_request_template.md`。
4. **版本与打包**：新增 `pyproject.toml`（项目元数据、依赖、`requires-python`，可选的 `document-synthesis` 命令行入口），在 `lib/__init__.py` 中定义 `__version__`，并在 `--doctor` 中输出。本轮不发布到 PyPI，只保证 `pip install .` 可用；是否上 PyPI 留到发布后再决定。
5. **证据体积（G7）**：保留各批次的 `status.json`、摘要和哈希；1 MB 以上的原始元数据压缩后作为 GitHub Release 附件，或移到独立的 `evidence` 分支，并在原位置保留哈希与下载说明。注意：这只影响以后的克隆，历史提交中的文件仍然存在，**不改写已公开的历史**。
6. **安全复核**：确认解析 docx 时使用的 lxml 解析器禁用了外部实体与网络访问；确认 zip 解压有大小和数量上限。缺少的项目列入 SECURITY.md 的已知限制，或者在本批修复。

### 退出条件

一位未参与开发的人（或一个全新的 agent 会话）在三个平台上只按 README 操作，就能完成安装、`--plan`、`--doctor`，以及与平台对应的构建；所有新文档中的链接都能打开（CI 加 markdown 链接检查）；`pip install .` 后命令行入口可用。

## 14. P9：发布候选回归与发布

### 执行动作

1. 冻结 `release/1.0` 分支，依次运行：三平台 CI；macOS 真实 Word（NW01–NW12、`remediation_word`）；Windows 真实 Word（WW01–WW12）；LibreOffice 草稿；P7 的 agent 试用。
2. 在全新目录中，对三个示例在每个平台上执行 `--plan`、`--doctor`、构建和 `smoke_test.py`，然后清理 `.work`、Office 锁文件、渲染缓存和 Python 缓存（沿用 Skill 中的发布要求）。
3. 生成 `docs/acceptance/public-release/status.json`：每项声称的能力都对应测试 ID、支持级别、平台、环境和证据路径。
4. 发布前最后检查：`test_repo_hygiene`；`git status` 干净；`git diff --check`；在全新克隆中跑一次快速开始。
5. 打 `v1.0.0` 标签，发布 GitHub Release：附变更说明、平台矩阵、已知限制和压缩的证据包。
6. 发布后一周内跟踪 issue，把首批平台问题整理为 `1.0.x` 修复。

### 退出条件

第 1.3 节矩阵中的每一格都有对应证据；`status.json` 中没有 `blocked` 或 `partial` 被写成 `verified`；Release 页面的能力描述与 README 完全一致。

## 15. 风险与应对

| 风险 | 影响 | 应对 |
| --- | --- | --- |
| 没有可用的 Windows + Word 测试机 | P3/P4 无法验收 | 尽早确认设备或云桌面；在 P4 完成之前，README 中 Windows 只能写 `draft` 支持，不能提前写 `exact` |
| CI 中无法运行真实 Word | 回归只能手动执行 | 真实 Word 测试继续按环境变量开启；每次发布前手动执行并保存证据，CI 只运行离线测试和假后端测试 |
| macOS 与 Windows Word 分页不一致 | 同一夹具在不同平台得到不同页码 | 以平台各自的 `exact` 结果为准；文档说明“页码只对生成它的 Word 环境有效”，P4 记录全部差异 |
| COM 卡在模态对话框 | 构建挂起 | 子进程加超时，结束时按 PID 清理，并把常见对话框原因写入 doctor 提示 |
| 字体缺失引起分页变化 | 用户误以为结果错误 | P6 在 doctor 和 plan 中提前报告，回退映射必须显式启用 |
| 证据瘦身被误认为删除历史 | 审计链不完整 | 保留哈希与下载说明；不改写 git 历史 |
| 各 agent 的 Skill 发现路径变化 | Skill 不被加载 | 只维护一份源文件，用同步脚本分发；每次发布前按官方文档核对路径，并在 P7 中做实际试用 |

## 16. 待确认事项

以下事项会影响范围，建议在 P0 开始前确认：

1. 是否有可用的 Windows + Microsoft Word 测试环境，以及 Word 版本。
2. 最低 Python 版本选 3.9 还是 3.10。
3. 是否接受可选的 LibreOffice 草稿后端（会增加维护量），还是只做 `none` 加 TOC 域的草稿模式。
4. 证据瘦身采用 Release 附件还是独立分支。
5. 是否计划发布到 PyPI（影响 P8 的打包深度和包名选择）。
