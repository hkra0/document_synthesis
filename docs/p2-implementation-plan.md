# P2 实施计划：Office 后端抽象（macOS 行为不变）

> **文档状态**：已完成 (Completed)  
> **基线依赖**：P0（脱敏与代码基线）、P1（跨平台加固与 CI）  
> **核心目标**：建立 `lib/office/` 抽象协议与多后端架构，将 AppleScript 从业务模块中剥离并原样迁移到 `mac_applescript.py`；生产代码中对 `osascript` 与 `sys.platform == "darwin"` 实施严格隔离；统一 `--doctor` 退出码约定；确保现有 macOS 行为与页码记录 100% 不变。

---

## 1. 架构设计与模块划分

### 1.1 模块结构 (`lib/office/`)

```
lib/office/
├── __init__.py          # get_backend(), set_backend(), 异常与数据结构重导出
├── base.py              # OfficeBackend Protocol, BackendStatus, BookmarkReport, OfficeBackendError
├── mac_applescript.py   # macOS AppleScript 原样实现 (MacAppleScriptBackend)
├── null.py              # 无 Office / 不支持平台后端 (NullBackend)
├── win_com.py           # Windows COM 后端 (P3 占位 / 待实施)
└── libreoffice.py       # LibreOffice 草稿后端 (P5 占位 / 待实施)
```

### 1.2 核心接口定义 (`lib/office/base.py`)

- **`OfficeBackend(Protocol)`**：
  - `name: str`（如 `"mac-word"`, `"win-word"`, `"libreoffice"`, `"none"`）
  - `fidelity: str`（`"exact"` | `"draft"` | `"unsupported"`）
  - `static_status() -> BackendStatus`：静态可用性检查，不触发授权、不启动应用。
  - `probe() -> BackendStatus`：`--doctor` 专用只读探针，验证真实自动化控制能力。
  - `convert_doc_to_docx(src: Path, dst: Path) -> None`：旧版 `.doc` 转换为现代 `.docx`。
  - `convert_pptx_to_pdf(src: Path, dst: Path) -> None`：`.pptx` 转换为 PDF。
  - `export_pdf(docx: Path, pdf: Path) -> None`：DOCX 导出 PDF。
  - `inspect_bookmarks(docx: Path, pdf: Path, names: Sequence[str]) -> BookmarkReport`：排版重新分页并获取各书签物理页与打印页码，输出导出的 PDF。
- **`BackendStatus`**：包含 `available: bool`, `reason: str`, `fidelity: str`, `details: Dict[str, Any]`。
- **`BookmarkReport`**：包含 `bookmarks: Dict[str, Dict[str, Any]]`, `total_pages: int`, `pdf_path: Path`, `fidelity: str`。
- **`OfficeBackendError`** / **`OfficeUnsupportedError`**：统一定义异常体系。

### 1.3 后端选择与优先级 (`lib/office/__init__.py`)

- 优先级：显式参数 `--office-backend` > 环境变量 `DOCUMENT_SYNTHESIS_OFFICE_BACKEND` > `auto` 自动探测。`--office-backend` 只接受 `auto`/`word`/`none`；未传参数时不覆盖环境变量；`libreoffice` 在 P5 之前明确报错，不静默退回其他后端。
- `auto` 策略：
  - macOS (`sys.platform == "darwin"`) → `MacAppleScriptBackend`
  - Windows (`sys.platform == "win32"`) → `WinComBackend` (P3) / 当前为 `NullBackend`
  - 其他平台 (Linux) → `NullBackend`
  - *注：LibreOffice 永不被 auto 选中，必须显式指定以防草稿混淆*。
- 支持 `set_backend()` 依赖注入，便于测试套件无缝挂载假后端。

---

## 2. 现有业务模块迁移映射

| 业务模块 | 现有原样代码 | 迁移后方案 | 隔离原则 |
| --- | --- | --- | --- |
| `lib/qa.py` | `_word_access_directory` | 移入 `MacAppleScriptBackend`，非 mac 后端使用 `tempfile`；`lib/qa._word_access_directory()` 作为委托薄包装保留 | `lib/qa.py` 内部不再出现 `/private/tmp` 与 `darwin` |
| `lib/qa.py` | `word_export_status()` | 委托 `get_backend().static_status()` | 保留向后兼容签名 `(bool, str)` |
| `lib/qa.py` | `word_automation_status()` | 委托 `get_backend().probe()` | 保留向后兼容签名 `(bool, str)` |
| `lib/qa.py` | `export_docx_to_pdf()` | 委托 `get_backend().export_pdf()` | 捕获 `OfficeBackendError` 返回布尔值 |
| `lib/qa.py` | `get_exact_printed_heading_pages()` | 无生产代码调用；增加 `DeprecationWarning` 弃用告警 | 保留实现以兼容历史外部调用与测试 |
| `lib/pagination.py` | `_inspection_script`, `inspect_document` | `_inspection_script` 移入 `mac_applescript.py`；`inspect_document` 委托 `backend.inspect_bookmarks()` 并保留页码标签后处理 | 生产代码中彻底移除 `osascript` 字符串 |
| `lib/engine.py` | `prepare_conversions`, `_run_applescript` | `prepare_conversions` 委托 `backend.convert_doc_to_docx()` 与 `backend.convert_pptx_to_pdf()`；`_run_applescript` 移入 `mac_applescript.py`，原位置留薄包装 | 移除 `lib/engine.py` 中的 `osascript` 与 `darwin` |

---

## 3. `--doctor` 退出码规范

在 `synthesize.py:run_doctor` 与 CLI 中规范三类退出码：
- **`0`**：环境完全就绪，支持 `exact` 精确构建（Python 依赖齐全、输入合法、Office 探针通过且 fidelity 为 exact）。
- **`2`**：Python 依赖就绪，但无可用精确 Office 自动化（例如 Linux/无 Word 的 macOS/Windows 环境）；系统支持分析类命令与草稿模式。
- **`1`**：缺少核心 Python 依赖（`docx`, `pymupdf` 等）或命令行输入参数无效。

更新 `README.md` 与 `.github/workflows/ci.yml` 验证该约定。

---

## 4. 测试与验收策略

### 4.1 新增假后端契约测试 (`tests/test_office_backends.py`)
1. **成功构建全路径**：用假后端驱动完整文档合成流程，产出合规文档并保留元数据。
2. **后端超时**：模拟后端超时或阻塞，断言捕获清晰的错误提示，且临时工作目录不泄漏、交付物不被污染。
3. **书签缺失**：后端返回的书签报告缺失必需书签，门禁断言拦截并报错。
4. **导出 PDF 缺失**：后端未能生成导出 PDF，门禁断言拦截并报错。
5. **草稿一律拒绝**：后端 fidelity 不是 `"exact"` 时一律拒绝构建，且没有绕过开关。2026-10-10 复验发现先前的 `--allow-draft` 会发布不带草稿标注的交付物，已移除；该开关连同文件名、诊断与终端标注一起在 P5 实施。
6. **架构隔离断言**：扫描 `lib/` 目录下全部生产 Python 代码，断言：
   - `"osascript"` 字符串字面量仅出现在 `lib/office/mac_applescript.py`。
   - `sys.platform == "darwin"` / `'darwin'` 仅出现在 `lib/office/mac_applescript.py` 与 `lib/office/__init__.py`。

### 4.2 回归验证
1. 离线测试集：执行完整无 Office 测试，保证通过数量完全对齐基线。
2. 真实 Word 回归：在开启 `DOCUMENT_SYNTHESIS_WORD_TEST=1` 时运行 NW01–NW12 和 `test_remediation_word`，确保书签页码记录与基线 100% 逐项一致。
3. 代码与仓库卫生：执行 `tests/test_repo_hygiene.py`，确保无敏感路径残留。

---

## 5. 详细执行清单

- [x] 1. 创建 `lib/office/base.py`（Protocol、数据模型与异常）
- [x] 2. 创建 `lib/office/mac_applescript.py`（原样迁移 AppleScript 与平台支持逻辑）
- [x] 3. 创建 `lib/office/null.py`（多平台友好诊断信息）
- [x] 4. 创建 `lib/office/__init__.py`（后端选择器与注册机制）
- [x] 5. 改造 `lib/qa.py`、`lib/pagination.py`、`lib/engine.py` 为薄包装
- [x] 6. 标记 `qa.get_exact_printed_heading_pages` 弃用
- [x] 7. 实现 `--doctor` 退出码约定 (0 / 2 / 1)，更新 `synthesize.py`、`README.md` 与 CI 配置
- [x] 8. 编写 `tests/test_office_backends.py`（5 大故障与成功场景 + 平台隔离断言）
- [x] 9. 执行离线全套测试与真实 Word 测试，比对基线

---

## 6. 验收与回归结果总结

### 6.1 测试集执行记录

1. **离线全量单元测试**：
   - 运行命令：`python3 -m unittest discover -s tests -v`
   - 结果：**345 项测试全部通过（6 项跳过），0 失败，0 错误**（基线 338 项 + 新增 7 项契约测试）。
2. **真实 macOS Word 回归测试**：
   - 运行命令：`DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest tests/test_remediation_word.py tests/test_word_formatting.py tests/test_post_review_word.py -v`
   - 结果：**27 项真实 Word 测试全部通过**（覆盖 NW01–NW12、SW01–SW07、W04–W07），页码测量与版式质检 100% 对齐。
3. **架构与平台隔离断言**：
   - 运行命令：`python3 -m unittest tests.test_office_backends.OfficeBackendContractTests.test_production_code_isolation_assertions -v`
   - 结果：生产代码（`lib/`）扫描断言完全通过：
     - `"osascript"` 仅存在于 `lib/office/mac_applescript.py`。
     - `sys.platform == "darwin"` / `'darwin'` 仅存在于 `lib/office/mac_applescript.py` 和 `lib/office/__init__.py:get_backend()`。
4. **仓库卫生与脱敏检查**：
   - 运行命令：`python3 -m unittest tests/test_repo_hygiene.py -v`
   - 结果：4 项检查全部通过，无用户主目录或敏感邮箱残留。
5. **`--doctor` 退出码规范**：
   - `0`：`exact` 精确构建就绪。
   - `2`：仅支持 `draft` 或分析类功能。
   - `1`：依赖缺失或参数错误。
   - CI 配置 `.github/workflows/ci.yml` 与 `README.md` 已同步更新。

### 6.2 复验修正（2026-10-10）

独立复验确认：离线 345 项通过（6 项跳过）；macOS 真实 Word 回归（`test_post_review_word`、`test_remediation_word`、`test_word_formatting`、`test_m1_acceptance`、`test_thesis_acceptance`、`test_pagination`、`test_publication_gates`）共 45 项通过、0 跳过。同时修正了以下问题：

1. 草稿门禁：移除 `--allow-draft` 与 `DOCUMENT_SYNTHESIS_ALLOW_DRAFT`，非 `exact` 结果一律拒绝（`test_draft_fidelity_always_rejected`）。
2. 后端选择：`--office-backend` 默认值改为未设置，环境变量不再被覆盖；无效值返回退出码 1（`test_cli_backend_option_precedence`）。
3. 未实现后端：`libreoffice` 明确报“尚未支持”；`NullBackend` 的提示不再指向不存在的命令（`test_unimplemented_backends_are_reported_not_substituted`）。

说明：P0 基线只记录了真实 Word 的通过数量，没有逐书签页码，因此“页码与基线逐项一致”无法逐项比对，只能以全部真实 Word 用例通过作为等价证据。
