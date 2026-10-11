# P3 实施记录：Windows COM 后端

> **文档状态**：已实施，待验收（Implemented, pending acceptance）  
> **基线依赖**：P2（Office 后端抽象，`main` = `335a821`）  
> **核心目标**：按[执行计划](windows-and-public-release-plan.md)第 3.3 节与第 8 节，在 Windows 上用 Microsoft Word / PowerPoint COM 实现 `OfficeBackend` 协议，使 Windows + Word 能以 `exact` 完成构建；四种故障（强制结束 Word、目标文件被占用、书签缺失、超时）都给出可读报错，且不残留 `WINWORD.EXE` 和半成品交付物。  
> **验收分工**：Windows 侧由本批次在本机完成；macOS 真实 Word 回归（45 项基线）由 macOS 环境复核后，才能把状态改为“已完成”。

---

## 1. 模块与职责

```
lib/office/
├── win_com.py          # WinComBackend（父进程）：静态检查、调度 worker、校验产物、原子替换、进程回收
└── win_com_worker.py   # worker 子进程：CoInitialize、DispatchEx、COM 调用、HRESULT 翻译
```

| 组件 | 职责 |
| --- | --- |
| `WinComBackend.static_status` | 只查注册表 `HKCR\Word.Application\CLSID` 与 pywin32 能否导入，不启动 Word；附带 `powerpoint_registered` |
| `WinComBackend.probe` | 启动隔离实例读取 `Version`、`Build`、`Path` 后退出，供 `--doctor` 使用；失败时返回不可用状态，不抛异常 |
| `convert_doc_to_docx` / `export_pdf` / `inspect_bookmarks` | 把输入复制到 `%TEMP%\document-synthesis-word-access\job-<随机>`，移除 `Zone.Identifier`，交给 worker；产物校验通过后才替换目标 |
| `convert_pptx_to_pdf` | PowerPoint 未注册时只让本操作报 `unsupported`；用户已打开 PowerPoint 时不 `Quit` |
| `run_worker` | `python -m lib.office.win_com_worker`，stdin 一行请求、stdout 一行结果；超时结束 worker，并按 pid 文件结束本次启动的 Office |
| `cleanup_recorded_office` | 只结束映像名与创建时间都和记录一致的进程；启动阶段卡住时，只结束“快照之后新出现、由 COM 启动（命令行带 `/Automation -Embedding`）、唯一”的进程 |
| `com_failure_message` | 常见 HRESULT 与错误描述翻译为中文提示 |

COM 调用、数值常量和导出参数见 [Office 后端对照](office-backends.md)。

## 2. 关键决策

1. **每次操作一个 worker、一个 Word 实例**。COM 同步调用无法从进程内中断，只有子进程能被可靠地超时结束。代价是每次操作多约 5 秒的 Word 启动时间；最小示例的完整构建约 30 秒。
2. **PID 取得方式**：给隐藏实例设置唯一的 `Application.Caption`，用 `FindWindow("OpusApp", caption)` 取窗口，再取进程 ID。记录 PID 的同时记录创建时间，防止 PID 复用后误杀。
3. **启动阶段的孤儿进程**：Word 卡在首次运行或许可证对话框时，`DispatchEx` 不返回，worker 来不及记录 PID。worker 在启动前写下 `WINWORD.EXE` 快照；超时后，父进程只在候选唯一时结束它。有多个候选（例如另一个构建同时在启动 Word）时宁可不结束。
4. **跨盘符替换**：工作目录默认在 `C:` 的临时目录，输出可能在其他盘符，`os.replace` 无法跨盘。产物先复制到目标所在目录的临时文件，再在同一目录内原子替换，失败时保留原目标。首次在 `D:` 盘仓库构建时发现了这个问题（`WinError 17`），修复后补充了说明。
5. **不使用 `gencache`**：全部用数值常量与动态分派，避免 `gen_py` 缓存损坏。
6. **NullBackend 提示**：Windows 上的 `auto` 现在选择 `WinComBackend`，“未安装 Word”由它报告；`NullBackend` 在 Windows 上只表示后端被显式设为 `none`（或 COM 后端模块无法导入）。

## 3. 测试

| 文件 | 类型 | 内容 |
| --- | --- | --- |
| `tests/test_win_com_backend.py` | 离线，三平台 CI 均运行 | HRESULT 翻译；worker 用 COM 假对象验证调用序列、常量、finally 中关闭与 `Quit`；PowerPoint 单实例；父进程用假 runner 验证产物校验、失败不替换目标、工作目录清理；用替身子进程验证成功、错误分类、异常退出、超时；pid 记录核对与孤儿判定；Windows 上额外验证超时只结束记录的进程、不动同名的未记录进程 |
| `tests/test_win_word_faults.py` | 真实 Word，需 Windows + `DOCUMENT_SYNTHESIS_WORD_TEST=1` | 分节重新编号文档的物理页/打印页；书签缺失；目标 PDF 被占用；启动阶段与处理阶段超时；操作中强制结束 Word；构建级超时不发布交付物。每项都检查没有残留 `WINWORD.EXE` 和工作目录 |
| `tests/test_office_backends.py` | 离线 | 隔离断言扩展：`win32com`、`pythoncom`、`DispatchEx`、`WINWORD` 只能出现在两个 Windows 后端文件中，且后端中不得出现 `taskkill` 或 `/IM` |

## 4. 本机验证结果

环境：Windows 11 Pro（10.0.26300）；Microsoft 365 Word 16.0（build 16.0.20430，64 位，Click-to-Run），PowerPoint 已安装；Python 3.14.4；已加载的 Word COM 插件包括 iSlide Tools 与 Acrobat PDFMaker。日期：2026-10-11。

| 项目 | 命令 / 范围 | 结果 |
| --- | --- | --- |
| `--doctor` | 三个示例各一次 | 退出码 `0`；报告 Word 版本、位数、Click-to-Run、PowerPoint 可用 |
| 最小示例 | `examples/minimal-demo` 构建 + `smoke_test.py` | 通过（`exact`，4 页） |
| 中文示例 | `examples/chinese-demo` 构建 + `smoke_test.py` | 通过（`exact`） |
| 自定义格式示例 | `examples/custom-format-demo` README 全流程（分析、编译、预览、迁移、映射、doctor、构建、`smoke_test.py`） | 全部退出码 `0` |
| 全量测试（含真实 Word） | `DOCUMENT_SYNTHESIS_WORD_TEST=1 python -m unittest discover -s tests` | 410 项通过，1 项跳过（`test_N6_C06_migration_rejects_symlink_alias`：未开启开发者模式，无权创建符号链接）；结束后无 `WINWORD.EXE` 残留 |
| 故障注入 | `tests/test_win_word_faults.py` | 6 项通过，报错如下 |

| 故障 | 实际报错 | 目标文件 | 残留 |
| --- | --- | --- | --- |
| 书签缺失 | `Word 文档中缺少书签: absent_mark` | 未生成 | 无 |
| 目标 PDF 被占用 | `无法更新分页诊断文件（目标文件可能正在 Word 或 PDF 阅读器中打开）: …` | 保留原内容 | 无 |
| 超时（8 秒） | `等待 Microsoft Word 响应超时（8 秒），已结束本次启动的 Office 进程。请检查文档是否会触发 Microsoft Word 的对话框（密码、修复、宏或许可证提示）。` | 未生成 | 无 |
| 强制结束 Word | `Microsoft Word 进程在自动化过程中意外退出（可能被强制结束或崩溃），本次操作未完成，已停止。（The remote procedure call failed.）` | 未生成 | 无 |
| 构建级超时（3 秒） | 构建停止，报超时 | 未发布 `demo_complete.docx` | 无 |

实施中发现并修复的问题：

1. **跨盘符替换失败**：仓库在 `D:`、工作目录在 `C:`，首轮三个示例构建都报 `WinError 17`（当时没有发布任何交付物）。修复见第 2 节第 4 条。
2. **插件窗口使孤儿判定失效**：首轮全量测试中途挂起一次，事后发现一个由 COM 启动的 `WINWORD.EXE` 显示着 iSlide 插件窗口，最初“无可见窗口”的孤儿判定因此放过了它（该轮被手动停止，父进程的清理没有执行）。改为按 `/Automation -Embedding` 命令行识别。那次挂起的确切调用栈没有取到；之后两轮全量测试与故障注入都没有复现。插件在自动化实例中照常加载，本后端不修改插件设置（`COMAddIn.Connect` 会改写用户的 LoadBehavior）。
3. **现有测试的 Windows 假设**：`test_post_review_word.py`（NW06、NW09、NW10）与 `test_remediation_word.py`（SW03）打开交付 DOCX 的压缩包后没有关闭，Windows 上临时目录无法删除，已改为 `with`；`test_pagination.py` 的导出失败用例只替换了 `subprocess.run`，没有固定后端，在装有 Word 的 Windows 上会选中 COM 后端并真实导出成功，已改为与相邻用例一样固定 macOS 后端。其余真实 Word 测试没有发现 macOS 专属假设（`osascript`、`/private/tmp`、`darwin`）。

## 5. 未完成与后续

- **macOS 回归**：本批次修改了 `NullBackend` 的 Windows 提示、`--doctor` 的附加信息与依赖声明，未改动 `mac_applescript.py`；仍需在 macOS 上跑真实 Word 回归（45 项基线）确认行为不变。
- **P4**：WW01–WW12 验收矩阵、跨平台页码对照、较旧 Office 版本、`environment.json` 证据，均不在本批次范围内。在 P4 完成前，README 中 Windows 只写“预览支持”。
- **未实现的计划项**：超过 259 字符的输入路径目前依赖工作目录较短（输入总是先复制到短路径的工作目录），没有单独测试；字体检查属于 P6。
