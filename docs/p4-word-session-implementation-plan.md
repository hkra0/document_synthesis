# P4 后续实施记录：一次构建复用一个 Word 实例

> **文档状态**：已完成 (Completed)；macOS 侧仅涉及构建元数据的取值来源，待 macOS 会话回归确认  
> **基线依赖**：P4（Windows 真实 Word 验收矩阵，已合并，`main` = `466b2a7`；本分支基于其中的 `3405aef`）  
> **核心目标**：P4 实测中一次构建要多次启动 Word（每次 4–6 秒）。本批次让一次构建内的全部 Word 操作共用同一个 worker 子进程与同一个 Word 实例，构建元数据不再为记录可用性额外启动 Word；P3 建立的可靠性约定（逐请求超时、只结束本次启动的 Word、孤儿判定、Quit 看门狗、诊断日志、不影响用户文档）全部保持。macOS 后端（`lib/office/mac_applescript.py`）不改动。

---

## 1. 问题

P3 的设计是“每次操作一个 worker、一个 Word 实例”。一次 `exact` 构建中启动 Word 的位置：

| 位置 | 调用 | 次数 |
| --- | --- | --- |
| `lib/engine.py` `prepare_conversions` | `convert_doc_to_docx`（目录型输入中的 `.doc`） | 每个 `.doc` 一次 |
| `lib/delivery.py` `build_deliveries` | `inspect_document` → `inspect_bookmarks`，目录回填最多 3 趟；模板引入新的 PAGEREF 目标时同一趟再测一次 | 每份交付物 1–6 次 |
| `lib/engine.py` `generate_build_metadata` | `word_automation_status()` → `backend.probe()`，只为在元数据中记录 Word 是否可用 | v3 构建每次 1 次 |
| `smoke_test.py` | `validate_delivery` → `inspect_document` | 每份交付物 1 次 |

P4 实测（`ww-results.json`，Windows 11 + M365 Word 16.0.20430）：每次启动 Word 约 4–6 秒；一次构建 9–20 秒，其中大部分是重复启动 Word。用优化前代码（`aa1f57b`）对 WW01–WW12 的夹具逐个构建并每 0.1 秒采样进程表，构建期间分别出现 2–3 个 `WINWORD.EXE`：1–2 次分页核验，加上元数据探针 1 次。

## 2. 设计

### 2.1 组件

```
lib/office/
├── __init__.py         # office_session()：块内 get_backend() 返回同一个后端实例
├── win_com.py          # WordSession（父进程会话）、WinComBackend.session()/session_status()
└── win_com_worker.py   # serve_session()（会话主循环）、launch_word/shutdown_word、run_word_op
```

| 组件 | 职责 |
| --- | --- |
| `lib.office.office_session()` | 后端有 `session()` 时，在块内把它设为活动后端（块内各处 `get_backend()` 拿到同一个实例），并进入其会话；没有 `session()` 的后端（macOS、空后端）原样执行。`UnifiedSynthesizer.synthesize` 与 `smoke_test.test_project` 各包一层 |
| `WinComBackend.session()` | 上下文管理器：块内 Word 操作（`probe`、`convert_doc`、`export_pdf`、`inspect`）交给 `WordSession`；`convert_pptx` 启动的是 PowerPoint，仍走单次模式；嵌套调用复用外层会话；`DOCUMENT_SYNTHESIS_WORD_SESSION=0` 时不复用 |
| `WordSession`（父进程） | 第一个请求时才启动 worker；每个请求写一行带 `id` 的 JSON，后台线程读取 stdout/stderr，按请求计时；超时、崩溃、实例失效时回收，下一个请求启动新 worker |
| `serve_session`（worker） | 第一个 Word 请求时启动 Word；之后每个请求在同一实例上执行并输出一行结果；首次启动时把 Word 版本随结果带回；`close` 或 stdin 结束时退出 Word |
| `WinComBackend.session_status()` / `lib.qa.word_build_status()` | 构建元数据用：返回本后端实例实际调用过的 Word（版本、位数等）；从未调用过时用 `static_status()`。都不启动 Word |

### 2.2 协议

单次模式不变：stdin 一行请求，stdout 一行结果。会话模式：

```
→ {"op": "session", "pid_file": "<工作目录>/job-session-<随机>/office.pid"}
→ {"op": "inspect", "id": 1, "timeout": 600, "docx": ..., "pdf": ..., "names": [...]}
← {"ok": true, "page_map": "...", "word": {"word_version": "16.0", ...}, "id": 1}
→ {"op": "inspect", "id": 2, ...}
← {"ok": false, "kind": "error", "error": "Word 文档中缺少书签: x", "id": 2}
→ {"op": "close"}
```

worker 在每个请求开始时向 stderr 写一行 `request: <op> #<id>`，之后照旧在每次可能阻塞的 COM 调用前写 `step:`；超时前 3 秒用 `faulthandler` 导出调用栈，请求结束后取消。

### 2.3 可靠性约定如何保持

| 约定 | 单次模式（P3） | 会话模式 |
| --- | --- | --- |
| 每个请求有超时 | worker 整体超时 | 父进程逐请求计时（`DOCUMENT_SYNTHESIS_OFFICE_TIMEOUT` 同样生效）；第一个请求的超时包含 Word 启动 |
| 超时后只结束本次启动的 Word | 结束 worker，按 pid 文件（PID + 映像名 + 创建时间）结束 Word，绝不按映像名 | 相同：结束 worker，立即按会话的 pid 文件结束 Word；会话随后在下一个请求时启动新的 worker 与 Word |
| 启动阶段的孤儿 | pid 文件只有启动前快照时，30 秒内寻找“快照之后新出现、命令行带 `/Automation -Embedding`、唯一”的进程 | 相同（复用 `cleanup_recorded_office`）。Word 只在会话的第一个请求（或重启后的第一个请求）中启动 |
| Quit 看门狗 | `Quit` 后 30 秒仍未退出，输出已完成的结果并结束 worker | 会话结束时同样武装；所有结果此前都已输出，看门狗只需结束 worker，父进程按 pid 文件结束 Word |
| 诊断日志 | 超时报错给出最后一次 COM 调用，stderr 另存为 `diagnostics/timeout-*.log` | 相同；“最后一次 COM 调用”只在最后一个 `request:` 分隔行之后查找，不会把上一个请求的步骤当成挂起点；日志保留整个会话的 stderr |
| 失败不留半成品 | 产物在作业目录生成，校验通过后才替换目标 | 不变：每个请求仍有自己的作业目录与私有副本 |
| 不影响用户文档 | 独立实例（`DispatchEx`），只结束记录的 PID | 相同，另见 2.4 |

请求失败的分类：

- **校验失败**（`WorkerFailure`，如书签缺失）：只影响当前请求，文档已在 `finally` 中关闭，实例保留，会话继续。
- **其他异常**（COM 错误、RPC 断连、对话框阻塞重试超限、启动失败）：实例状态不可信。结果带 `"restart": true`，worker 退出 Word 后结束；父进程按单次模式“正常退出”的规则回收（给 Word 15 秒自行退出，否则按 pid 文件结束），下一个请求启动新 worker。
- **两次请求之间 Word 已退出**（被用户或其他程序结束）：复用前先调用 `Application.Version`。失败且 pid 文件记录的进程已不存在时，直接启动新实例；进程仍在却无法调用时按上一条处理，不另起实例，以免留下一个挂住的 Word。

### 2.4 构建期间用户打开的文档

实测（Windows 11 + M365 Word 16.0.20430）：隐藏的自动化 Word 运行期间，用户在资源管理器中双击 `.docx`，文档会被系统路由进这个隐藏实例；随后的 `Quit(wdDoNotSaveChanges)` 会不经提示丢弃用户在其中的修改。P3 的单次模式同样存在这个窗口（每次操作数秒），会话把它延长到整个构建，所以本批次一并处理：

- worker 记录自己交给 Word 打开过的文件（`Documents.Open` 的源文件与 `SaveAs2` 的目标）。
- 退出 Word 前遍历 `Documents`：自己遗留的文档先关闭；仍有其他文档时不 `Quit`，改为恢复 `ScreenUpdating`、`DisplayAlerts`、`AutomationSecurity`（按界面设置）并设 `Visible = True`，再把 pid 记录改写为 `{"released_pid": ...}`。父进程读不到有效记录，不会结束这个 Word。
- 单次模式与会话模式共用这段逻辑。
- 局限：如果请求超时（worker 已无响应），父进程仍会按 pid 文件结束 Word，其中路由进来的用户文档无法保住。超时本身意味着 Word 已卡住，此时没有更安全的选择。

### 2.5 构建元数据

`generate_build_metadata` 改为调用 `lib.qa.word_build_status()`：

- Windows 会话中：首个启动 Word 的请求带回版本信息，`session_status()` 返回 `available = true`，`reason` 为 `Automation 会话成功: Microsoft Word 16.0 (build …), 64 位。`（格式与 `--doctor` 探针的 `Automation 探针成功: …` 相同）。
- 没有实际调用过 Word 的后端（macOS、`DOCUMENT_SYNTHESIS_WORD_SESSION=0`）：`static_status()`，如 `Microsoft Word 精确分页可用。`。

字段 `environment.word_available`、`environment.word_automation.available`、`environment.word_automation.reason` 的类型与含义不变（布尔值 + 说明文字）。v3 构建只有在 `exact` 分页核验全部通过后才生成元数据，此时 Word 实际可用已被证明，静态检查不会给出与事实不符的结论。`word_automation_status()` 保留，`--doctor` 与测试仍用它做真实探针。

## 3. 设计取舍

1. **长连接 worker，而不是在父进程中持有 COM 对象**。COM 调用无法从进程内中断，只有子进程能被可靠地超时结束；会话只是让同一个子进程处理多个请求，超时语义不变。代价是父进程需要后台线程读取 stdout/stderr 并逐请求计时（单次模式用 `communicate(timeout=…)` 即可）。
2. **会话边界是一次构建 / 一次 smoke 项目检查**，由 `office_session()` 显式进入，而不是进程级常驻 Word。常驻实例会把“用户文档被路由进隐藏实例”的窗口扩大到整个进程生命周期，也需要额外的空闲回收；构建结束即退出，行为与 P3 对用户的可见效果一致（构建结束后没有残留 Word）。`--doctor` 只有一次探针，不进入会话。
3. **懒启动**。会话进入时不启动 Word；`--structure-only`、未调用 Word 的路径不会多出一次启动。
4. **失败后整实例重启，而不是在原实例上继续**。除书签缺失这类在 worker 内判定的校验失败外，任何 COM 异常都让会话换新实例。这样实例状态（打开的文档、模态对话框、半完成的导出）不会带到下一个请求；失败本身已经让构建停止，重启的成本只在少数重试场景出现。
5. **复用前的健康检查只做一次廉价调用**（`Application.Version`）。它能发现两次请求之间被结束的 Word；仍在运行却不响应的实例不另起新实例，交给父进程按 pid 文件回收，避免一个挂住的 Word 与一个新 Word 同时存在。
6. **PowerPoint 不进会话**。`.pptx` 转换只在目录型输入中出现、每个文件一次，且 PowerPoint 是单实例应用（用户已打开时不能 `Quit`）；放进会话会让一个 pid 文件同时记录两种进程，得不偿失。
7. **保留单次模式与开关**。`DOCUMENT_SYNTHESIS_WORD_SESSION=0` 回到每次操作一个实例，用于排障与本次的前后对比；离线测试注入 `runner` 而未注入会话工厂时，会话自动不生效，旧测试不受影响。
8. **元数据不再探针，而不是把探针结果缓存到磁盘**。缓存会让元数据记录的是“某次过去的探针”；复用本次构建实际调用 Word 的结果更准确，也不需要失效策略。

## 4. 测试

| 文件 | 类型 | 新增内容 |
| --- | --- | --- |
| `tests/test_win_com_backend.py` `SessionWorkerLoopTests` | 离线，三平台 | `serve_session` 用内存请求行与假 COM 驱动：4 个不同操作只 `DispatchEx` 一次、只 `Quit` 一次，每个文档都关闭，首个结果带 Word 版本；逐请求武装/取消调用栈导出；没有请求时不启动 Word；书签缺失与会话不支持的操作保留实例；COM 错误与启动失败带 `restart` 并结束会话、不再读取后续请求；两次请求之间 Word 已退出时换新实例、不再调用旧实例；Word 仍在但不响应时结束会话；用户文档移交（不 `Quit`、可见、pid 记录撤销、父进程不结束）；遗留的自身文档先关闭再 `Quit`；单次模式同样移交 |
| 同上 `WordSessionProcessTests` | 离线，三平台（最后一项仅 Windows） | 父进程 `WordSession` + 替身会话 worker 子进程（`tests/fake_word_session_worker.py`：真实 `serve_session`、假 COM，按环境变量注入故障）：多请求只启动一个 worker，`close` 后退出码 0、工作目录删除；请求间书签缺失不重启；COM 错误、worker 崩溃（退出码 3）后下一个请求启动新 worker；第二个请求超时只报告该请求的挂起步骤（`Document.Repaginate`）、保存含 `request: inspect #2` 与调用栈的诊断日志，下一个请求恢复；会话结束时 `Quit` 卡住由看门狗在 1 秒（测试值）后结束；Windows 上会话超时只结束 pid 文件记录的替身进程，不动另一个同名进程 |
| 同上 `BackendSessionTests` | 离线，三平台 | 会话内 `inspect`/`export_pdf`/`probe` 进会话、`convert_pptx` 走单次；每个请求超时仍为 600/120 秒；失败时不替换已有 PDF、不留作业目录；会话块抛异常时也关闭会话；嵌套复用外层；`DOCUMENT_SYNTHESIS_WORD_SESSION=0` 关闭；注入 runner 时不启动真实会话；`session_status()` 不发请求；`office_session()` 块内 `get_backend()` 返回同一实例、块外恢复；macOS 与空后端原样通过 |
| 同上 `BuildMetadataWordStatusTests` | 离线，三平台 | `word_build_status()` 与 `generate_build_metadata` 不调用 `probe()`；字段与类型不变 |
| `tests/test_win_word_faults.py` | 真实 Word，`DOCUMENT_SYNTHESIS_WORD_TEST=1` | `test_session_reuses_one_word_for_several_requests`：会话中 3 次分页 + 1 次导出只有一个本次启动的 `WINWORD.EXE`；`test_session_recovers_after_timeout_and_killed_word`：会话中处理阶段 8 秒超时只结束会话的 Word，下一个请求换新实例成功；两次请求之间外部结束 Word，下一个请求成功；`test_document_opened_during_session_is_handed_over`：会话期间经 Shell 打开的文档确实被路由进会话的 Word，会话结束后该 Word 仍在、可见，文档仍打开且未保存的修改仍在 |

`test_build_timeout_publishes_nothing` 原用 3 秒超时，依赖“每次分页都重新启动 Word（4–6 秒）”必然超时；会话中后续分页请求不再包含启动时间，最小示例在 3 秒内即可完成每个请求。改为与启动阶段超时用例相同的 1 秒，并加 15 秒观察确认没有迟到的残留。`UserWordIsolationTests` 的“正常构建”一步现在经过会话，未改动即通过。

## 5. 前后耗时对比

方法：同一台机器、同一 Word 配置下，先在优化前代码（`aa1f57b`，独立 worktree）上运行 `word_matrix.py run --label windows-word --output ww-results-baseline-rerun.json`，紧接着在本分支运行 `--output ww-results-session.json`。两份结果都是 12 项 `verified`。表中同时列出 P4 原始的 `ww-results.json`（`561c9de`，同一环境）作为参照。单位：秒。

| 用例 | build 优化前（P4 原始） | build 优化前（同日复测） | build 优化后 | 用例总耗时 优化前（同日） | 用例总耗时 优化后 |
| --- | --- | --- | --- | --- | --- |
| WW01 | 13.8 | 15.2 | 6.3 | 30.6 | 21.4 |
| WW02 | 14.0 | 14.9 | 8.4 | 30.7 | 23.9 |
| WW03 | 11.2 | 11.2 | 7.8 | 26.7 | 25.6 |
| WW04 | 9.2 | 9.9 | 5.3 | 24.2 | 20.8 |
| WW05 | 9.6 | 9.2 | 5.8 | 24.1 | 20.5 |
| WW06 | 17.5 | 25.1 | 10.0 | 43.0 | 27.6 |
| WW07 | 20.2 | 15.9 | 7.2 | 31.4 | 22.5 |
| WW08 | 9.6 | 10.8 | 5.8 | 26.0 | 20.9 |
| WW09 | 17.7 | 19.4 | 9.6 | 37.3 | 26.9 |
| WW10 | 17.7 | 24.0 | 12.9 | 41.9 | 30.5 |
| WW11 | 10.8 | 12.8 | 7.8 | 31.7 | 26.5 |
| WW12 | 9.5 | 9.4 | 6.0 | 25.2 | 21.7 |

| 步骤 | 优化前（P4 原始）均值 / 范围 | 优化前（同日复测）均值 / 范围 | 优化后 均值 / 范围 | 变化（对同日复测） |
| --- | --- | --- | --- | --- |
| plan | 0.6 / 0.6–0.6 | 0.7 / 0.7–0.8 | 0.7 / 0.7–0.7 | 不调用 Word，无变化 |
| doctor | 4.6 / 4.1–5.0 | 4.4 / 4.1–5.3 | 4.2 / 4.1–4.4 | 只有一次探针，无变化 |
| build | 13.4 / 9.2–20.2 | 14.8 / 9.2–25.1 | 7.7 / 5.3–12.9 | −48%（合计 177.8 → 92.9） |
| smoke | 5.7 / 5.1–6.9 | 6.0 / 5.2–7.2 | 6.1 / 5.2–7.0 | 每个用例只有一份交付物、一次分页核验，本就只启动一次 Word，无变化 |
| 用例合计 | 29.2 / 23.8–36.5 | 31.1 / 24.1–43.0 | 24.1 / 20.5–30.5 | −23%（合计 372.8 → 288.8） |

每次构建启动 Word 的次数（同一夹具分别用两份代码构建、采样进程表）：

| 用例 | WW01 | WW02 | WW03 | WW04 | WW05 | WW06 | WW07 | WW08 | WW09 | WW10 | WW11 | WW12 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 优化前 | 3 | 3 | 2 | 2 | 2 | 3 | 3 | 2 | 3 | 3 | 2 | 2 |
| 优化后 | 1 | 1 | 1 | 1 | 1 | 1 | 1 | 1 | 1 | 1 | 1 | 1 |

`examples/minimal-demo` 与 `examples/chinese-demo` 的构建同样从 2 次降到 1 次，耗时分别为 12.9 → 7.2 秒、11.8 → 7.6 秒。

节省的时间基本等于“少启动的 Word 次数 × 4–6 秒”：多趟目录回填的用例（WW01、WW02、WW06、WW07、WW09、WW10）节省最多。一次 Word 启动仍保留在每次构建中；`doctor` 与 `smoke` 是独立进程，各自仍要启动一次。

## 6. 页码一致性

用 `word_matrix.py` 的 `compare_results`（与跨平台对照同一逻辑：书签的 `physical_page`、`printed_page`、`observed_label`，以及交付物导出 PDF 的页数与页码标签）对照：

| 对照 | 可对照用例 | 完全一致 | 有差异 |
| --- | --- | --- | --- |
| `ww-results.json`（P4 原始）↔ `ww-results-session.json` | 12 | 12 | 0 |
| `ww-results-baseline-rerun.json`（同日复测）↔ `ww-results-session.json` | 12 | 12 | 0 |

复用同一个 Word 实例没有改变任何页码。

## 7. 已知事项与后续

- **全量测试（含真实 Word）**：`DOCUMENT_SYNTHESIS_WORD_TEST=1 python -m unittest discover -s tests`，共 466 项：465 项通过，1 项跳过（`test_N6_C06_migration_rejects_symlink_alias`：未开启开发者模式，无权创建符号链接，与 P3 相同；跳过不计为通过）。`tests/test_win_word_faults.py` 10 项（原 7 项 + 会话 3 项）全部通过，结束后无 `WINWORD.EXE` 残留。推送前在全新克隆（`c60edfc`）中运行离线全量（不设 `DOCUMENT_SYNTHESIS_WORD_TEST`）：429 项，`OK (skipped=9)`；跳过的是需要真实 Word 的用例类与符号链接用例。
- **macOS**：`mac_applescript.py` 未改动；`office_session()` 对没有 `session()` 的后端原样执行。唯一的行为变化是构建元数据的 `word_automation.reason` 由探针结果（`Automation 探针成功: Microsoft Word。`）变为静态检查结果（`Microsoft Word 精确分页可用。`），构建不再为此执行 `tell application "Microsoft Word" to get name`。需要 macOS 会话跑一次离线全量与真实 Word 回归确认。
- **超时中的用户文档**：见 2.4 的局限。
- **会话外的调用**：`--doctor`、`lib.qa.export_docx_to_pdf`、直接调用 `inspect_document` 的测试仍是单次模式；如有需要可在调用方包一层 `office_session()`。
- **smoke 与构建之间**：两者是独立命令，各自启动一次 Word；矩阵中 smoke 每例约 6 秒，主要是这一次启动。跨命令复用需要常驻进程，不在本批次范围（见第 3 节第 2 条）。
