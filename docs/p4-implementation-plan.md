# P4 实施记录：Windows 真实 Word 验收矩阵

> **文档状态**：Windows 侧已完成；跨平台页码对照待 macOS 运行同一矩阵后补齐  
> **基线依赖**：P3（Windows COM 后端，`main` = `2ae6e63`）  
> **核心目标**：按[执行计划](windows-and-public-release-plan.md)第 9 节，在 Windows 上建立与 NW01–NW12 对应的 WW01–WW12，按 plan → doctor → build → smoke 串行执行；记录环境；与 macOS 对照同一夹具的书签页码；覆盖较旧 Office 版本或如实说明；在有/无第三方 COM 插件两种环境下复现并定位 P3 中出现过的挂起。

---

## 1. 交付物

| 文件 | 内容 |
| --- | --- |
| `docs/acceptance/public-release/word_matrix.py` | 跨平台矩阵脚本：`run`（WW01–WW12）、`environment`（环境记录）、`compare`（跨平台页码对照） |
| `docs/acceptance/public-release/word_addins.py` | 插件对照工具（仅 Windows）：备份并禁用第三方 Word COM 加载项、按备份恢复、在隔离实例中核对加载状态 |
| `docs/acceptance/public-release/windows-word/environment.json` | Windows 环境记录 |
| `docs/acceptance/public-release/windows-word/ww-results.json` | WW01–WW12 结果（插件按用户设置加载） |
| `docs/acceptance/public-release/windows-word/ww-results-addins-disabled.json` | WW01–WW12 结果（第三方插件临时禁用） |
| `tests/test_public_release_matrix.py` | 对照逻辑与证据脱敏的离线测试 |

证据中的仓库路径、临时目录与用户目录替换为 `<repo>`、`<work>`、`<tmp>`、`<user-home>`，不含个人信息；交付物 DOCX/PDF 不入库，只记录 SHA-256、页码与检查结果。

## 2. 矩阵设计

WW01–WW12 与 `tests/test_post_review_word.py` 的 NW01–NW12 使用同一组匿名夹具（`docs/acceptance/create_word_followup_fixtures.py`、`create_complete_thesis_fixture.py` 与测试中的内联夹具），区别在于：

1. 每项都经生产 CLI 串行执行 `--plan`、`--doctor`、构建、`smoke_test.py`；NW 只跑构建，不跑 smoke。
2. 构建后用当前平台的 Word 再导出一次交付物 PDF，记录页数与 PDF 页码标签；连同 `build-metadata.json` 中的书签 `page_map`，作为跨平台对照的依据。
3. 状态口径：四步退出码全为 0 且全部检查通过才是 `verified`；doctor 报告精确 Office 不可用为 `blocked`；其余为 `failed`。skip 不算通过。

跨平台对照只比较 `physical_page`、`printed_page`、`observed_label`，以及导出 PDF 的页数和页码标签；核验状态等与平台无关的字段不参与对照。任一平台不是 `verified` 的用例不对照，并在报告中注明。

## 3. 实施中发现并修复的问题

| # | 问题 | 影响平台 | 修复 |
| --- | --- | --- | --- |
| 1 | `smoke_test.py` 调用 `validate_delivery` 时没有传 `page_sequences`；配置了 `layout.page_sequences` 的项目在 smoke 中报 `'NoneType' object has no attribute 'get'` | 全平台（NW 不跑 smoke，此前未暴露） | smoke 与构建使用同一组页码序列；`_enrich_page_map_labels` 在没有序列表时退回空表。回归测试见 `tests/test_page_sequences.py` |
| 2 | `docx_document` 的 `source.file` 位于子目录（如完整论文夹具的 `source/thesis-source.docx`）时，smoke 按“来源目录 + 节点文件名”拼接，报“源文件不存在” | 全平台 | smoke 对 `docx_document` 策略改用构建同一个 `resolve_docx_document_path`。回归测试见 `tests/test_smoke_outputs.py` |
| 3 | P3 的 COM 消息过滤器从未生效：pywin32 没有 `CoRegisterMessageFilter`（也没有 `IID_IMessageFilter`），注册时的异常被日志吞掉 | Windows | 改为在调用层识别 `RPC_E_CALL_REJECTED`、`RPC_E_SERVERCALL_RETRYLATER` 并退避重试（最长约 60 秒），覆盖打开、分页、读取页码、导出、另存、PowerPoint 打开与导出 |
| 4 | 超时后无法知道挂在哪一步 | Windows | worker 在每次可能阻塞的 COM 调用前把步骤写入 stderr，并在超时前 3 秒用 `faulthandler` 导出全部线程的 Python 调用栈；父进程超时后把 stderr 保存为 `<工作目录>/diagnostics/timeout-*.log`，报错中给出最后一次 COM 调用与日志路径 |
| 5 | Word 卡在启动阶段时超时：worker 在 `DispatchEx` 返回前被结束，父进程立即查找孤儿进程，但 DCOM 可能在数秒后才把 Word 拉起，这个没有客户端的 Word 会一直残留。禁用插件后整体变慢，“强制结束 Word”两个用例在全量测试中因前一个 1 秒超时用例的迟到进程而失败 | Windows | 超时或 worker 异常退出且 pid 文件只有启动前快照时，父进程在 30 秒内持续查找本次启动的孤儿进程（仍须满足“快照之后新出现、命令行带 `/Automation -Embedding`、唯一”）。实测三次 1 秒超时中有两次孤儿进程在超时后约 7.6 秒才出现，均被结束。故障测试中的超时用例改为超时后再观察 15 秒，确认没有迟到的残留 |

## 4. 插件对照

本机 Word 加载两个第三方 COM 加载项：iSlide Tools 与 Acrobat PDFMaker。它们同时注册在 HKCU、HKLM 与 HKLM 的 WOW6432Node 下；实测只把 HKCU 的 `LoadBehavior` 改为 2 时，自动化实例仍然加载两者（为所有用户安装的加载项以 HKLM 为准）。因此“禁用插件”一组由用户在管理员终端运行 `word_addins.py disable`（先备份三处原值），对照结束后运行 `restore` 恢复；禁用后在隔离的自动化实例中确认两者均未连接，Microsoft 自带的加载项不受影响。

<!-- ADDIN-RESULTS -->

## 5. 未覆盖与待办

- **较旧 Office 版本**：没有可用的 Office 2019/2021 环境。README 的支持范围如实写“仅验证过 Microsoft 365”。
- **跨平台页码对照**：需在 macOS 上运行 `python3 docs/acceptance/public-release/word_matrix.py run --label macos-word` 与 `environment --label macos-word`，提交结果后运行 `compare macos-word windows-word`。按计划，字体完全一致时页码应相同；出现差异要逐项记录原因（字体替换、Word 版本、打印机驱动或默认纸张），不能直接接受为“正常差异”。macOS 的字体检测按字体文件名匹配，若与 Windows 注册表名称不一致，需在 macOS 侧补充别名。
