# P4 实施记录：Windows 真实 Word 验收矩阵

> **文档状态**：已完成 (Completed)；退出条件全部满足：WW01–WW12 在 Windows 上全部 `verified`，跨平台页码差异为 0，README 只列出已验证的 Word 版本  
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
| `docs/acceptance/public-release/macos-word/environment.json`、`ww-results.json` | macOS 环境记录与 WW01–WW12 结果（由 macOS 会话在 `8f53eca` 上生成） |
| `docs/acceptance/public-release/cross-platform-pages.json` | macOS 与 Windows 的逐用例页码对照 |

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
| 6 | Word 在关闭阶段卡住：插件加载的环境下，一次全量测试中 `test_m1_custom_format_demo_full_loop` 等满 600 秒超时。诊断日志显示测量与导出在 2 秒内完成，最后一步是 `Application.Quit`；调用栈停在 `handle` 返回处，即最后一次释放 COM 引用的跨进程调用在等待正在关闭的 Word | Windows | worker 在操作产物完成时记下结果；调用 `Quit` 时启动看门狗，30 秒内 worker 仍未退出就直接输出已完成的结果并结束 worker，父进程随后按 PID 结束卡住的 Word。结果先于 `CoUninitialize` 输出。等待从 600 秒失败变为约 45 秒成功 |
| 7 | 跨平台对照误报：重名书签改名为 `_SynthBm_<源文件绝对路径哈希>_<序号>`（`lib/renderers.py`），不同运行、不同机器必然不同；有/无插件两组对照时 WW09 因此被误判为差异，页码实际相同 | 对照工具 | `compare` 把路径哈希归一化为按页码排序的组号后再比较。书签名依赖本机路径这一点不影响页码，但意味着同一输入在不同机器上产出的 DOCX 内容不逐字节相同，记为观察项 |

## 4. 插件对照

本机 Word 加载两个第三方 COM 加载项：iSlide Tools 与 Acrobat PDFMaker。它们同时注册在 HKCU、HKLM 与 HKLM 的 WOW6432Node 下；实测只把 HKCU 的 `LoadBehavior` 改为 2 时，自动化实例仍然加载两者（为所有用户安装的加载项以 HKLM 为准）。因此“禁用插件”一组由用户在管理员终端运行 `word_addins.py disable`（先备份三处原值），对照结束后运行 `restore` 恢复；禁用后在隔离的自动化实例中确认两者均未连接，Microsoft 自带的加载项不受影响。

### 4.1 结果

| 环境 | 提交 | WW01–WW12 | 矩阵总耗时 | 全量测试（`DOCUMENT_SYNTHESIS_WORD_TEST=1`） | 挂起 |
| --- | --- | --- | --- | --- | --- |
| 插件按用户设置加载 | `561c9de` | 12 项 `verified` | 350 秒 | 429 项，1 项跳过（符号链接），1 项失败：仓库卫生检查命中本批次新增测试中的用户目录与邮箱字面量，已改为运行时拼接后单独复测通过 | 无 |
| 第三方插件临时禁用 | 中间版本（见下） | 12 项 `verified` | 568 秒 | 425 项，1 项跳过（符号链接），其余通过 | 无 |

- 两组的书签页码、导出 PDF 页数与页码标签逐项一致（12/12，经第 3 节第 7 条归一化后对照）。
- “禁用插件”一组运行于 `2ae6e63` 之上的工作区版本，已包含第 3 节第 1–4 条，尚未包含第 5、6 条；这两条只改变超时与关闭阶段卡住时的处理，不影响正常构建路径。禁用插件需要管理员权限，未在最终提交上重跑。
- 耗时受本机负载影响明显：同为“插件加载”，两次矩阵分别为 591 秒与 350 秒，不能据此判断插件对速度的影响。

### 4.2 挂起统计

| 环境 | 全量测试轮数 | 挂起次数 | 说明 |
| --- | --- | --- | --- |
| 插件加载 | 6（P3 三轮、P4 三轮） | 2 | P3 首轮：未取到调用栈；P4 第二轮：卡在 `Application.Quit` 之后的 COM 释放（第 3 节第 6 条），看门狗修复后的一轮未再出现 |
| 插件禁用 | 2 | 0 | 首轮的 2 项失败是第 3 节第 5 条的孤儿进程问题，不是挂起 |

样本很少，不能断定挂起只由插件引起，但两次挂起都发生在插件加载的环境中，且第二次明确卡在 Word 关闭阶段，与加载项在关闭时阻塞的表现一致。看门狗与 600 秒超时、诊断日志共同兜底：关闭阶段卡住时构建照常成功；其他阶段卡住时报错会给出最后一次 COM 调用与调用栈日志。

### 4.3 耗时（插件加载，`561c9de`）

| 步骤 | 12 个用例合计 | 单个用例范围 |
| --- | --- | --- |
| `--plan` | 7 秒 | 不足 1 秒 |
| `--doctor` | 55 秒 | 约 4–5 秒（启动一次 Word 读取版本） |
| 构建 | 161 秒 | 约 10–20 秒 |
| `smoke_test.py` | 68 秒 | 约 5–8 秒 |

Windows 后端每次 Office 操作都启动一个独立的隐藏 Word（约 4–6 秒），一次构建会因目录回填的多趟分页与构建元数据中的 Word 探针启动多次；macOS 后端连接已运行的 Word，没有这部分开销。一次构建复用一个 Word 实例、元数据不再启动 Word 的优化另行处理，不在本批次范围内。

## 5. 跨平台页码对照

| 项目 | macOS | Windows |
| --- | --- | --- |
| 系统 | macOS 27.0.1（arm64） | Windows 11 Pro 26H2（10.0.26300，x64） |
| Word | 16.113.4（16.113.26100421） | Microsoft 365 16.0.20430.20146（Current Channel，64 位） |
| 默认打印机 | 无（no system default destination） | Microsoft Print to PDF |
| 字体 | 8 个字体族全部检测为已安装 | 8 个字体族全部检测为已安装 |
| 结果文件提交 | `8f53eca` | `561c9de`（插件加载组） |
| WW01–WW12 | 12/12 `verified` | 12/12 `verified` |
| 矩阵总耗时 | 83 秒 | 350 秒 |

`compare macos-word windows-word` 的结论：12 个用例全部可比、全部一致，**差异为 0**。对照内容包括每个书签的物理页、打印页码与 PDF 中观察到的页码标签（共 49 个书签，WW06 与 WW10 各 10 个），以及 Word 导出交付物 PDF 的页数与逐页页码标签。两平台默认打印机不同（macOS 没有默认打印机），但没有影响分页。

macOS 侧还确认：两个 smoke 修复在 macOS 上同样生效（12 个用例的 smoke 步骤全部通过）；本批次改动了共用代码（`smoke_test.py`、`lib/delivery.py`），macOS 真实 Word 45 项基线全部通过，离线 402 项通过、15 项跳过（Windows 专属与未开启真实 Word 的用例）。Windows 侧在 `8f53eca` 之后只修改了 Windows 专属测试（`2a83560`、`aa1f57b`），不影响对照结论。

macOS 的字体按文件名匹配：楷体_GB2312 对应 `楷体_GB2312.ttf`，方正小标宋简体对应 `方正小标宋_GBK.ttf` 与 `方正小标宋简体.ttf`；Windows 按字体注册表名称匹配（`KaiTi_GB2312`、`FZXiaoBiaoSong-B05S`）。两个平台的字体来源不同，但本矩阵的分页结果一致。字体可移植性的系统处理属于 P6。

耗时差异来自后端设计：macOS 连接已运行的 Word；Windows 每次 Office 操作启动一个独立的隐藏 Word（约 4–6 秒），一次构建会启动多次。一次构建复用一个 Word 实例的优化另行处理。

## 6. 未覆盖与后续

- **较旧 Office 版本**：没有可用的 Office 2019/2021 环境。README 的支持范围如实写“仅验证过 Microsoft 365”。
- **插件禁用组的代码版本**：见第 4.1 节，运行于包含第 3 节第 1–4 条修复的中间版本，未在最终提交上重跑。
- **偶发挂起**：两次挂起都发生在插件加载的环境中，其中一次已定位到 Word 关闭阶段并由看门狗兜底；另一次（P3 首轮）未取到调用栈。之后遇到挂起时，按诊断日志中的最后一次 COM 调用与调用栈继续定位。
