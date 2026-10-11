# P5 实施记录：草稿模式（目录页码域）

> **文档状态**：已实施，macOS 已验证；Windows 复核：精确构建不受影响，排版完成后更新域的页码正确，但发现双击打开草稿时目录页码显示为 1 的问题，待修复（见第 6、7 节）  
> **基线依赖**：P4 与 Word 会话复用（`main` = `838db05`）  
> **核心目标**：按[执行计划](windows-and-public-release-plan.md)第 10 节，在没有可精确分页的 Microsoft Word 时（主要是 Linux），用 `--allow-draft` 生成草稿交付物：不调用 Word、不推算页码，目录页码由 Word 打开时更新；草稿标注不可关闭，永远不覆盖精确交付物。
> **范围决定（2026-10-11）**：只做页码域草稿，不实现 LibreOffice 后端。

---

## 1. 行为

| 场景 | 结果 |
| --- | --- |
| 能精确分页（macOS / Windows + Word） | 照常 `exact` 构建；`--allow-draft` 不起作用 |
| 不能精确分页，没有 `--allow-draft` | 构建失败，提示可用 `--allow-draft`；已有交付物不被替换 |
| 不能精确分页，有 `--allow-draft` | 生成 `*_draft.docx`，目录页码为 dirty 的 `PAGEREF` 域，打开时由 Word 更新 |
| 后端自称只能给出 `draft` 保真度的测量页码 | 仍一律拒绝（P2 的门禁不变）；草稿只来自“不测量、交给 Word 更新”这一条路径 |

保真度由 `lib/drafting.py` 的 `build_fidelity()` 决定：当前后端 `static_status()` 可用且为 `exact` 时总是 `exact`；否则只有显式 `--allow-draft`（环境变量 `DOCUMENT_SYNTHESIS_ALLOW_DRAFT=1`）才是 `draft`。

## 2. 实现

| 文件 | 改动 |
| --- | --- |
| `lib/drafting.py` | 新增：保真度判定、`*_draft` 文件名、草稿元数据文件名、终端提示 |
| `lib/composition.py` | `append_toc(..., draft=True)`：目录条目格式不变，页码换成指向书签、`w:dirty="true"` 的 `PAGEREF <书签> \h` 复杂域，更新前显示“?”；目录引用另一份交付物时只写“?”。`enable_update_fields_on_open()` 按 ECMA-376 `CT_Settings` 顺序写入 `w:updateFields` |
| `lib/field_updater.py` | `FieldUpdater(defer_pagerefs=True)`：草稿中保留 `PAGEREF` 原缓存，`SEQ`/`REF` 照常更新 |
| `lib/delivery.py` | `build_deliveries(draft=True)`：每份交付物只装配一次；执行结构核验与 v3 出现位置标记；跳过分页测量、目录收敛、`PAGEREF` 缓存核验与 PDF 页级 QA；引用另一份交付物的目录在终端说明页码保留“?” |
| `lib/engine.py` | 决定保真度并传给交付；v3 草稿元数据写入 `build-metadata_draft.json`（顶层与每个交付物均有 `fidelity`，`filename` 为实际发布的文件名）；终端最后一行为草稿提示。精确构建的元数据也新增 `fidelity: "exact"` |
| `smoke_test.py` | `--allow-draft`：检查 `*_draft.docx` 与草稿元数据，强制只做结构与文字核验，并确认 `updateFields` 已设置、没有残留占位标记；精确 smoke 拒绝 `fidelity` 不是 `exact` 的元数据 |
| `synthesize.py` | 新增 `--allow-draft` |
| `lib/office/null.py`、`lib/office/win_com.py` | 没有 Word 时的提示指向 `--allow-draft` |
| `.github/workflows/ci.yml` | 三平台（CI 没有 Word）对三个示例执行 `--allow-draft` 构建与 `smoke_test.py --allow-draft` |

## 3. 测试

`tests/test_draft_mode.py`（离线，三平台都运行）：

1. 没有 `--allow-draft` 时构建失败，已有交付物字节不变。
2. 草稿标注：只发布 `demo_complete_draft.docx`；每个目录条目一个 dirty `PAGEREF`；`updateFields` 已设置；没有占位标记残留；终端最后一行是草稿提示。
3. 草稿不覆盖同目录下已有的精确交付物。
4. 能精确分页（假后端）时 `--allow-draft` 不起作用，产出精确文件名且目录不含 `PAGEREF`。
5. v3 草稿：没有 `build-metadata.json`，`build-metadata_draft.json` 的 `fidelity` 为 `draft`，交付物记录无 `page_map`。
6. 拆分打印目录：目录在另一份交付物中时没有 `PAGEREF`，只有“?”，终端给出说明。
7. `smoke --allow-draft` 通过；把草稿改名冒充精确交付物时，精确 smoke 因保真度不符失败；删除 `updateFields` 后草稿 smoke 失败。
8. `FieldUpdater(defer_pagerefs=True)` 保留 `PAGEREF` 缓存、更新 `SEQ`；`updateFields` 排在 `compat` 之前且不重复。

`tests/test_office_backends.py`：无 Word 的提示现在应包含 `--allow-draft`。

## 4. 验证记录（macOS 27.0.1 arm64，Word 16.113.4，Python 3.14.0；Python 3.10 离线另测）

1. **离线全量**：450 项通过，19 项跳过（均为 Windows 专属或需手动开启的真实 Word 用例）；Python 3.10 且 `DOCUMENT_SYNTHESIS_OFFICE_BACKEND=none` 下同样 450 项通过、19 项跳过。
2. **三个示例的草稿构建与草稿 smoke**（`--office-backend none --allow-draft`）：全部通过。最小示例 2 个、中文示例 11 个目录条目均为 dirty `PAGEREF`；custom-format 示例只有正文部件，没有目录。
3. **WW01–WW12 夹具的草稿构建**：10 个准备函数（WW01、WW03–WW05、WW08、WW09、WW11、WW12、编号夹具、论文夹具）全部构建成功，没有占位标记残留；来源文档自带的 `PAGEREF` 保持原样，由 `updateFields` 一并更新。
4. **在 Word 中更新域后与精确构建对照**（退出条件）：
   - 中文示例：11 个目录条目，更新后页码 `1,1,1,2,2,2,2,3,3,3,3`，与精确构建逐项一致。
   - 论文夹具：7 个目录条目 `1,3,5,7,8,9,10` 一致；正文中来源自带的 `PAGEREF` 更新后为 `8`，与精确构建的缓存一致。
   - 同时用精确构建使用的测量方法核对了草稿文档的书签页码：物理页与打印页码和精确构建完全相同，说明草稿与精确构建的版面一致。
   - 注意：通过 AppleScript 自动化时，Word 只做 `repaginate` 就更新域，会在排版完成前求值，所有页码都得到 `1`；先查询一次文末位置的页码迫使 Word 排完全部页面后再更新，结果才正确。正常打开文档时版面已排好，不受影响；但这说明在 Word 排版完成前更新域会得到错误页码，README 的提示是“打开后按提示更新域（或全选后按 F9）”。
5. **精确构建回归**：见第 5 节。

## 5. 精确构建回归（macOS）

本批次改动了共用的装配、交付、引擎、元数据与 smoke 代码，需要重新跑真实 Word 基线：

- macOS 真实 Word 45 项基线：全部通过。
- WW01–WW12（`--output ww-results-p5.json`）：12/12 `verified`；与已合并的 `macos-word/ww-results.json` 对照 `page_map`（物理页、打印页码、实际页码标签，`_SynthBm_` 哈希归一化）、导出页数与页码标签，12/12 一致。该结果文件未入库。

## 6. Windows 复核（Windows 11 Pro 26H2 build 26300，Microsoft 365 Word 16.0.20430.20146 x64 Current Channel，Python 3.14.4，提交 `13dc4c4`）

在独立 worktree 中执行；本机 `templates/` 下没有私有模板。

1. **离线全量**：440 项通过，9 项跳过：8 项需要 `DOCUMENT_SYNTHESIS_WORD_TEST=1`，1 项为符号链接用例（本机未开启开发者模式，无权创建符号链接）。
2. **真实 Word 全量**（`DOCUMENT_SYNTHESIS_WORD_TEST=1`）：477 项通过，1 项跳过（同一个符号链接用例，属于环境限制，不算通过）。
3. **WW01–WW12**（`--output ww-results-p5.json`，未入库）：12/12 `verified`。直接用 `word_matrix.compare_results` 与已入库的 `windows-word/ww-results.json`（`561c9de`）对照（没有使用 `compare` 子命令），`page_map`（物理页、打印页码、实际页码标签，`_SynthBm_` 哈希归一化）、导出页数与页码标签 12/12 一致。
4. **草稿在 Word 中更新域后与精确构建对照**：中文示例与论文夹具（`prepare_thesis`）各做一次精确构建与精确 smoke，以及一次 `--office-backend none --allow-draft` 草稿构建与草稿 smoke，全部通过。草稿中目录条目分别有 11 个、7 个 dirty `PAGEREF`，更新前显示“?”。用独立 Word 实例（`DispatchEx`，`DisplayAlerts=0`）打开草稿副本，按三种方式更新域，保存后按 `_Synth_entry_` 书签读取目录页码：

   | 更新方式 | 中文示例（精确：`1,1,1,2,2,2,2,3,3,3,3`） | 论文夹具（精确：`1,3,5,7,8,9,10`） |
   | --- | --- | --- |
   | 先对文末取 `Range.Information(3)` 迫使排版完成，再 `Fields.Update()` | 一致 | 一致；正文来源自带的 `PAGEREF` 为 `8`，与精确构建相同 |
   | 先 `Repaginate()`，再 `Fields.Update()` | 一致 | 一致 |
   | 打开后立即 `Fields.Update()` | `1,1,1,2,2,2,2,2,2,2,2`，不一致 | `1,2,2,2,2,2,2`，不一致 |

   与 macOS 不同：Windows 上 `Repaginate()` 能让排版完成，之后更新域的结果正确。共同点是：排版完成前更新域会得到错误页码。目录条目标题三种方式都与精确构建一致。

5. **像用户一样打开草稿（发现问题）**：
   - 用 `os.startfile` 打开中文示例草稿：Word 弹出对话框“This document contains fields that may refer to other files. Do you want to update the fields in this document?”，默认焦点在 **No**（截图确认）。本次观察中，对话框大约 2.6 秒后在**没有人、也没有程序点击**的情况下消失，原因未查明。
   - 文档打开后，目录 `PAGEREF` 已经从“?”变为全部 `1`，**不是正确页码**；排版完成（文末页码为 5）并等待数分钟后，仍然全是 `1`。论文夹具相同：8 个 `PAGEREF`（含正文来源自带、正确值为 8 的那个）全部为 `1`。
   - 原因定位：用 COM 打开四种变体，只要保留 `w:dirty="true"` 或 `w:updateFields` 中的任意一个，打开时就会在排版完成前更新域，结果全是 `1`；两个都去掉才保持“?”。也就是说，打开时的这次更新发生在分页之前，与第 4 项“打开后立即更新”的错误同源。
   - 补救：在用户打开的 Word 中排版完成后全选并更新域（用 `Content.Fields.Update()` 代替 F9），中文示例和论文夹具的页码都与精确构建一致。
   - **未验证**：在对话框中人工点击“是”之后的页码。用户不在场，没有人工点击；也没有用程序代替点击。
   - 影响：草稿在 Windows 上被直接打开时，目录显示的是看似合理但错误的页码 `1`，而不是醒目的“?”。用户不按 F9 就不会发现。README 中“打开后按提示更新域”的说法在 Windows 上不足以得到正确页码。按约定，本次只记录问题，没有修改代码。

## 7. 未完成与后续

1. **修复双击打开时页码显示为 1 的问题**（见第 6 节第 5 项）：需要决定修复方式，例如不在打开时自动更新，改为明确提示用户全选后按 F9，或采用其他能保证在分页之后求值的做法；修复后需在 Windows 与 macOS 上重新验证交互式打开。
2. **交互式打开的人工确认**：Windows 上弹出的“更新域”对话框中点击“是”之后的结果、以及对话框无人点击自行消失的原因，都尚未人工确认；macOS 上交互式打开也没有人工验证。
3. **已发现的独立问题（不属于 P5）**：中文示例清单中的封面模板路径按清单目录解析后不存在，引擎静默改用全局 `templates/` 目录中的任意 `.docx`；在有私有模板的本机上，中文示例的精确 smoke 因多出两张图片失败。已作为独立任务记录，修复后需重新核对依赖中文示例输出的证据。
