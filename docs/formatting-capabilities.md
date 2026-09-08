# OOXML 排版能力矩阵与高风险技术验证

日期：2026-09-06。历史 R11 状态为 `verified`；post-review N0–N10 当前状态见 [`acceptance/post-review/status.json`](acceptance/post-review/status.json)，当前声明范围整体为 `verified`（N7 独立匿名冻结集已达到自动接受门槛，N8 已通过真实 Word 矩阵，N9 已通过已知 Quartz orphan-xref 诊断）。本轮离线与真实 Word 验证记录见 [`acceptance/post-review/offline-verification.json`](acceptance/post-review/offline-verification.json)。

本表区分“模块已有代码”“端到端证据”和“当前可对外声明的能力”。B0 探针、离线装配测试和真实 Word 测试不是同一层证据；当前只对有对应契约和真实 Word/PDF 证据的窄范围能力使用 `supported`，不把它扩展为全部学校规范或完整匿名论文的无条件支持。

## 1. 运行环境与证据边界

- 保留的真实 Word 补测记录见 [`docs/acceptance/word-followup/`](acceptance/word-followup/)；W04–W06 的输入文件哈希固定在 [`r0_contracts.json`](../tests/fixtures/formatting/r0_contracts.json)，W07 使用独立的完整匿名论文夹具。
- F01–F09 的新回归通过离线夹具或替换 Word 边界验证，测试名称明确包含 `integration without Word`，不计入真实 Word 验收；NW01–NW12 的真实 Word 结果单独记录在 [`n8-status.json`](acceptance/post-review/n8-status.json)。
- `tests/test_thesis_acceptance.py` 是装配级回归；`tests/test_post_review_word.py` 是本轮真实 Word/PDF 验收组。真实 Word 测试统一使用固定 `DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access`；Word 环境不可用时，结果为 `BLOCKED`，不等同于通过。

## 2. 当前能力矩阵

| 能力标识 | 模块实现 | 端到端验证 | 当前可声明状态 | 证据/缺口 |
| :--- | :--- | :--- | :--- | :--- |
| `styles.fonts.mixed` | `implemented` | B0 探针；R4/R7 离线 preserve/mixed 回归；M1 真实 Word 格式闭环 | `preserve_only` | 真实交付已验证样例复用与格式门禁，但尚未建立独立混合字体 PDF 基线，因此不扩展为全量 restyle 支持。 |
| `fields.complex` | `implemented` | R9 15 项字段回归（含跨 run/跨段与前向 REF）；W06 真实链路验证受控 PAGEREF | `preserve_only` | 复杂字段结构与依赖更新已有离线证据；真实 Word 复杂字段专项尚未单独纳入，因此只声明保留/受控更新边界。 |
| `notes.footnotes` | `implemented` | F08/R6/R9 离线通过；R5 选区闭包通过 | `preserve_only` | 脚注/尾注引用、定义、ID 重映射和关系闭包已验证；完整跨文档注释的真实 Word/PDF 基线仍缺。 |
| `pagination.roman` | `implemented` | R8 离线回归；NW07 真实 Word/PDF 标签与序列验证 | `supported`（NW07 窄范围） | 当前矩阵已覆盖罗马前置序列；不扩展为所有学校页码规范的无条件支持。 |
| `layout.odd_page` | `implemented` | R8 离线回归；NW07 真实 Word/PDF 页级 QA | `supported`（NW07 窄范围） | 当前矩阵已覆盖 oddPage 物理奇偶与白名单空白页；不扩展为任意复杂分节布局。 |
| `numbering.multilevel` | `implemented` | R4/R5/R6 离线；W05 真实 Word 交付 | `preserve_only` | 原生编号在导入/重排中保留；尚未对所有多级编号变换发布通用支持声明。 |
| `layout.parts.v1` | `implemented` | F01/F07/R6/R9 离线；W05/W06 真实交付使用 | `preserve_only` | 选区关系闭包、ExpectedInventory 范围和 ID 映射已验证；复杂 story 的全量真实矩阵仍缺。 |
| `fields.managed_update.v1` | `implemented` | R9 离线；NW01/NW02/NW10 真实 Word/PDF 受控更新与字段指令验证 | `supported`（NW01/NW02/NW10 窄范围） | 当前矩阵覆盖无目录/有目录 PAGEREF、对象字段指令和最终缓存；已知 Quartz orphan-xref 警告由 N9 非破坏性分类，其他 PDF 结构异常仍阻断。 |

状态含义：`implemented` 只说明存在模块代码，`preserve_only` 只允许保护既有语义，`unsupported` 不允许作为已验证交付能力，`blocked` 表示外部环境未提供必要证据，`supported` 仅在完整契约、真实 Word/PDF 和发布门禁均通过后使用。

## 3. R0 契约证据

| 范围 | 固定资产 | 当前结果 |
| :--- | :--- | :--- |
| F01–F10 | `tests/fixtures/formatting/r0_contracts.json`、`test_formatting_pipeline_contracts.py`、`test_publication_gates.py` | 契约测试已全部转为正常回归并纳入当前 post-review 286 项全套；F05/F06 的页码/字段反例已由 R8/R9 修复后关闭，F03/F04/F07/F08/F09 继续通过。 |
| W04–W07 | `docs/acceptance/word-followup/m2-fixtures/`、`docs/acceptance/create_complete_thesis_fixture.py`、`test_word_formatting.py` | 已在 `DOCUMENT_SYNTHESIS_WORD_TEST=1` 的本机 Word 环境执行并通过 4/4；W07 覆盖完整匿名论文 DOCX/PDF；未满足该环境时仍应记录为 BLOCKED。 |

## 4. 门禁与实现约束

1. 只有 required 检查全部真实执行并通过，交付物才可发布；`not_run`、`failed`、`unsupported` 都不能折算为 `passed`。
2. `preserve_only` 能力若需要结构性改写，必须在 plan/预检阶段明确阻止，不能静默转换或回退全文。
3. 格式包中的 `required_capabilities` 必须与当前引擎支持注册表和验收证据一致；规划中、未接入和未验证能力不能提前宣传为 supported。

## 5. R3 证据

- `tests/test_r3_build_plan.py` 覆盖共享准备、公开计划正文脱敏、显式 `heading.2` 样式与导航、重复标题、过期 RoleMap 和渲染前源篡改。
- `lib/build_plan.py` 统一维护输入/格式/映射哈希、原始 NodeRef、最终 assignments 和交付声明；`lib/engine.py` 在创建 run 目录前完成准备与校验。
- 当前状态仍只对 preserve/mixed 的已验证边界作声明；R10 已补齐 M1 与 W04–W07 的真实 Word/PDF 证据，多来源复杂 story 的全量矩阵不在本次支持声明内。

## 6. R4 证据

- `RenderContext` 已贯穿 v3 单文件、多来源递归、DOCX/PDF/图片/PPTX 导入和横竖版恢复；旧 `fonts` 与 `LAST_RENDERED_LANDSCAPE` 仅保留为兼容入口。
- `apply_section_spec()` 实际写入纸张、方向、边距、页眉页脚距离和 `docGrid`；preserve/source 复制源节几何与简单页眉页脚，restyle/target 使用目标版心。
- preserve 跳过源样式和直接格式清洗，mixed 按文件与 `node_range` 限制托管范围；生成标题、目录条目、页脚、图片和表格统一使用当前格式上下文。
- F03 stale RoleMap、F04 多来源 37 pt 标题与 preserve/source 几何页脚契约已通过；新增页面规格、preserve 无变更和 mixed 隔离测试。
- 离线回归与 W04 真实 Word/PDF 页面几何、页脚和发布门禁均已通过；W04–W07 统一真实 Word 记录见 [`r8-status.json`](acceptance/r8-status.json) 与 [`r10-status.json`](acceptance/r10-status.json)。

## 7. R5 证据

- 选区按 `[start,end)` 解析；`p`、`tbl`、`sdt` 及无文字图片/公式/字段/脚注引用均纳入实质对象判定。重叠、非法边界、跨字段/书签边界和未知顶层对象均拒绝继续装配。
- `slice_document_by_region()` 通过 `PackageImporter` 导入源元素及关系闭包；`ImportResult` 记录节点、书签、关系、样式、编号和脚注/尾注映射，目标 `sectPr` 保持合法顺序。
- 关系校验覆盖 rId 存在性、关系类型、目标部件和关键内容类型；脚注内部媒体/XML 部件会克隆进目标包，不再直接引用源包。
- F01 图片/样式/关系闭包、F07 图片漏选/非法选区和新增跨边界、未知对象、选择性脚注测试已通过。R5 详细命令与哈希见 [`r5-status.json`](acceptance/r5-status.json)。

## 8. R6 证据

- `ExpectedInventory` 以 `PreparedBuild.source_order`、`parts`/`source_region` 和源 DOCX 清单为输入，支持多文件按声明顺序合并，并把生成文字、替换映射、合法 ID 重映射、来源 provenance 与 coverage 结构化保存。
- 内容门禁严格使用段落多重集、表格拓扑/单元格位置/重复数、媒体实例、OMML 实例、脚注/尾注定义与引用、字段指令/依赖、超链接真实目标和书签名称；受控字段缓存更新不改变正文要求。
- 格式门禁使用 `EffectiveStyleEvaluator` 的最终有效值，包含文档默认、样式继承、直接格式、主题字体、显式 false、每个 run、页面几何、缩进、行距和 keep/widow；每项报告都保存 expected/actual/provenance/coverage。
- engine 与 smoke 共用同一构造器和校验器；R6 全量离线回归为 `225` 项运行、`222 passed`、`3 expected failures`、`2 skipped`、`0` 非预期失败、`0` 错误。证据哈希和命令见 [`r6-status.json`](acceptance/r6-status.json)。

## 9. R7–R10 历史能力证据与当前边界

> 本节保留 R7–R10 的历史批次记录，用于追溯当时的夹具与环境；它们不覆盖当前 post-review N0–N10 的发布结论。当前应以 [`acceptance/post-review/status.json`](acceptance/post-review/status.json) 和 [`n8-status.json`](acceptance/post-review/n8-status.json) 为准；N8 的 NW01–NW12 已有真实 Word 证据，N9 已对当前复现的 Quartz orphan-xref 模式完成诊断。

当前补全批次的发布边界以 [`../post-n0-n10-remediation-plan.md`](post-n0-n10-remediation-plan.md) 和 [`remediation/status-summary.json`](acceptance/remediation/status-summary.json) 为准。历史 R7 的人工决策流程仍可追溯；S5 已在声明范围内完成真人审阅与跨目标交付，证据见 [`remediation/s5-human-review-evidence.json`](acceptance/remediation/s5-human-review-evidence.json)。20 份冻结集的 aggregate `manual_correction_count` 仍为 null，不能把单次真人操作外推为冻结集人工一致性保证。

- **R7**：`tests.test_r7_workflows` 8 项通过。样例分析、人工决策、虚构近似预览、manifest 迁移和跨目标格式/RoleMap 复用已形成闭环；独立匿名评估集 20 份、100 个标题正例、500 个正文负例，各级 precision/recall 1.0，正文误报 0；原始样例正文不会进入格式包。
- **R8**：`tests.test_page_sequences tests.test_document_parts tests.test_delivery_pipeline tests.test_formatting_pipeline_contracts` 54 项通过；W05 真实 Word/PDF 通过。`PageRecord` 的 expected/observed label、序列重启和 oddPage 空白页白名单均进入门禁。
- **R9**：`tests.test_r9_fields_and_notes tests.test_notes_and_fields` 15 项通过；W06 真实 Word/PDF 通过。复杂字段仍按能力边界管理，未知字段保留缓存，不被擅自重算。
- **R10**：离线全套 `246` 项为 OK（2 项条件跳过）；M1 真实 Word 3/3、W04–W07 真实 Word/PDF 4/4、M2 装配级论文 1/1 及完整匿名论文真实 Word/PDF 1/1。详细批次证据见 [`r7-status.json`](acceptance/r7-status.json)、[`r8-status.json`](acceptance/r8-status.json)、[`r9-status.json`](acceptance/r9-status.json) 和 [`r10-status.json`](acceptance/r10-status.json)。

## 10. R11 发布声明索引

> 下表是历史 R11 批次的声明索引，用于追溯当时的样例与环境；当前 post-review 的 NW01–NW12 已完成逐项真实 Word 复核。N7 的自动接受结论限于独立匿名冻结集，不把 N9 的已知诊断模式扩展为所有 PDF 结构的通用兼容保证。

R11 将历史批次中的每项声明绑定到测试 ID、支持级别和最近证据；其中的 `supported` 仅在当时声明的窄范围内成立。

| 声明范围 | 测试 ID | 支持级别 | 最近证据 |
| :--- | :--- | :--- | :--- |
| 样例分析、人工决策、格式包与虚构近似预览 | `tests.test_r7_workflows` | `supported`（流程闭环） | [`r7-status.json`](acceptance/r7-status.json)、[`examples/custom-format-demo/README.md`](../examples/custom-format-demo/README.md) |
| RoleMap 源哈希绑定与未映射节点门禁 | `tests.test_r3_build_plan`, `tests.test_r7_workflows`, `tests.test_r11_release` | `supported`（显式映射范围） | [`r11-status.json`](acceptance/r11-status.json) |
| v1/v2 清单迁移兼容 | `tests.test_r7_workflows`, `tests.test_r11_release` | `preserve_only`（非破坏迁移） | [`r7-status.json`](acceptance/r7-status.json)、[`r11-status.json`](acceptance/r11-status.json) |
| 自定义格式 v3 真实 Word/PDF 交付与完整元数据 | `tests.test_r11_release`, `smoke_test.py` | `supported`（当前示例/环境范围） | [`r11-status.json`](acceptance/r11-status.json)、`build-metadata.json` |
| 复杂字段、脚注/尾注、罗马页码、oddPage 与多级编号 | `tests.test_r9_fields_and_notes`, `tests.test_word_formatting` | 依各行矩阵的 `supported` 或 `preserve_only` | [`r8-status.json`](acceptance/r8-status.json)、[`r9-status.json`](acceptance/r9-status.json)、[`r10-status.json`](acceptance/r10-status.json) |

上述证据不覆盖 OCR/PDF 样例学习、双栏期刊、参考文献著录转换、多样例融合或全部学校规范认证；完整匿名论文的 W07 结果和 R11 自定义示例只证明各自夹具、配置和环境下的端到端闭环。
