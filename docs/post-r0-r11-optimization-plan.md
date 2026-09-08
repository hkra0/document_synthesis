# R0–R11 验收后的优化执行计划

日期：2026-09-06。状态：**已完成（声明范围内）**。N0–N10 的相应证据已完成：N7 独立匿名无大纲冻结集达到自动接受门槛，N8 已通过真实 Word 集成矩阵，N9 已完成已知 Quartz orphan-xref 警告诊断。批次编号为 N0–N10，与历史 R0–R11、缺陷 C01–C08 区分。机器可读状态见 [`acceptance/post-review/status.json`](acceptance/post-review/status.json)，离线与真实 Word 执行记录见 [`acceptance/post-review/offline-verification.json`](acceptance/post-review/offline-verification.json)。

本计划以 [R0–R11 独立验收报告](r0-r11-acceptance-report.md) 和当前代码为依据。目标是关闭已复现的错误发布与验证缺陷，完成可审计的自定义格式交付，再提高无结构样例的识别能力。已实施内容以对应 N 批次状态和测试/产物证据为准；若后续扩大范围或出现新异常，仍按 blocked/partial 记录，不能用状态文件替代证据。

当前离线复核基线：现有测试运行 286 项，`OK (skipped=3)`；干净虚拟环境同样为 286 项通过、3 项条件跳过。在固定 `DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access` 后，NW01–NW12 的最终真实 Word 矩阵为 12/12 通过（72.668 秒），没有新的授权弹窗；此前随机临时目录导致的 NW02 一次超时作为历史观察保留。同源真实 Word→PDF 导出的 Quartz wrong pointing object offset 警告已通过 xref 与跨解析器对照确认是未引用 orphan 条目，N9 不再阻断这一已知模式，但其他 PDF 结构异常仍按门禁处理。

## 1. 交付目标与范围

本轮交付分成三个里程碑。

| 里程碑 | 用户可获得的结果 | 放行条件 |
| --- | --- | --- |
| A：交付可信 | 错误字段、错误封面、漏检节点、错序/重复正文不会发布；合法保留格式能够通过 | N0–N6 完成，N8 中对应真实 Word 用例通过；既有 M1 兼容回归通过 |
| B：学习流程可用 | 结构化与手工格式分别识别、展示证据、允许校正，格式包可应用于另一稿件 | N7 完成，学习闭环真实 CLI/Word 验证通过；按实测区分辅助校正与自动识别范围 |
| C：论文范围可声明 | 已声明的选区、前置部件、页码、横版、页眉页脚及对象字段均有具体证据 | N8–N10 完成；完整矩阵中的必需用例无 skip、无未解释偏差 |

本轮保持 Python/现有 CLI 架构，继续使用 `synthesize.py` 作为构建入口。先修复当前实现，不引入新的 Web 后端、云端模型服务或另一套排版引擎。学校特定论文规范、自动引文编排、任意复杂 Word 对象等不自动纳入支持声明；只有明确测试覆盖的能力才可放行。

用户已授权 Word 访问匿名验收临时目录。实施中使用新的匿名输出目录，按 plan → doctor → build 执行，无需再次请求相同目录访问授权。运行 Word 测试时顺序执行。

## 2. 批次与依赖

工作量是相对尺度：S 为局部边界修复，M 为跨模块接线，L 为数据契约或多场景验证；不是固定工期承诺。

| 批次 | 内容 | 缺陷关联 | 依赖 | 工作量 | 退出标志 |
| --- | --- | --- | --- | --- | --- |
| N0 | 固定失败断言、基线及状态 | C01–C08 | 无 | S | 原反例稳定失败，正常对照通过 |
| N1 | 字段依赖与分页收敛 | C01 | N0 | M | 无目录 DOCX/PDF 引用一致，未收敛拒绝发布 |
| N2 | v3 封面来源与生成内容契约 | C02 | N0 | M | 默认不吸入旧模板，生成字段按部件验证 |
| N3 | 预期清单与输出节点映射 | C03/C04/C05 | N0 | L | 预期由源计划决定，映射覆盖每个预期实例 |
| N4 | 按策略验证格式与保护区 | C03/C04 | N3 | M | preserve/mixed 正确通过，70 pt 漏检反例被拒 |
| N5 | 有序内容与生成内容精确核验 | C02/C05 | N2/N3 | M | 错序、重复、额外来源实例均被拒 |
| N6 | 非破坏迁移保护 | C06 | N0 | S | 同文件及别名拒绝，源哈希不变 |
| N7 | 手工格式候选、校正与独立评测 | C07 | N0；最终闭环依赖 N4/N5 | L | 分层指标与校正闭环满足本计划门槛 |
| N8 | 真实 Word 集成矩阵 | C01–C05/C08 | N1/N2/N4/N5/N6；学习用例依赖 N7 | L | 必需场景逐项有最终成果断言 |
| N9 | PDF 解析警告诊断与导出证据 | 验收报告第 6 节 | N0；最终复核使用 N8 成果 | M | 警告已解释或相关能力明确阻断 |
| N10 | 发布回归、示例与能力状态同步 | C08及全部缺陷 | N0–N9 | M | 证据、CLI、元数据、文档一致 |

建议提交顺序为 N0 → N1 → N2 → N3 → N4 → N5 → N6 → N7 → N8 → N9 → N10。N6、N9 的调查可提前完成；不要等待识别算法优化才处理错误发布。每批可分为“失败断言 → 最小修复 → 边界回归”三个可审阅提交，按当前分支管理约定执行，不覆盖既有未提交改动。

## 3. 共同数据契约

### 3.1 在预期与实际之间建立可追踪关系

当前 `PreparedBuild` 已有来源哈希、RoleAssignment、选区和交付清单，`RenderedBody` 只有 `path/nodes`；最终验证仍从输出样式猜角色。建议新增 `lib/verification_contracts.py`，承载以下内部数据类型，并扩展已有类，不另建一条构建流程。

| 数据类型（拟新增） | 必需字段 | 产生时机与作用 |
| --- | --- | --- |
| `SourceOccurrence` | source_sha256、part_uri、element_path、delivery_id、part_id、occurrence_index、source_order_index | 准备阶段产生；同一源节点合法复用两次是两个实例，不能按文字或源哈希去重 |
| `ExpectedNode` | occurrence_id、role、story_type、语义摘要、受控属性、保护属性、provenance | 从来源检查和已解析策略生成，不能读取渲染结果反推期望 |
| `ExpectedDelivery` | 有序来源实例、预期节、生成内容许可、字段依赖、必需检查 | 每份交付物独立，不把整篇来源当作每份交付物的预期 |
| `OutputNodeMap` | occurrence_id → 输出 part_uri/结构位置/文本区间，变换与 ID 重映射记录 | 渲染与装配时逐步维护；保存后重新定位、核验，支持一对多合法变换 |
| `VerificationContext` | ExpectedDelivery、OutputNodeMap、配置与源哈希、格式包哈希 | 同时交给 engine 与 smoke；缺少上下文不能声称完成全面格式验收 |
| `FieldObservation` | 稳定 field_id、story、指令、目标、缓存、测量记录、状态 | 每轮收敛与最终 DOCX 核对使用 |

`occurrence_id` 由来源 NodeRef、交付部件及复用序号生成，不能用正文文字作为唯一键。输出路径会因前插封面、表格导入而改变，应在实际装配结束时重新建立索引。必要时使用不影响排版的内部书签辅助定位；必须处理书签冲突并验证关系闭包，不能在作者文字中插入可见标签。

行内属性以逻辑文本区间记录，不能假设 run 数量和边界在导入、字段更新后不变。表格按表/行/列和合并拓扑定位；脚注、尾注、页眉页脚使用各自 story 地址。映射只用于找到实际对象，独立的内容/格式检查仍须检查对象本身，不能把“映射记录存在”当作对象存在。

### 3.2 验证接口与覆盖分母

拟新增内部统一入口：

```python
build_verification_context(prepared, rendered, delivery_spec) -> VerificationContext
verify_delivery_artifact(docx_path, context) -> DeliveryReport
verify_final_fields(docx_path, field_index, measured_pages) -> CheckResult
```

`verify_delivery_format()` 可保留兼容适配入口，但正式 engine/smoke 必须传入完整上下文。旧入口缺少预期映射时，只能返回有限检查结果；不得伪造全面覆盖。

覆盖率的分母来自 **ExpectedDelivery**，包括应改写和应保留的属性。`expected=5, verified=4` 必须失败；合法 preserve 不能因为“没有目标格式托管角色”而失败，应验证其来源属性。零内容的合法目录/封面交付，按对应生成内容和页面检查判断，不硬编码 `verified_count>0`。

沿用现有 `CheckResult` 的 `not_run/passed/failed/unsupported` 状态，属性级未定位或未求值记录为 `unverified` 诊断。必需检查中存在未验证项时，汇总状态应为失败或未执行，不能用 optional 绕过必需能力。元数据新增契约版本、expected/verified/missing 数量、失败 NodeRef 和 expected/actual/provenance。

公开元数据以哈希、节点身份和属性摘要为主；完整来源文字只在需要时保留于本地诊断，不将原文全量复制到格式包。smoke 必须核对上下文与当前输入、输出哈希，不能直接信任历史 `passed` 字段。

## 4. N0：把验收反例变成正式失败断言

**改动位置：** `tests/`、`tests/fixtures/formatting/`、`docs/acceptance/`、`docs/formatting-progress.md`、`docs/formatting-capabilities.md`。

1. 从 `docs/acceptance/review_r0_r11_probes.py` 提取最小匿名夹具和正常对照，建立 `tests/test_post_review_contracts.py`。将观察布尔值改为正确行为断言，不继续使用“输出观察结果但 exit 0”的方式充当测试。
2. 字段不一致使用旧证据作离线解析反例，再通过 N8 的新 Word 运行验证修复；旧 PDF 不作为修复后的成功证据。
3. C01–C06 设置固定测试 ID；C07/C08 建立覆盖清单。不要用 skip、expectedFailure 或放宽断言把已知缺陷变绿。N0 阶段允许测试失败，失败列表必须与缺陷清单吻合。
4. 建立 `docs/acceptance/post-review/status.json`（新文件），记录 baseline、问题 ID、测试 ID、状态与证据路径。现有 R 批次历史结果保留，当前摘要链接最新验收结论，避免继续显示全批次已通过。

**退出条件：** 六类代码缺陷有独立失败断言；正常对照仍通过；复现不修改用户输入；当前失败与历史证据能一一对应。

## 5. N1：让分页字段独立于目录收敛

**改动位置：** `lib/field_updater.py`、`lib/delivery.py`、`lib/pagination.py`、`lib/pagination_types.py`、`lib/composition.py`；测试扩展 `test_r9_fields_and_notes.py`、`test_delivery_pipeline.py`、`test_publication_gates.py`。

### 实施步骤

1. 在 `FieldUpdater` 周围提取可复用 `FieldIndex`：统一遍历简单/复杂字段，记录跨 run 指令、受控缓存区间、依赖和稳定 ID。重复指令的字段用位置/实例区分。复用现有解析和循环诊断，不再另写一套不一致的正则扫描。
2. 扩展 `required_bookmarks()`：测量集合为部件边界、目录目标与所有受支持 PAGEREF 目标的并集。普通段落或表格内书签也要测量，不能只遍历 `toc_nodes()`。不支持的 story/开关在更新前明确诊断，不能留下旧缓存后当作已验证。
3. 移除 `has_toc=False → exact={} → all(empty) → break` 的收敛捷径。目录决定是否生成目录条目，不能决定是否处理分页依赖。
4. 定义每轮稳定状态：目标物理页、编号序列与打印数值、受控字段结果、目录条目标签、节边界与补白状态。记录每轮差异；存在依赖时必须有有效测量，缺目标不能默认页 1。
5. 更新 staging DOCX 后再导出、再测量。Word 若更新私有副本，要单独证明其受控字段与即将发布的 DOCX 相同；发布前从磁盘重新读取最终字段缓存，与测量结果逐一比对。
6. 页脚是否显示编号与 PAGEREF 的格式化结果分开建模：页脚无标签不等于引用为空。根据 PageSequence 的数值与字段开关计算引用；必要时扩展 PageRecord 的关联上下文，不直接把空 `expected_label` 当作 PAGEREF 结果。
7. 保留有界迭代和当前默认上限，通过内部参数注入上限测试；需要公开配置时再同步 Schema/CLI。来源正文只渲染一次，分页循环处理装配及字段，避免每轮重导入改变 ID。

### 必测与退出条件

| ID | 场景 | 必须断言 |
| --- | --- | --- |
| N1-T01 | 无目录，目标在第 4 页 | 最终 DOCX 缓存、PageRecord、PDF 对应引用均为 4 |
| N1-T02 | 同稿增加目录 | 引用随实际分页更新，不能沿用无目录页数 |
| N1-T03 | 罗马序列、无显示页脚 | 引用按目标编号规则正确，显示策略不吞掉引用 |
| N1-T04 | 普通段落/表格书签、重复 PAGEREF | 每个字段有明确目标和正确缓存 |
| N1-T05 | 缺目标、循环或持续分页振荡 | 结构化失败、逐轮差异、旧成果及旧元数据哈希不变 |
| N1-T06 | 最终保存后注入旧缓存 | 最终门禁拒绝，不能只验证导出副本 |

N1 离线通过后尽早运行 T01 的真实 Word 验证；C01 只有在新生成最终 DOCX/PDF 一致后才关闭。

## 6. N2：明确封面来源并验证生成内容

**改动位置：** `lib/config.py:get_template_path`、`lib/build_plan.py`、`lib/composition.py:build_cover`、`schemas/project-v3.schema.json`、`lib/content_integrity.py`；测试扩展 `test_composition.py`、`test_format_config.py`、`test_build_plan.py`。

### 配置决策

为 v3 解析出内部 `CoverSpec`。建议新增可选 `cover.mode`，并采用以下无歧义规则：

| 输入 | 解析结果 |
| --- | --- |
| v3 未指定模板/模式，或 `template: false` | generated；用配置字段生成通用封面 |
| v3 显式模板路径、未指定 mode | template；执行占位符绑定 |
| `mode: template` | 必须提供有效模板路径，验证配置字段对应的占位符 |
| `mode: static_template` | 显式使用固定封面；禁止同时设置会被忽略的动态封面字段 |
| v1/v2 | 保留当前自动发现兼容路径，并在 plan 说明实际来源 |

省略字段时按以上规则推导，不给 v3 偷偷回落到仓库模板。配置冲突在 plan 阶段报错。未启用封面部件的交付不应因无关模板报错。

### 实施步骤与测试

1. 在准备阶段解析并冻结 CoverSpec，记录模板 SHA-256 和配置 provenance。构建前复核模板未变化，避免 plan/render 重新发现不同模板。
2. 生成模式将 main_title（缺省使用 project_name）和实际非空字段转换为 `GeneratedContentExpectation`；模板模式扫描段落/表格及拆分 run 占位符，列出缺失字段，不把正文任意包含相同字符串当作封面通过。
3. 先实现局部封面门禁以关闭 C02，再由 N5 接入统一生成内容清单；不要等待所有映射重构才阻止旧封面发布。
4. 添加：仓库有旧模板而 v3 未指定；显式有效模板；拆分 run/表格占位符；缺占位符；固定模板与动态字段冲突；无封面交付；v1/v2 兼容。
5. 增加封面区间断言：预期标题必须出现在 cover 部件，正文或目录出现同名文字不能抵消封面缺失。Word 用例检查对应页面中的标题与作者/日期。

**退出条件：** C02 原夹具不靠手工补 `template:false` 即可正确生成论文封面；显式模板来源和内容可核验；旧版基线无未解释变化。

## 7. N3：贯通预期实例和装配后节点映射

**改动位置：** 新增 `lib/verification_contracts.py`；扩展 `lib/build_plan.py:PreparedBuild/RenderedBody`、`lib/engine.py:render_body`、`lib/package_importer.py:ImportResult`、`lib/composition.py:assemble_document`、相关来源渲染器。

1. 先固定第 3 节的数据类型及序列化版本，扩展已有 `ExpectedInventory`，避免两套预期模型长期并存。
2. `prepare_project_build()` 按选区、来源顺序、交付部件产生来源实例和来源有效属性快照。深拷贝或使用不可变值对象冻结验证相关字段，不能仅依赖外层 `frozen=True` 保护可变 config/dict。
3. 渲染返回 OutputNodeMap 的正文阶段结果；导入器扩展现有 ImportResult，将表格、样式、关系与脚注 ID 重映射关联到来源实例。文件名、样式前缀、正文字符串均不能单独作为身份。
4. 装配追加封面/目录及多部件时更新实际位置；对合法拆段/合并建立显式变换条目。不能映射的已声明能力应提前失败或明确未验证，不得跳过对应节点。
5. 保存 DOCX 后重新解析映射位置，核对节点类型、故事流与摘要。按每份交付物计算 missing/duplicate/unexpected 实例，形成上下文 sidecar，并绑定最终 DOCX 与输入哈希。

**必测：** 同文不同段、同源重复选区、两来源同路径/同标题、前插封面、表格合并单元格、脚注 ID 重排、删除节点、输出映射缺一项、修改 config 后与冻结预期不一致。

**退出条件：** 每个预期实例都有经过实际文件验证的定位；合法复用次数明确；改变输出样式名称不改变检查对象；映射缺失不能缩小覆盖分母。

## 8. N4：按实际策略检查目标格式和保护属性

**改动位置：** `lib/content_integrity.py:verify_delivery_format`、`lib/style_applier.py`、`lib/engine.py`、`smoke_test.py`、`lib/docx_inspector.py`；必要时提取公共属性策略解析器，避免渲染/验证分别解释 FormattingPolicy。

### 属性执行矩阵

| 情况 | 期望来源 | 验证方式 |
| --- | --- | --- |
| preserve | 来源有效属性及对象语义 | 与来源快照比较，不强行套目标 A4/正文样式 |
| restyle 托管属性 | resolved_format + 显式覆盖 | 验证最终有效值，包括继承和直接格式 |
| mixed | 按来源/选区/节点解析出的策略 | 保护区按来源，托管区按目标；重叠规则必须有确定优先级 |
| 行内强调 preserve | 来源逻辑文本区间的粗/斜体 | 允许合法强调；保护区强调丢失也必须失败 |
| 页面 source/target | 每节的来源或目标 PageSpec | 按节核验纸张、边距、页眉页脚距离，不能整篇使用单一页面期望 |
| 模板与生成部分 | CoverSpec/部件生成规则 | 与该部件的期望比较，不能因无来源映射而全部跳过 |

1. 将预期值构建与实际值求值分开，继续复用 EffectiveStyleEvaluator。目标样式里 `bold=false` 不应覆盖策略明确要求保留的局部粗体。
2. 删除“未知角色直接 continue 后整体通过”的行为：在预期集合中的节点必须找到并验证；改变样式 ID 不影响原始 role。非托管但要求保留的属性进入另一类必需检查。
3. 加入属性级 expected/actual/provenance 和覆盖数。沿用已声明的单位换算误差；不得为通过反例随意放宽字号、页面或页码容差。
4. engine 和 smoke 调用相同上下文及验证函数。结构检查模式明确报告未执行项目，不能升级为发布验收。

**必测：** C03 两个正例、保护属性被破坏、mixed 两区域、粗/斜体叠加、显式 false、第二个 run 错字号、Normal/未知样式 70 pt、5 个预期只定位 4 个、纯封面/目录交付、表格与页眉页脚保护。

**退出条件：** C03/C04 反例关闭；正确 preserve/强调在最终装配后与真实 Word 路径通过；相同输入下 engine/smoke 的必需检查结论一致。

## 9. N5：严格验证来源顺序、次数和生成内容

**改动位置：** `lib/content_integrity.py:build_expected_inventory/verify_content_integrity`、`lib/composition.py`、`lib/engine.py`、`smoke_test.py`；测试扩展 `test_content_integrity.py`、`test_source_regions.py`、`test_publication_gates.py`。

1. 从 N3 的有序实例清单构建每个内容部件的期望序列。使用 NodeRef/实例与语义摘要共同核验；多重集作为数量诊断，不再作为顺序正确性的替代。
2. 实际内容部件中，期望来源实例的出现次数必须精确匹配。不能简单把所有文本计数改为相等后误伤合法标题替换、字段结果及生成段落。
3. 将生成许可定义为“部件 + 类型 + 位置/关联来源 + 预期内容或结构 + 次数”，而非全局字符串白名单。同一句话在目录允许生成，不意味着正文可以多复制一次。
4. 将分页/编号字段的受控缓存从普通文字比对中分离，交给 N1 检查；保留字段指令、依赖及非缓存文字的完整性核验。空段、分节承载段和重复表格单元格也要有明确结构规则，不全局忽略。
5. 表格按拓扑与单元格顺序，媒体/公式按实例，脚注/尾注按定义和引用，超链接按目标验证。对合法 ID 重映射核验前后语义，不能只检查对象 XML 存在。

**必测：** `[A,B]→[B,A]`、`[A,B]→[A,B,A]` 分别失败；两来源本来均有 A 则 `[A,A]` 正确通过；声明复用同一区域两次正确通过；生成目录重复文字合法而正文多复制失败；字段缓存合法更新通过；图片/公式少一个实例失败；两来源注释冲突合并正确。

**退出条件：** C05 关闭；C02 的生成内容检查纳入统一交付门禁；每个失败报告可定位到部件与来源实例。

## 10. N6：保护迁移源文件与输出事务

**改动位置：** `lib/manifest_migration.py:migrate_manifest_file`、`synthesize.py` 对应 CLI 错误处理、`tests/test_r7_workflows.py`、`tests/test_r11_release.py`。

1. resolve 源/目标后比较规范路径；目标存在时再用 `samefile()` 检查硬链接身份。符号链接、相对路径别名必须纳入检测。
2. 源与目标为同一文件时，无论 replace 是否为 true，均返回 `MIGRATION_SOURCE_EQUALS_OUTPUT` 等明确诊断。独立目标的 replace 语义保留。
3. 完成解析、迁移及 ProjectConfig 校验后再写入。使用同目录临时文件与原子替换；未允许覆盖时保留排他创建语义，失败不得留下半份目标或覆盖原目标。
4. 测试迁移到另一目录后的引用重定位，确保安全修复没有破坏已有 provenance 与相对路径行为。

**必测：** 同路径、`..` 别名、符号链接、硬链接、独立已存在目标无 replace/有 replace、无效源 JSON、迁移校验失败、写入失败；所有拒绝路径检查源 SHA-256 不变。

**退出条件：** C06 关闭，CLI 与函数层都保护源文件，正常独立输出示例通过。

## 11. N7：提高无结构样例的候选识别与评测可信度

**改动位置：** `lib/format_analysis.py`、`lib/role_mapper.py`、`lib/format_review.py`、`lib/contracts.py`、必要的 analysis/decisions Schema；新增共享候选推导模块 `lib/role_candidates.py`；扩展 `docs/acceptance/r7-evaluation/` 与 `tests/test_format_analysis.py`、`test_format_review.py`、`test_r7_workflows.py`。

### 11.1 推导规则与用户流程

1. 保留显式 RoleMap 的最高优先级；大纲/标题样式作为结构证据。无结构情况下综合有效字号相对正文分布、字体、段前后距、缩进、粗体比例、文本长度、编号模式及相邻段落结构形成候选。不能用最大字号、首段居中或粗体单信号确定标题。
2. 先识别“像标题的段落/簇”，再推导层级；层级不足时返回候选列表和缺失层级，不凭空补齐 1–9。列表编号、题注、短正文必须作为竞争角色考虑。
3. 每个候选保留信号、反证、样本数、影响节点及分数。启发式分数不直接称为统计概率；自动接受阈值必须根据验证集校准。
4. 样例格式候选与目标稿件角色映射共用候选能力但保留两个决策步骤。学习到的样例字号不应直接决定另一稿件的语义角色；显式人工映射仍优先。
5. HTML 展示“当前候选/其他候选/继承预设/人工指定”，支持批量校正簇及例外节点；继续绑定 report_id/source_sha256，未决项不得被编译暗中接受。
6. 分析导出显式 false/0 属性，避免把“未设置”与“设置为否/零”混淆。新增字段同步 Schema、兼容读取和测试；跨版本不兼容时给出明确版本诊断。

### 11.2 评测设计

保留原 20 份结构化合成样本作为回归；去大纲副本作为同源对照，不能将原件与副本算作 40 份独立样本。新增按不同文档结构设计的开发集和冻结验收集，按文档家族分割，同一模板及其变体不得跨集合泄漏。

冻结验收集至少覆盖公文、报告、论文、手工格式各 5 份，共至少 100 个标题正例、500 个正文负例；其中必须包含全部 Normal 且无 outlineLvl 的文档。标注通过人工核对，不能根据待测识别规则生成答案。已有用户文档仅在明确授权后使用；否则采用独立设计的匿名材料并标明局限。

分别输出结构化/无结构、各类别、各层级的 precision、recall、自动接受覆盖率、正文误报率、候选覆盖率、校正操作数和未决数。自动接受 precision 的分母是实际自动接受预测，不能用真实标签数量代替，也不能漏算把正文预测为标题的误报。

### 11.3 本轮工程门槛

以下为拟实施门槛，不是已测结果或行业标准：

| 能力 | 验收要求 | 未达时处理 |
| --- | --- | --- |
| 结构化识别 | 原有独立集无回退，按层级报告结果 | 修复回归，不用整体平均掩盖某层失败 |
| 无结构辅助识别 | 冻结集标题候选覆盖率 ≥90%；候选最多 3 个角色，另报候选节点精确率防止全篇撒网 | 保持 partial，改进候选规则与样本 |
| 标题自动接受 | 实际自动接受 precision ≥98%，同时覆盖率 ≥50%；分组公开样本数与置信区间 | 保留辅助校正模式，自动识别支持状态不关闭 |
| 人工校正闭环 | 选定困难样例的最终角色/属性 100% 等于人工基准；应用另一目标得到预期有效格式 | 修复 UI、Schema、编译或目标映射接线 |
| 数据诚实性 | 无预测 precision 为 null；对照不冒充独立样本；冻结验收集不用于调阈值 | 重新划分数据并重跑 |

这些门槛约束具体冻结集，不能外推为任意真实论文的准确率保证。独立报告中列出样本局限、校正成本及自动接受范围。

**退出条件：** C07 的评测覆盖缺口已关闭；独立人工基线、分组指标、手工校正 smoke 与自动接受门槛均已形成可复现证据。自动识别结论仅对当前匿名冻结集有效，扩大真实文档范围时必须新增独立冻结集并重新校准。

## 12. N8：用真实 CLI 与最终成果关闭集成矩阵

**改动位置：** `tests/test_word_formatting.py`、`test_m1_acceptance.py`、`test_thesis_acceptance.py`，新增 `tests/test_post_review_word.py`、匿名夹具生成器及预期清单。

每条用例从生产 CLI 进入，先 plan/doctor，保存当次导出的原始 PDF。Word 顺序执行。Mock 可用于失败注入和快速单元测试，不能替代以下真实通过证据。

| ID | 夹具 | 核心断言 |
| --- | --- | --- |
| NW01 | 无目录多页正文 + PAGEREF | 最终缓存、目标实测编号、PDF 引用一致且不为默认 1 |
| NW02 | 同来源有目录/独立目录变体 | 目录标签、内部链接及引用文档的最终页码一致 |
| NW03 | v3 默认封面，仓库存在旧模板 | 配置标题、作者、日期在 cover 部件；无旧业务封面吸入 |
| NW04 | preserve、180×240 mm、已有页脚与局部强调 | 来源有效属性和 PDF 纸张保留；检查正常通过 |
| NW05 | mixed，托管正文与保护附件 | 托管区改变、保护区不变；按节和节点验证 |
| NW06 | 独立声明、中英摘要、目录、正文、参考文献、附录 | 使用真实 source.regions/部件；来源无缺失/重复，摘要语言与配置对应 |
| NW07 | 罗马前置序列、正文重启、右页起章 | 每个部件边界、每个目录目标、空白页原因及编号全部可解释 |
| NW08 | 横版表格节 + 不同首页/奇偶页眉 | 各节纸张与方向、页眉实际文字及适用页面符合配置 |
| NW09 | 两来源同 noteId，注释含链接/图像 | 注释定义/引用/关系完整，内容不串来源 |
| NW10 | 图片、表格、列表、OMML、SEQ/REF/PAGEREF | 对象实例、顺序、字段指令与最终值均正确 |
| NW11 | 样例→校正→包→另一目标→Word | 包不含样例正文，目标映射生效，最终属性匹配已确认规则 |
| NW12 | 最终字段/格式/内容故障注入与发布失败 | 必需门禁拒绝；旧成果/元数据哈希不变；诊断可定位 |

NW06 使用真正英文摘要，不能仅把中文段落标题改为 Abstract。NW08 若暴露配置无法表达每节页面或页眉，先补齐相应 PageSpec/PartSpec/Schema 和实施路径，再运行；不得通过手工修改最终 DOCX 完成夹具。无法支持时明确阻断该能力，不能把未覆盖项记为通过。

每份成功成果至少断言：结构及来源清单、托管/保护属性覆盖、每节页面、实际页码标签、生成内容、对象关系、最终受控字段、发布事务与输入哈希。逐页查看所有关键成果的 PDF，记录封面、目录、节边界、图表/注释页；视觉查看补充裁切/重叠判断，不替代机器断言。

**退出条件：** C08 原矩阵缺口逐项关闭；“完整论文通过”链接到具体 NW 用例和能力边界，不能只靠某份大样例的 exit 0 与 XML 标签存在。

## 13. N9：定位 PDF 对象偏移警告

**改动位置：** `lib/qa.py`、`lib/pagination.py` 和现有 Word 导出辅助路径；新增 `docs/acceptance/post-review/pdf-diagnostics/` 证据。

1. 固定现有两份原始警告 PDF 的哈希，不改写历史证据。以同一匿名 DOCX 在相同 Word 环境重新导出，记录导出返回时间、文件关闭/稳定状态、输入 DOCX 哈希及原始 PDF 哈希。
2. 对同一个 PDF 运行 pypdf、Poppler 及项目当前页码提取路径，比较页数、每页 MediaBox/CropBox、文字/标签和书签目标。收集原始 stderr 与工具版本，不只截图。
3. 区分导出未完成、文件被追加写入、Quartz 生成结构、解析器兼容等可能来源；以对照实验确认，不能凭警告措辞推断根因。未证实根因前不静默过滤日志。
4. 若为导出时序，修正等待完成机制并增加重复导出一致性测试；若为已确认且不影响所需语义的兼容警告，形成有适用条件的诊断规则和跨解析器证据。任何页数、尺寸、字段标签差异均阻断相关交付。
5. 不自动重写原始 PDF 以消除警告。若未来引入规范化副本，应单独标识原件/派生件及哈希，不把副本冒充当次原始 Word 输出。

**退出条件：** 根因、影响范围和处理路径有可重复证据；本轮已知 Quartz orphan-xref 模式满足该条件，采用非破坏性诊断分类；若出现被引用的缺失对象、页数/尺寸/字段差异或其他未解释警告，相关 PDF 兼容能力仍为 blocked，不能宣称全部解析/视觉 QA 通过。

## 14. N10：收尾、证据与发布状态

**改动位置：** `README.md`、`DEVELOPMENT.md`、`docs/formatting-progress.md`、`docs/formatting-capabilities.md`、历史计划的当前状态摘要、`examples/custom-format-demo/`、`tests/test_build_metadata.py`、`tests/test_r11_release.py`。

1. 更新实际可复现 CLI 示例，增加默认通用封面、显式模板、preserve/mixed、无目录 PAGEREF 及学习校正流程。示例使用新的匿名输出目录。
2. 干净虚拟环境安装 `requirements.txt` 并运行离线套件及 plan；本机 Word 环境单独记录依赖、字体和 Word 版本。不要把环境缺失计为功能通过。
3. 元数据保存上下文版本与哈希、逐交付物检查、预期/实测覆盖、每轮字段收敛摘要、最终 DOCX/PDF 哈希、PDF 警告分类。确保 QA 后未再次修改将发布的 DOCX；若修改必须重验。
4. 同步能力表和当前状态。历史 R 批次记录保留，通过新的 N 批次证据关闭 C 缺陷。实现完成但尚未 Word 验证，状态仍为 implemented/partial，不能直接 verified。
5. 保存正式证据后清理本轮 `.work`、Office 锁文件、渲染缓存及 Python 缓存；只清理本轮临时目录，不动用户输入和历史证据。

### 验证命令

以下新测试文件需在对应批次建立后运行；不存在不是可跳过的验收项。

```bash
python3 -m unittest discover -s tests -p 'test_post_review_contracts.py' -v
python3 -m unittest discover -s tests -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p 'test_post_review_word.py' -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p 'test_word_formatting.py' -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p 'test_m1_acceptance.py' -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p 'test_thesis_acceptance.py' -v
git diff --check
```

既有 opt-in 分页测试若使用不同环境开关，应读取测试代码补充正确命令，不能以设置一个变量推断所有 Word 用例都已执行。测试报告同时统计运行数、失败、错误及 skip；发布必需用例不得 skip。

### 每批证据格式

`docs/acceptance/post-review/nX-status.json` 至少包括：

```json
{
  "batch": "N1",
  "status": "todo",
  "closes": ["C01"],
  "code_sha256": {},
  "fixture_sha256": {},
  "commands": [],
  "tests": {"run": 0, "failed": 0, "errors": 0, "skipped": 0},
  "word_cases": [],
  "artifact_sha256": {},
  "remaining_limits": []
}
```

`status` 使用 todo/in_progress/blocked/verified；示例初值不得用作完成证据。每条 Word case 附实际断言和环境，不能只有文件名。已确认缺陷关闭须附原失败→新通过的对应测试，最终成果必须来自受检代码。

## 15. 最终验收清单

- [x] C01：有/无目录分页字段均在最终 DOCX 与 PDF 一致；失败不发布。
- [x] C02：v3 默认封面来源正确，必需字段在正确部件；旧版兼容已验证。
- [x] C03：preserve/mixed/行内强调按策略通过，保护属性被破坏会失败。
- [x] C04：预期覆盖分母固定，未知样式和丢失映射不能漏检。
- [x] C05：来源顺序及实例次数精确，生成内容许可按部件与位置限制。
- [x] C06：迁移始终保护源文件和别名，输出失败无半成品覆盖。
- [x] C07：独立评测、分组指标和困难样例校正闭环齐备；自动能力按门槛标注。
- [x] C08：NW01–NW12 必需范围有真实生产路径证据，未实现能力明确阻断。
- [x] PDF 警告已解释或影响能力保持 blocked，原始证据未被覆盖。
- [x] 既有回归和新负向回归通过；必需 Word 用例无 skip；示例/元数据/文档一致。

C07 已勾选：独立匿名无大纲冻结集包含 20 份文档、100 个标题正例和 500 个正文负例，标签绑定独立人工可见段落基线；候选覆盖率 100%、候选块 precision 1.0、自动接受覆盖率 100%、自动接受 precision 1.0、正文误报 0，达到 98% precision/50% coverage 门槛。结论限定于该冻结集，不外推为任意真实文档。PDF 项按计划的允许路径关闭：本轮已知 Quartz orphan-xref 警告已解释并形成非破坏性分类规则，历史原始证据未被覆盖；其他未解释结构异常仍保持 `blocked`。

满足里程碑 A 后可先交付可靠的格式执行能力；里程碑 B 的自动识别若未达到门槛，应明确提供辅助校正模式。只有里程碑 C 的矩阵与证据齐备，才关闭本轮完整论文验收。不得仅修改状态文件、提高测试计数或缩小检查覆盖来完成计划。
