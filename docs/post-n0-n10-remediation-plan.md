# N0–N10 复验后的补全执行计划

日期：2026-09-06；截至 2026-09-08 状态：**已完成（声明范围内）**。本轮批次编号为 **S0–S8**，缺陷沿用 **D01–D06**，不覆盖历史 R/N 批次记录。当前逐批证据索引见 [`docs/acceptance/remediation/status-summary.json`](acceptance/remediation/status-summary.json)；能力结论仍限定在声明的匿名夹具、冻结集和 Word 环境。

依据：[N0–N10 独立复验报告](n0-n10-acceptance-report.md)、[复验反例及结果](acceptance/n0-n10-review/probe-observations.json)、当前生产代码。下文新增的类型、函数、测试和证据文件均为实施设计，不能当作已经完成的能力。

## 1. 本轮目标与执行原则

本轮将工作集中在六个已经确认的问题：来源格式保护误判、页眉页脚漏检、额外图片实例漏检、标题层级推断错误、自动接受指标口径错误、真实 Word 场景覆盖不足。

历史基线是离线 286 项通过（3 项条件跳过），真实 Word NW01–NW12 为 12/12 通过；这些数量不能代替退出条件。执行后的最新索引为离线 324 项通过（6 项条件跳过）、NW01–NW12 为 12/12、remediation_word 为 11/11；这些数量仍不能代替逐批退出条件。已确认修复的默认封面、无目录 PAGEREF、迁移源文件保护继续作为回归基线，不重新设计相应功能。

执行必须满足以下原则：

1. 每项保护契约同时有“来源保持正确应通过”和“只破坏该属性应失败”的成对测试。单独通过合法样例或单独拒绝某个坏文件都不足以关闭问题。
2. 预期在来源检查/准备阶段产生；不得根据输出已有内容、样式、story 或对象数量缩小检查范围。
3. 把“识别为标题”和“识别为正确层级”分开度量；先修评估器，再调整候选算法。
4. 每个集成场景同时登记配置前提、夹具事实、生产调用和最终断言。测试名包含某种能力，不代表实际覆盖了该能力。
5. Word 按 plan → doctor → build 顺序执行，测试串行。沿用用户对匿名验收目录的授权和固定访问目录 `/private/tmp/document-synthesis-word-access`。
6. 保留原有报告及失败证据。本轮只增加新证据和更新当前状态摘要，不回写历史测试结果。

## 2. 批次、依赖与交付物

| 批次 | 核心改动 | 缺陷 | 依赖 | 主要交付物 |
| --- | --- | --- | --- | --- |
| S0 | 固定成对反例和场景清单 | D01–D06 | 无 | 正式失败测试、场景契约、基线哈希 |
| S1 | 全部受保护字符属性按逻辑区间验证 | D01 | S0 | 区间比较器、完整属性覆盖、真实 preserve 成果 |
| S2 | 页眉页脚 story、节绑定及距离核验 | D02 | S1 | story 清单、节期望、页级使用证据 |
| S3 | 非文字对象精确实例与引用位置核验 | D03 | S0；story 集成依赖 S2 | 对象实例契约、生成对象许可、负向回归 |
| S4 | 修正精确角色评估指标和数据约束 | D05 | S0 | 指标 v2、评估器手算对照、数据分割记录 |
| S5 | 编号层级、冲突与弃权策略 | D04 | S4 | 层级候选证据、困难样例评估、校正闭环 |
| S6 | 补全论文/页眉/注释的真实夹具与断言 | D06 | S1/S2/S3；识别用例依赖 S5 | 独立前置部件、奇偶页眉、注释对象矩阵 |
| S7 | 验证失败构建与发布事务，固定完整证据 | D06 | S1/S2/S3/S6 | 实际第二次失败构建、事务故障报告、成果证据 |
| S8 | 全套回归与发布范围同步 | 全部 | S0–S7 | 最终验收索引、示例、能力与状态更新 |

建议执行顺序：S0 → S1 → S2 → S3 → S4 → S5 → S6 → S7 → S8。S4 的独立指标修复可以提前，但 S5 不得先于 S4 调参。

每个批次以“失败断言 → 实现 → 对照回归 → 证据”形成可审阅的改动。不设以测试数量为目标的验收，也不以固定日期替代质量门槛。

## 3. S0：固化本轮反例与验收清单

**当前进度（2026-09-08）**：D01–D06 的历史反例观测已与本轮正常/损坏成对测试逐项对照，并写入 `acceptance/remediation/scenario-contracts.json`；S0 sidecar 已标为 `verified`。这不改变历史报告，只补充当前批次的可追溯证据。

**文件位置**：新增 `tests/test_remediation_contracts.py`、`tests/test_role_evaluation.py`；复用 `docs/acceptance/review_n0_n10_probes.py` 的匿名输入；新增 `docs/acceptance/remediation/`。

### 执行动作

1. 将现有探针转换成正式断言。探针中的 `passed=true/false` 观察不能继续以脚本 exit 0 充当通过。
2. 拆开 D02 的合并损坏例，分别修改页眉文字、页脚文字、字体、header_distance、footer_distance。这样每个检测缺口都有独立失败证据。
3. 将 D03 拆成原件、额外图片、删除图片、移动图片四种输入，并保留合法关系和不同 docPr ID，确保测试失败来自契约而非 DOCX 损坏。
4. 固定 D04 的缺层级文档与人工期望：`1. Introduction → heading.1`，`1.1.1. Detailed protocol → heading.3`；分别检查预测、是否自动接受、最终校正结果。
5. 用不调用识别器的固定预测数据测试 D05：两个自动接受标题一对一错时，精确层级 precision 必须为 0.5，不能仍为 1.0。
6. 建立机器可读 `scenario-contracts.json`，每条包含 `id/defect/preconditions/fixture_assertions/command/artifact_assertions/required/evidence`。初始化为未运行，禁止自动把存在同名测试转换为 verified。

### 退出条件

原缺陷对应测试在修复前确实失败，正常对照能够通过；失败原因与报告一致。S0 允许缺陷测试处于失败状态，但不能使用 expectedFailure/skip 掩盖后续批次的退出条件。

## 4. S1：按逻辑文本区间验证全部保护属性

**现状（2026-09-08）**：`BlockInspection.effective_run_spans` 与最终门禁已统一使用逻辑字符区间，中文字体、西文字体、字号、颜色、粗体和斜体均逐区间核验，并记录 expected/actual/provenance。S1-T01–T07 及真实 Word SW01 正常/损坏成对场景已通过；其余状态以 [`s1-status.json`](acceptance/remediation/s1-status.json) 为准。

**文件位置**：`lib/docx_inspector.py`、`lib/content_integrity.py:_logical_inline_segments/verify_delivery_format`、`lib/verification_contracts.py:_serialize_run_spans`、`lib/style_applier.py`；相关格式/保护测试。

### 实施方案

1. 将逻辑区间坐标与 NodeRef 的内容身份哈希分开。定义版本化 `TextSpan`：`start/end/effective_properties/origin`，偏移按确定的 Unicode 字符序列计算；空格、制表、换行须有明确表示，不能比较时临时删空白后重新拼出区间。
2. 普通文字、字段指令/受控缓存、无文字对象分别建模。分页字段缓存允许变化，不应推动其后普通文字保护区间错位；用字段实例身份进行对齐。零文字 drawing/公式由 S3 的对象锚点负责。
3. 将现有 `_logical_inline_segments()` 提取或泛化为共享比较器，例如 `compare_protected_spans(expected, actual, attributes)`。合并来源和实际区间的全部端点，在每个非空子区间比较对应有效属性；run 拆分/合并不改变结果。
4. preserve 和 mixed 保护区的中文字体、西文字体、字号、颜色、粗体、斜体使用各自来源有效值。restyle 托管区按目标；只有策略允许保留的属性覆盖目标期望。不要把新增来源属性检查无条件套到全部 restyle 内容。
5. 有效属性通过现有 `EffectiveStyleEvaluator` 求值，区分缺失、继承与显式 false/0。所需属性无法求值时记录未验证并拒绝相关必需检查，不能退回首 run 期望。
6. 统一报告节点覆盖和属性区间覆盖：输出 `expected_intervals/verified_intervals/missing_intervals` 及 expected/actual/provenance。正文节点定位 1/1 不代表其所有字符属性已核验。
7. 若区间序列化发生不兼容变化，提升验证上下文契约版本；旧 sidecar 可重新从可信来源构建，不能静默将旧区间当作新坐标。相应更新元数据读取与 smoke 验证。

### 成对测试

| ID | 应通过 | 只改变该项后应失败 |
| --- | --- | --- |
| S1-T01 | 同段 12/24 pt 保留 | 第二段文字改为 12 pt |
| S1-T02 | 同段不同中英字体/颜色 | 改坏第二个区间的字体或颜色 |
| S1-T03 | 相同有效格式 run 任意拆分 | 拆分后的某个子区间字号错误 |
| S1-T04 | 合并同格式相邻 run | 跨不同属性区间强制统一格式 |
| S1-T05 | 保留空格/tab/换行两侧不同属性 | 移动一个区间边界使属性作用到错误文字 |
| S1-T06 | 受控字段缓存合法改变 | 字段后的普通文字格式被改坏 |
| S1-T07 | mixed 托管区改写、保护区保留 | 保护区发生目标样式覆盖 |

**真实验收**：用上轮 `mixed_sizes` 匿名源重新走 CLI；最终 DOCX 与 PDF 保留 12/24 pt，构建成功。再将对应损坏例送入真实最终门禁，必须失败。不能仅修改原失败测试的预期来使其通过。

## 5. S2：将页眉页脚与节几何纳入预期清单

**文件位置**：`lib/docx_inspector.py:inspect_docx/NodeRef/SectionInspection`、`lib/verification_contracts.py`、`lib/content_integrity.py`、`lib/composition.py`、`lib/pagination.py`、`lib/package_importer.py`，以及角色分析调用点。

### 5.1 数据模型

扩展现有 DocumentInspection，建议增加独立 `stories` 与 `section_bindings`，避免直接把页眉段落塞进正文 blocks 后误入正文顺序、目录或样例正文聚类。

| 类型（拟新增/扩展） | 内容 |
| --- | --- |
| `StoryInspection` | 源文件哈希、实际 part_uri、story 类型、段落/表格、字段与对象实例 |
| `SectionStoryBinding` | section_id、header/footer、default/first/even、实际引用/继承来源、适用开关 |
| `ExpectedSection` | 来源/目标身份、纸张/方向/四边距、header_distance/footer_distance、预期 story 绑定 |
| story 输出映射 | 来源 story 实例 → 输出 part_uri 与节绑定；允许包导入改名但不改变语义 |

NodeRef 不再为所有节点硬编码 `word/document.xml`。两个不同 part 中相同 paragraph path 必须能区分。直接通过 OOXML 关系读取 story，避免为了检查而访问会创建空页眉页脚的可写 API；检查前后源哈希应一致。

共享同一 header/footer part 时，内容按唯一对象核验一次，节绑定和适用页面分别核验。被保留但当前不显示的 first/even 内容仍属于来源保护对象；只对实际启用的变体要求 PDF 可见文字证据。

### 5.2 执行动作

1. 解析各节引用及继承链、different_first_page 与奇偶页开关；缺失关系、循环/无法解析绑定应有明确诊断。
2. 将 story 的来源文字、字段、表格、有效样式和对象加入 ExpectedDelivery；通用节点定位器支持不同 part，S1 比较器复用于 story 文本属性。
3. 页面检查覆盖每个预期节的纸张、边距与两个距离，避免现有仅从六项 target_geometry 中取字段的做法。页面方向和宽高保持一致性检查。
4. 根据来源选区与生成部件建立输出节对应关系，不能仅按“全部源节按下标对输出节”比较；增加封面/目录后仍需找到正确来源节。
5. Word 页级证据记录实际节、编号序列、所用变体及该页观测文字。首页/奇偶页选择规则与 Word 实测对照，不能仅依据 PDF 物理页奇偶做未经验证的推断。
6. 更新 format_analysis/RoleMapper 等读取检查结果的代码，明确仅正文或显式支持的 story 参与正文角色识别；页眉字体不能污染正文/标题候选统计。

### 必测与退出条件

- 普通、不同首页、奇偶页三种变体分别放入不同匿名文字，并确保实际文档足够长，使每种启用变体出现。
- 两节解除链接与两节共享链接各有对照；修改某一节不应串改其他节。
- 分别破坏 story 内容、字号、距离、关系目标及启用开关，所有必需检查应失败。
- 验证横版节与前插生成封面后，来源节和 story 仍正确映射。
- 真实 PDF 中核对各页独有文字与配置绑定；不能只断言 `section.header` 含某字符串。

**D02 关闭条件**：原五类损坏反例均被检出，正确源内容通过；覆盖报告含必需 story 与节属性，不只显示正文 1/1。

## 6. S3：精确核对非文字对象实例

**文件位置**：`lib/content_integrity.py:SemanticInventory/extract_semantic_inventory/verify_content_integrity`、`lib/verification_contracts.py`、`lib/package_importer.py:ImportResult`、`lib/notes_merger.py`、`lib/composition.py`。

### 实施方案

1. 在资源哈希之外记录 `ObjectOccurrence`：所属交付部件、story、宿主来源实例、对象类型、局部序号/锚点、语义指纹和关系目标。文件中一个图片资源被引用两次是两个实例。
2. 预期实例由源选区和显式复用次数产生。实际实例从最终文件提取，不能从输出倒推允许次数。
3. 按部件/宿主锚点核对位置和精确次数，报告 missing/extra/moved/replaced。相同哈希只用于识别资源，不能替代出现位置与次数。
4. 为生成封面 logo、目录元素及合法其他生成对象建立有限许可：部件、生成规则、位置/锚点、类型与次数。不得采用全局“这个哈希出现过，因此任意额外实例都允许”的白名单。
5. 同步补查公式、表格、脚注/尾注引用等当前使用下限判断的路径。注释定义和引用次数分开：合法重复引用同一定义，不应被强制变成多个定义；导入 ID 重映射必须保持引用→定义→图片/链接对应关系。
6. 区分共享页眉的文件内对象实例与页面上重复渲染次数。保护页眉的一张图可在多页显示，不能按 PDF 出现次数误报重复。

### 必测矩阵

| 情形 | 结果 |
| --- | --- |
| 原 1 图 → 同段 2 图、关系和 docPr 均合法 | 拒绝，extra=1 |
| 原 2 个相同资源引用 → 1 个 | 拒绝，missing=1 |
| 两来源各有 1 个相同图，按顺序合并 | 通过，预期实例为 2 |
| 声明同一选区复用两次 | 通过；第三次额外实例拒绝 |
| 图数量相同但移动到另一来源段落 | 拒绝，除非存在明确许可的布局变换 |
| 明确生成封面图 + 正文来源图 | 各按所在部件核验；不能互相抵消缺失 |
| OMML/表格/注释引用额外或缺失 | 精确检查对应实例和关系，给可定位诊断 |

**退出条件**：D03 原反例由 true 转为明确失败；来源对象与生成对象的正常对照全部通过，并在包含图片的真实 Word 构建中确认不误伤正常导入。

## 7. S4：先修评估器，明确自动角色指标

**文件位置**：`docs/acceptance/r7-evaluation/evaluate_unstructured_dataset.py`、相关结构化评估器及状态写入脚本；新增 `tests/test_role_evaluation.py`。

### 7.1 固定统计口径

对具有完整人工标签的节点，设 H 为真实标题节点集合，A 为实际被自动接受为标题角色的预测集合（包括被误接受的正文），C 为 A 中预测角色/层级与标签完全相同的集合。

| 指标 v2 | 定义 | 用途 |
| --- | --- | --- |
| `heading_detection_precision` | A 中真实标题数 / count(A) | 仅评价标题/正文二分类 |
| `exact_role_auto_precision` | count(C) / count(A) | 自动 RoleMap 放行指标 |
| `auto_acceptance_coverage` | A 中真实标题数 / count(H) | 自动接受覆盖，单独暴露错误层级 |
| `correct_auto_coverage` | count(C) / count(H) | 正确自动完成比例 |
| `candidate_exact_role_coverage` | 真实角色出现在候选中的标题数 / count(H) | 不能用“任意标题候选存在”代替 |

文档主标题 title 与 heading.1…9 按标签定义分别统计。所有自动接受节点都必须有标签或明确评测状态；未标注节点不能静默从预测分母中剔除。无自动接受时 precision=null；空集合及无标题文档不以 1.0 填补。

### 7.2 实施与测试

1. 先用固定 labels/predictions/accepted 集合实现纯统计函数，不调用候选算法。修复 `auto_accept_correct += 1` 无条件计入错误层级的问题。
2. 升级 metrics_schema_version，停止将旧 `automatic_acceptance_precision` 直接作为精确角色门槛；消费旧指标的文档/状态脚本同时更新。旧报告保留并标注口径，不覆盖为新结果。
3. 手算回归至少包括：两标题一对一错（exact=0.5）；正文误报；全部弃权；缺层级；主标题与一级标题混淆；没有标题的文档。
4. 用上轮 D04 的固定预测重评，必须出现二分类 precision=1.0、精确角色 precision=0.5，而不是两个都等于 1.0。
5. 人工校正操作数只有实际动作记录才可填实数；仅根据标签计算的次数继续叫 estimate。存在 baseline 不能自动将真实人工操作数写成 0。

**退出条件**：统计函数与手算一致；新旧口径明确；D05 关闭之后才允许 S5 使用新指标调阈值。

## 8. S5：处理编号深度、层级缺失与证据冲突

**当前进度（2026-09-08）**：候选证据、四类 DOCX 困难夹具和跨目标显式决策闭环已补齐；HTML 审阅页已补充绑定 `report_id/source_sha256` 的 `review_audit` 操作轨迹及人工确认开关，并严格限制只有 `origin=browser_ui` 且 `automation=false` 才能声明人工确认。真实审阅者已完成匿名困难样例审阅并导出 2 个操作的 `reviewer_attested=true` 决策；该决策已编译格式包、目标 RoleMap，并在另一份目标稿件上完成真实 Word build/smoke。S5 在声明范围内标为 `verified`；20 份冻结集的 aggregate `manual_correction_count` 仍为 null，不外推为冻结集人工一致性通过。

**文件位置**：`lib/role_candidates.py`、`lib/format_analysis.py`、`lib/role_mapper.py`、`lib/format_review.py`、必要的 analysis/decisions Schema 与 `lib/contracts.py`。

### 实施步骤

1. 将目前编号布尔值扩展为编号证据：编号形式、分量、候选深度、可信度及反证。阿拉伯多级编号的分量可提供层级线索；日期、小数、版本号、列表及题注必须经过排除或降低置信度。
2. 将“是否标题”与“标题层级”分成两个决策结果。字号相对大小只作为格式证据，不再把不同字号的排序直接映射到连续 heading.1…n。
3. 明确优先级：人工映射最高；有效结构化证据优先；一致的编号与上下文作为候选；格式证据辅助。证据冲突时保留候选并标为待确认，不能为了自动覆盖率将分数固定抬到 0.985。
4. 缺失中间层级时不压缩层级、不插入不存在的标题。D04 中应给出 heading.3 的候选；若无法可靠自动判定，应弃权并让人工确认，禁止高分自动选 heading.2。
5. 单个样式只出现一次不再被描述为“重复样式证据”。若使用重复频率或校准分数，记录真实支持数、开发集和校准方式。
6. 在混合结构样例中逐节点使用证据；不能因全篇某一段存在 Heading 样式就禁用其他 Normal 段的候选推导。此项作为相关回归，不预先宣称它已被本轮复验单独证明为缺陷。
7. HTML 展示候选层级、编号解释与冲突原因，批量选择与单点例外均能导出；明确区分自动接受、人工确认和未决。未决项不得被编译静默接受。

### 数据与放行门槛

保留旧固定五层样例作为回归集。另建立按文档家族隔离的开发/校准/冻结验收集；同一模板及其去大纲或改字号变体不能跨分区。冻结集至少覆盖缺层级、同层不同字号、不同层同字号、正文编号、短正文粗体、题注、部分结构化和缺标题文档。

继续沿用此前规模与目标：冻结验收至少四类各 5 份、100 个标题正例、500 个正文负例；新增困难样例必须实际占有样本，不只是写在 limitations。标注与审阅记录绑定源哈希、标注版本与修改记录；脚本自动输出 `manually_reviewed` 不构成人工审阅证据。匿名合成样本的结论限定在该集合。

自动 RoleMap 的门槛使用 S4 新口径：**exact_role_auto_precision ≥98%，auto_acceptance_coverage ≥50%**，同时报告 correct_auto_coverage、各层级结果、样本数与区间。阈值仅在校准集确定，冻结集不能用于反复调参。

如果安全弃权关闭了 D04 的高分错判，但自动覆盖尚未达标，可先发布辅助校正模式；S5 自动能力保持 partial，不得据此把整轮自动识别验收记为完成。

**退出条件**：D04 不再错误自动接受；精确指标经 S4 验证；困难样例能通过“候选→人工决策→格式包/目标映射→另一文档”的实际流程。自动支持声明是否放行由上述门槛决定。

**本批次结果**：上述真人审阅、格式包编译、目标 RoleMap、真实 Word build 与 smoke 已完成；证据和哈希见 [`s5-human-review-evidence.json`](acceptance/remediation/s5-human-review-evidence.json)。后续若扩大模板、冻结集或 Word 环境范围，应新建对应复验，不得把本批次结果外推。

## 9. S6：补齐真正的论文、页眉和注释场景

**当前进度（2026-09-08）**：NW01–NW12 与 remediation_word 的真实 Word 夹具均已通过，包含完整论文部件、横向节故事、多来源注释、对象/封面许可及缺层级人工确认；状态见 [`s6-status.json`](acceptance/remediation/s6-status.json)。

**文件位置**：改进 `tests/test_post_review_word.py` 的 NW06/NW08/NW09，新增 `tests/test_remediation_word.py`；新增独立匿名夹具生成器与静态期望清单，保留旧夹具作为基础回归。

### 9.1 每个场景先检查夹具事实

在运行 Word 之前断言夹具确实包含目标特征；不满足时测试直接失败。例如要求奇偶页眉，就必须检查开关和两个不同关系/内容，不能只创建普通 header。

| ID | 夹具必须真实包含 | 最终成果必须断言 |
| --- | --- | --- |
| SW01 | S1 的 12/24 pt 与不同字体文本区间 | preserve 构建成功，最终 DOCX 区间属性匹配，PDF可见区间不异常 |
| SW02 | default/first/even 不同文字；足够页面；共享/解除链接各一例 | 各节绑定、两个距离、每种启用变体在对应 PDF 页出现；不串节 |
| SW03 | 封面、声明、中文摘要、实际英文摘要、目录、两章、参考文献、附录；明确 source.regions/独立内容部件 | 每个部件来源范围正确、正文无缺失/重复，前置序列与正文重启符合配置 |
| SW04 | 纵向正文与横向表格节，含 SW02 的部分页眉变化 | 每节与每页纸张/方向正确，横向页内容和页眉符合绑定 |
| SW05 | 两个不同来源都用 noteId=1；A/B 注释文字不同且各含不同链接/图片 | A 引用仍指向 A 定义及关系，B 同理；ID 唯一、定义/引用数正确，不只是字符串至少出现两次 |
| SW06 | 来源图、合法重复资源引用、生成封面图 | 每类对象在正确部件/锚点，计数与许可匹配；多复制一图的故障例被拒 |
| SW07 | 缺层级样例及另一个目标稿件 | 错误层级不自动应用；人工确认后正确角色和格式实际生效 |

中文/英文摘要使用两份语言正确的匿名文字，不再复用中文生成函数仅改标题。预计页码由声明的序列和实际 Word 测量共同核验；不可手填 pages、手工合并最终文档或绕过门禁来完成场景。

### 9.2 断言组织

每个测试输出“夹具事实 → plan → doctor → build → 最终 OOXML → 同次 PDF → 发布状态”。对本批新增特征设置显式断言，不能只检查文件存在、退出码为 0 或某个 XML 标签出现。

最终 PDF 需要检查封面、节边界、奇偶页眉、横版图表及注释页的可见内容。机器检查证明对象、属性与标签；视觉检查补充裁切/重叠，不以截图相似代替语义。

对每个缺失场景至少设计一个会让断言失败的夹具变体，例如关闭奇偶页开关、让 B 引用 A 注释、移除英文摘要选区。先证明测试能发现这些错误，再将正常场景标为通过。

**退出条件**：原 D06 的 NW06/NW08/NW09 缺口被真实输入与最终断言覆盖；配置无法表达的必要能力须补实现或保持 blocked，不只调整测试名称。

## 10. S7：在真正的失败构建中检查发布事务

**文件位置**：`tests/test_publication_gates.py`、`tests/test_delivery_pipeline.py`、NW12；必要的测试宿主与生产 I/O 依赖注入接口。复用 `lib/engine.py:_publish_deliveries` 的现有事务路径。

### 三层验证，分别记录证据

1. **纯生产 CLI 第二次失败构建。** 首次构建成功，保存全部已声明 DOCX 和 build-metadata.json 哈希；第二次用无效/过期映射或确实不支持的配置运行同一输出目录，断言非零退出、明确失败诊断、全部旧成果未变。此例证明实际 CLI 失败保护，但不单独证明最终格式门禁或发布中途回滚。
2. **真实门禁故障注入。** 在测试宿主中，等待实际装配/必要 Word 导出完成后，仅损坏待检 staging 文件（S1 字号、S2 header、S3 额外图片），随后调用真实验证器和真实发布控制流程。可通过内部依赖注入或 CLI main 的测试宿主实现；不得 mock `passed` 或直接返回验证结果，不引入面向用户的绕过 QA 参数。明确标记为测试宿主注入，不能冒充完全无注入的命令行路径。
3. **发布中途失败。** 至少两个交付物，定点注入第二个替换操作的 I/O 异常、元数据写入异常，验证回滚范围包括全部 DOCX 和元数据，文件不存在状态也要恢复。故障只针对测试目录/明确一次操作，不破坏其他任务。

### 必需断言

- 失败的第二次操作确实到达所测阶段，日志记录注入位置、实际失败检查或 I/O 操作。
- 旧成果及旧元数据逐文件 SHA-256 不变；不出现混合批次成果、半份元数据或暂存成果冒充正式交付。
- 本次新诊断包含失败阶段和对象定位，且不改写上次成功元数据为本次失败状态。
- 没有预先存在成果时，失败后不留下宣称成功的文件；存在多个成果时检查全部文件，而非只检查第一个。
- 字段不收敛用有界测量序列的测试注入验证；Word 正常成功路径仍需真实执行，二者证据分开。

**退出条件**：NW12 不再是“损坏另存副本后检查从未被写入的原文件”；三个层次均有对应失败证据与事务断言，原 D06 的事务缺口关闭。

## 11. S8：最终回归、证据固定与能力状态

**当前进度（2026-09-08）**：完整离线套件为 325/325 通过（6 项条件跳过）；真实 Word/NW、remediation、格式、M1、论文、分页及发布门禁索引已固定，最小匿名示例已完成 plan→doctor→build→smoke_test；S5 真人审阅与跨目标真实 Word 交付证据已加入。S8 在声明范围内标为 `verified`，剩余限制仅是能力不外推到未复验的模板、冻结集或环境。

### 11.1 回归范围

完成定向检查后再运行完整离线与 Word 组。命令从仓库根目录执行；新增测试文件在对应批次创建，缺文件不能当作无测试可跳过。

```bash
python3 -m unittest discover -s tests -p 'test_remediation_contracts.py' -v
python3 -m unittest discover -s tests -p 'test_role_evaluation.py' -v
python3 -m unittest discover -s tests -v

DOCUMENT_SYNTHESIS_WORD_TEST=1 DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access python3 -m unittest discover -s tests -p 'test_remediation_word.py' -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access python3 -m unittest discover -s tests -p 'test_post_review_word.py' -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access python3 -m unittest discover -s tests -p 'test_word_formatting.py' -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access python3 -m unittest discover -s tests -p 'test_m1_acceptance.py' -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access python3 -m unittest discover -s tests -p 'test_thesis_acceptance.py' -v
git diff --check
```

其他分页测试若有独立 opt-in 开关，读取实际代码并补充正确调用。逐项记录 skip 原因；本轮必需 Word 场景不得 skip。检查格式验证、元数据契约的变更未使旧封面、无目录 PAGEREF、迁移和已通过的文字实例检查回退。

### 11.2 正式证据

新增 `docs/acceptance/remediation/sX-status.json` 和每个 SW 场景目录，保存：

- 代码/配置/格式包/源夹具哈希、执行命令及环境版本；
- 夹具事实断言、正常对照和损坏例的结果；
- 成功成果的最终 DOCX、同次原始 PDF、对应 metadata；
- 被拒绝 staging 文件及拒绝原因，明确不能作为合格示例；
- 节/story/对象/属性覆盖、实际字段与页码、事务前后哈希；
- 新指标版本、分区数据哈希、人工标注/决策证据、剩余范围。

保存最终成果后再清理本轮 `.work`、锁文件、渲染与 Python 缓存。测试 TemporaryDirectory 不得在证据复制完成前删除待验收成果。原始 PDF 的已知 xref 警告继续保留诊断，不为消除警告重写原件。

### 11.3 状态更新规则

更新 `docs/formatting-progress.md`、`docs/formatting-capabilities.md`、上一轮计划的当前摘要，以及 README/示例中实际受影响的说明。历史成功记录保留，追加本轮缺陷关闭证据。

每批仅在其退出条件满足后标为 verified。候选算法实现完成但自动门槛未达，应标记“辅助校正已验证、自动映射 partial”；未覆盖的 story/对象/节组合明确未验证。不能把较窄范围的通过写成 S0–S8 全部完成。

## 12. 最终验收对照表

| 问题 | 正常对照 | 破坏性对照 | 关闭证据 |
| --- | --- | --- | --- |
| D01 | 12/24 pt、不同字体及 run 重组后正确通过 | 改坏某逻辑区间被拒 | S1 + SW01，含真实 CLI 成果 |
| D02 | 各 story/节绑定、文字和距离正确 | 独立修改内容/属性/关系/距离被拒 | S2 + SW02/SW04，含 PDF 页级对应 |
| D03 | 合法资源复用与生成对象通过 | 多/少/错位对象被拒 | S3 + SW06，含实例级诊断 |
| D04 | 缺层级可正确推断或明确待确认 | 不再高分自动压平层级 | S5 + SW07，新冻结集证据 |
| D05 | 手算指标与评估器一致 | 一对一错不能报告 exact=1.0 | S4，指标 v2 回归与重评 |
| D06 | 夹具事实与最终断言均满足场景 | 缺关键特征或第二次构建失败被检出 | S6/S7，场景清单逐条关闭 |

可靠格式交付的先行放行要求 D01–D03 关闭并完成相关真实 Word/事务验证；自动角色映射放行要求 D04/D05 与精确指标门槛同时满足；完整论文范围放行要求 D06 所列具体场景完成。保留这种分项结论，避免再次用全套测试的总数掩盖尚未覆盖的能力。
