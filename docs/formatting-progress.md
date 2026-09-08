# 自定义格式与样例识别：实施进度与下一步规划

日期：2026-09-06。历史 R11 状态为 `verified`（历史 R0–R11 批次完成了当时范围的实现与验收）；post-review N0–N10 是前置基线，当前补全执行以 [`post-n0-n10-remediation-plan.md`](post-n0-n10-remediation-plan.md) 及 [`acceptance/remediation/status-summary.json`](acceptance/remediation/status-summary.json) 为准。当前 S0–S8 已在声明范围内验证；离线全套 325/325 通过（6 项条件跳过），NW01–NW12 为 12/12，`remediation_word` 为 11/11。S5 真人审阅与跨目标真实 Word 交付哈希见 [`acceptance/remediation/s5-human-review-evidence.json`](acceptance/remediation/s5-human-review-evidence.json)；结论不构成无条件的全校规范支持声明。

历史章节记录了曾经的模块实现和测试结果，不构成当前发布声明。当前状态以补全计划、缺陷报告、契约测试和真实 Word/PDF 证据为准；历史 R7/R11 的流程 `verified` 不等于当前 S5 自动角色映射已放行。

## 0. Post-review N0–N10 前置基线

> 本节保留前置 N0–N10 的基线。当前 S0–S8 补全状态不覆盖或改写历史结论，详见 [`acceptance/remediation/status-summary.json`](acceptance/remediation/status-summary.json)。

优化计划的机器可读状态见 [`acceptance/post-review/status.json`](acceptance/post-review/status.json)，逐批证据见同目录的 `n0-status.json` 至 `n10-status.json`；离线与真实 Word 链路记录见 [`offline-verification.json`](acceptance/post-review/offline-verification.json)。前置 N0–N10 已完成其声明范围的退出验证，N7 已达到冻结集自动接受门槛，N9 已完成已知 Quartz orphan-xref 模式诊断。该批次历史基线为离线 286 项通过、3 项条件跳过，固定 Word Automation 目录后的真实 Word 矩阵为 12/12 通过；当前补全回归数字以本文件顶部及 [`acceptance/remediation/status-summary.json`](acceptance/remediation/status-summary.json) 为准。真实产物与环境哈希见 [`n8-status.json`](acceptance/post-review/n8-status.json)；历史 R0–R11 状态文件仍保留，不覆盖当前 S0–S8 补全结论。

N7 的独立无大纲集包含 20 份匿名虚构文档、100 个标题正例和 500 个正文负例，标签绑定到独立人工可见段落基线。该前置批次的前三候选标题覆盖率为 100%，候选块 precision 为 1.0；自动接受覆盖率为 100%、precision 为 1.0，正文误报为 0，达到 98% precision/50% coverage 门槛。能力结论限定于该匿名冻结集；当前 S5 仍需真实审阅者操作证据，不能仅以此历史冻结集结论放行自动角色映射。

N9 的 PDF 诊断见 [`acceptance/post-review/pdf-diagnostics/summary.json`](acceptance/post-review/pdf-diagnostics/summary.json) 与 [`n9-status.json`](acceptance/post-review/n9-status.json)。同源真实 Word→PDF 导出可由 pypdf、PyMuPDF 和 pdftotext 读取 6 页；pypdf 报告的 3 个 Quartz wrong pointing object offset 警告已确认对应未引用、文件中不存在的 orphan xref 行，页数和页面框一致，N9 通过非破坏性诊断分类。其他 PDF 结构异常仍应阻断。

## 1. 总体实施进展总览

进度表将“模块实现”与“端到端验证”分开。测试数量或离线装配通过，不能替代真实交付证据；未验证的高风险能力不得标记为 `supported`。

| 批次编号 | 阶段名称 | 模块实现 | 端到端验证 | 主要证据与当前结论 |
| :--- | :--- | :--- | :--- | :--- |
| **B0** | 基线锁定与高风险能力验证 | `verified` | `verified`（探针层） | 夹具与环境记录已留存；不等同于 M1/M2 发布验收。 |
| **B1** | 格式契约与兼容适配 | `implemented` | `not_verified` | 格式包、分析/决策/映射 v1 契约已补齐；真实交付验证仍待后续批次。 |
| **B2** | 只读 DOCX 检查与有效属性 | `implemented` | `partial`（R6 离线） | 有效属性按文档默认→样式→直接格式求值，逐 run 记录证据；真实 Word/PDF 仍未验证。 |
| **B3** | 统一样式应用与旧规则隔离 | `implemented` | `partial`（F04/R6 离线） | preserve/mixed、样式托管属性、生成标题/目录与逐项格式证据已接入；完整发布矩阵待 R8/R10。 |
| **B4** | 渲染上下文与包导入 | `implemented` | `partial`（F04/R5 离线已通过） | RenderContext、页面规格、版心适配和选区 OPC 导入已接入；复杂 story 与真实矩阵仍待后续批次。 |
| **B5** | 通用单 DOCX 与角色映射 | `implemented` | `verified`（F03/F04） | RoleMap 消费、preserve/source 和多来源格式执行已接入离线生产链路。 |
| **B6** | 样例识别与 HTML 校正 | `implemented` | `verified`（R7） | 字段证据、人工决策、近似预览、格式包及 RoleMap 跨目标复用已通过；样例正文不进入格式包。 |
| **B7** | M1 验收与发布元数据 | `implemented` | `verified`（R10） | M1 三项真实 Word 验收通过；DOCX/PDF/配置/页码证据与原子发布元数据已固定。 |
| **B8** | 文档抽象部件与源选区 | `implemented` | `verified`（离线契约；真实 W05/W06 交付） | 选区对象、关系闭包、交付范围清单和失败回退已验证；完整复杂 story 仍不作无条件支持声明。 |
| **B9** | 编号区间与真实 Word 测量 | `implemented` | `verified`（R8/W05） | 罗马/阿拉伯标签、重启、oddPage 物理页与空白页白名单均有真实 PDF 证据。 |
| **B10** | 论文对象、字段与 M2 验收 | `implemented` | `verified`（R9/W06；R10/W07） | SEQ/REF/PAGEREF、脚注/尾注闭包与 OMML 有离线契约；W06 真实收敛与 W07 完整匿名论文真实 Word/PDF 矩阵均通过。 |
| **R0** | 反例、能力状态与证据固定 | `implemented` | `verified` | `r0_contracts.json`、离线契约测试和独立 Word 测试组已加入。 |
| **R1** | 错误发布与虚假 passed 封堵 | `implemented` | `verified` | 逐交付物必需门禁、页面几何、内容范围和独立失败诊断已接入；F02 负向回归通过。 |
| **R2** | 配置、分析、决策与 RoleMap 契约 | `implemented` | `verified` | 三份 v1 schema、共享角色注册表、跨目录 provenance 和隔离环境 CLI 闭环已验证；R3 接消费。 |
| **R3** | 共用只读 BuildPlan，消费显式映射 | `implemented` | `verified` | `PreparedBuild` 冻结配置与 NodeRef/哈希；显式 RoleMap 控制样式和导航，过期映射及渲染前源篡改在写 run 之前拒绝；R4 接 preserve/mixed 多来源。 |
| **R4** | 统一样式、页面与多来源上下文 | `implemented` | `verified`（离线；Word/PDF 补测待环境） | RenderContext、preserve/mixed、`apply_section_spec()`、目标版心和生成样式已接通；F03/F04 正常回归通过。 |
| **R5** | 原始选区与 OPC 无损导入 | `implemented` | `verified`（离线） | `[start,end)`、PackageImporter、关系闭包、ID 映射和 F01/F07 已验证；复杂 story 与完整交付范围待后续批次。 |
| **R6** | 按交付范围验证内容和有效格式 | `implemented` | `verified`（离线） | `ExpectedInventory`、严格内容多重集、脚注/链接/字段依赖、逐 run 有效格式与 engine/smoke 同契约已验证；真实 Word/PDF 待 R8/R10。 |
| **R7** | 样例校正、预览与复用 | `implemented` | `verified` | 分析证据、人工决策契约、虚构近似预览、manifest 迁移和格式包/RoleMap 跨目标复用已通过。 |
| **R8** | 分节、页码与物理页面 QA | `implemented` | `verified`（离线 + W05） | PageRecord、真实 PDF 标签、序列重启、oddPage 物理奇偶和有意空白页白名单已通过。 |
| **R9** | 脚注/尾注与字段依赖更新 | `implemented` | `verified`（离线 + W06） | SEQ、精确书签范围、前向 REF、类型化 PAGEREF、循环诊断和尾注重映射已通过。 |
| **R10** | 完整 Word 回归与 PDF 证据 | `implemented` | `verified`（M1/W04–W07） | 246 项离线全套、M1 真实 Word 3/3、W04–W07 真实 Word 4/4、M2 装配级 1/1 及完整匿名论文真实 Word/PDF 1/1；R11 发布收尾已在后续批次完成。 |
| **R11** | 文档、示例与发布收尾 | `implemented` | `verified` | 自定义格式示例完成分析、最终决策、格式包、虚构近似预览、RoleMap、v1 迁移、plan/doctor、真实 Word build 和完整 smoke；正式状态见 [`r11-status.json`](acceptance/r11-status.json)。 |

## 1.1 R3 执行记录

- 新增 [`lib/build_plan.py`](../lib/build_plan.py)，将配置 provenance、源/格式/映射哈希、来源顺序、原始 NodeRef 索引、选区计划、最终 assignments、分节、交付声明和诊断集中到 `PreparedBuild`；运行时 `Document` 句柄不进入计划序列化。
- `UnifiedSynthesizer.plan()` 与 `synthesize()` 共用 `prepare_build()`；`render_body()` 只消费已经准备好的 assignments/tree，并在渲染前复核源、RoleMap 和格式继承链哈希。
- RoleMap 现在按完整 NodeRef 绑定至实际源节点，检查 source SHA-256、part URI、element path、story type 和 text hash。显式映射不再被无参 `RoleMapper()` 覆盖；`on_unmapped=error` 和 `preserve` 均有明确行为。
- 公开 `plan()` 输出为 JSON 可序列化结构，去除源正文；新增 `tests/test_r3_build_plan.py` 覆盖 heading.2 实际样式/导航、重复标题、过期映射和源篡改。
- 验证：全量 203 项，0 非预期失败、0 错误、2 skipped、9 expected failures；R3 定向 3 项通过；`minimal-demo --plan` 和 `git diff --check` 通过。

## 1.2 R4 执行记录

- `RenderContext` 贯穿 v3 单 DOCX、多来源递归、DOCX/PDF/图片/PPTX 导入和横竖版恢复；字体与旧全局横版状态只作为兼容入口保留。
- `apply_section_spec()` 现在实际写入纸张尺寸、方向、边距、页眉页脚距离和 `docGrid`；`preserve/source` 复制源节几何与简单页眉页脚，`restyle/target` 统一落目标规格。
- `apply_roles()` 的 preserve 为严格无操作；mixed 通过文件与 `node_range` 过滤 RoleMap，`on_unmapped=preserve` 不再隐式降级为正文。行内粗斜体、上下标、公式、超链接和原生编号不进入旧全局清洗。
- 生成标题、目录条目和页脚改走 `add_styled_heading()`、`add_styled_toc_entry()` 与 `setup_styled_footer()`；目录制表位按当前版心计算，PDF 缓存键纳入页面/格式/清洗策略，图片与表格按当前版心适配。
- F03 stale RoleMap、F04 多来源 37 pt 标题和 preserve/source 几何页脚契约均已转为正常回归；新增页面规格、preserve 无变更和 mixed 未托管段落测试。
- 验证：全量 `206` 项，`197 passed`、`0` 非预期失败、`0` 错误、`2 skipped`、`7 expected failures`；`git diff --check` 通过。

## 1.3 R6 执行记录

- `lib/content_integrity.py` 新增 `ExpectedInventory`、多来源合并和选区/部件边界投影；文本、表格、媒体、公式、脚注/尾注、字段依赖、超链接目标和书签均以严格清单核验。标题替换、生成内容和合法 ID 重映射显式记录，不再以整篇 body 或任意子串推断完整性。
- `lib/docx_inspector.py` 的有效样式求值补齐文档默认段落属性及 keep/widow、缩进等字段，并输出每个 run 的有效值；`verify_delivery_format()` 对 role→style、页面属性和原始 NodeRef 逐项写入 expected/actual/provenance/coverage。
- `lib/engine.py` 与 `smoke_test.py` 均调用同一清单构造/验证契约；`CheckResult.evidence` 将期望清单或格式报告随交付门禁写入 QA 元数据。
- R6 回归新增于 `tests/test_content_integrity.py`，并将 F08 转为正常回归；定向测试全部通过。全量：`225` 项运行、`222 passed`、`3 expected failures`、`2 skipped`、`0` 非预期失败、`0` 错误。
- 本批只证明离线 DOCX/装配层。真实 Word/PDF 的字段更新、页眉页脚、罗马标签、奇数页和物理页面证据仍按 R8–R10 处理。

---

## 1.4 R7–R10 当前执行记录（2026-09-05）

本节是对历史 B6–B10 章节的当前验收覆盖，批次证据分别见 [`r7-status.json`](acceptance/r7-status.json)、[`r8-status.json`](acceptance/r8-status.json)、[`r9-status.json`](acceptance/r9-status.json)、[`r10-status.json`](acceptance/r10-status.json) 和 [`r11-status.json`](acceptance/r11-status.json)。

- **R7 `verified`**：分析报告现在包含样本计数、样式使用/定义差异、字段证据、不支持特性与候选置信度；HTML 决策支持数值/候选修改，未决角色不能编译。近似预览使用虚构内容；manifest 迁移非破坏；格式包和 RoleMap 已在另一份目标上实际消费。`tests.test_r7_workflows`：8 项通过。独立匿名评估集 20 份、100 个标题正例、500 个正文负例，各级 precision/recall 1.0，正文误报 0，自动接受覆盖率/精确率 1.0。
- **R8 `verified`**：引入类型化 `PageRecord`，Word 导出后从同次 PDF 提取并核对实际标签；罗马/阿拉伯序列、重启、oddPage 物理奇偶与有意空白页白名单已接入发布门禁。离线 R8 定向回归 54 项通过；W05 真实 Word/PDF 通过。
- **R9 `verified`**：字段更新器支持 SEQ 格式/重启/大小写、简单及跨 run 复杂字段、精确跨段书签范围、前向 REF、类型化 PAGEREF 与循环诊断；尾注 ID 重映射和关系闭包有回归。`tests.test_r9_fields_and_notes tests.test_notes_and_fields`：15 项通过；W06 真实 Word/PDF 通过。
- **R10 `verified`（当前范围）**：离线全套 246 项通过（2 项条件跳过）；M1 真实 Word 3/3 通过；W04–W07 真实 Word/PDF 4/4 通过；M2 匿名论文装配级回归 1/1 及完整匿名论文真实 Word/PDF 矩阵 1/1 通过。元数据固定 DOCX/PDF/配置/覆写哈希、PageRecord、QA 与 Word 环境信息。
- **R11 `verified`**：根目录可复现命令已统一到当前 CLI；自定义示例在干净输出目录完成分析、决策编译、预览、RoleMap、迁移、plan、doctor、真实 Word build 与完整 smoke。迁移输入未被覆盖，预览元数据明确虚构/近似且不含样例正文，build metadata 保存逐交付物内容、格式、页码与环境证据；R11 状态文件记录命令、结果、哈希和清理动作。
- **边界**：当前证据仍不外推为 OCR/PDF 学习、双栏期刊、参考文献著录转换、多样例融合或全部学校规范支持。

## 2. B0 实施详细记录

### 2.1 运行环境基线
- **操作系统**：macOS (Darwin 25.3.0, arm64)
- **Python 环境**：Python 3.14.0
- **Microsoft Word**：Microsoft Word for Mac，AppleScript automation 探针响应成功
- **核心依赖**：`jsonschema` 4.26.0, `python-docx` 1.1.2, `lxml` 6.0.0, `pymupdf` 1.26.x, `pypdf` 6.0.x, `Pillow` 12.0.x
- **代码库基线**：保留全部用户未提交工作区修改，原有 67 项单元测试全部 PASS，`minimal-demo` 与 `chinese-demo` 只读 `--plan` 执行正常。

### 2.2 测试夹具生成器 (`tests/fixtures/formatting/`)
- 创建了 `tests/fixtures/formatting/README.md`，确立测试样本零业务材料与隔离运行准则。
- 创建了 `tests/fixtures/formatting/generate_samples.py`，支持编程式生成：
  - `mixed_fonts.docx`：中西文多字体与直接格式取消加粗 (`w:b w:val="0"`)；
  - `complex_fields.docx`：简单字段 (`w:fldSimple`) 与跨 run 复杂字段 (`w:fldChar` begin/separate/end 状态机)；
  - `sections_and_pagination.docx`：分节页码格式 (`upperRoman` vs `decimal`) 与奇数页起章 (`oddPage`)。

### 2.3 高风险能力状态矩阵 (`docs/formatting-capabilities.md`)
通过真实脚本与真实 Word 自动化探针验证，确立支持矩阵：
1. **中西文混合字体与直接格式覆盖 (`styles.fonts.mixed`)**：`supported`。可精确解析 `w:rFonts` 的 `eastAsia` 与 `ascii`，`w:b w:val="0"` 成功解析为显式取消加粗。
2. **跨段/跨 run 复杂字段 (`fields.complex`)**：M1 `preserve_only`，M2 `supported`。严禁粗暴清空 run 重建段落，必须维护 `fldChar` 闭合。
3. **脚注引用 (`notes.footnotes`)**：M1 `preserve_only`，M2 `supported`。通过 OPC Part/Relationship 维系单文件副本，跨文件合并需 ID 重映射。
4. **罗马页码与分节标签 (`pagination.roman`)**：M1 `unsupported`（保留阿拉伯整数），M2 `supported`。真实标签需结合 PDF 文本提取核验。
5. **奇数页起章留白 (`layout.odd_page`)**：M1 `unsupported`，M2 `supported`。真实 Word 导出 PDF 奇数起章留白需在 QA 中做规则追踪。

---

## 3. B1 实施详细记录

### 3.1 契约规范与 JSON Schema (`schemas/`)
- `schemas/format-v1.schema.json`：定义格式包结构、单位约束（mm/pt/char）与语义角色注册表；
- `schemas/project-v3.schema.json`：定义 v3 清单规范，强制 `format` 引用与 `formatting` 策略，彻底禁止顶层混用 `page_setup` 与 `fonts`。

### 3.2 数据模型与单位度量 (`lib/format_schema.py`)
- 定义了不可变数据类：`LengthValue`, `LineSpacing`, `RunStyle`, `ParagraphStyle`, `StyleDefinition`, `PageSpec`, `RoleSpec`, `TocSpec`, `HeaderFooterSpec`, `FormattingPolicy`, `ResolvedFormat`, `Diagnostic`；
- 实现了统一的单位换算（mm/pt/twip/char），版心有效几何校验与首行/悬挂缩进互斥校验；
- 确立了稳定诊断码（`FORMAT_CYCLE`, `INVALID_GEOMETRY`, `UNSUPPORTED_CAPABILITY` 等）。

### 3.3 内置官方预设 (`formats/presets/`)
- `legacy-official/1.0.0.json`：党政机关公文格式基线（GB/T 9704 风格版式）；
- `report-basic/1.0.0.json`：现代商务与技术报告版式；
- `academic-basic/1.0.0.json`：基础学位论文规范版式（宋体/Times New Roman、1.5倍行距、首行2字符缩进）。

### 3.4 解析与继承引擎 (`lib/format_resolver.py`)
- 支持 `preset:` 与相对路径解析，`extends` 单继承循环检测（最大深度 8 层）；
- 纯函数合并与叶子属性来源追踪（`provenance`），64位内容哈希计算，能力匹配拦截。

### 3.5 历史版本适配器与清洗隔离 (`lib/legacy_format.py` & `lib/config.py`)
- `from_project_config` 将 v1/v2 配置映射为规范 `ResolvedFormat`；
- `LegacyPolicy` 隔离旧清洗逻辑，v3 默认关闭；
- `lib/config.py` 支持 v3 严格校验，`ProjectConfig` 暴露 `resolved_format` 与 `formatting`。

---

## 4. B2 实施详细记录 (只读 DOCX 检查与有效属性求值)

### 4.1 规范节点身份 (`NodeRef`)
- 在 `lib/docx_inspector.py` 中实现了不可变类 `NodeRef(source_sha256, part_uri, element_path)`；
- 路径采用确定性 XML 元素层级索引（如 `/w:document/w:body/w:p[1]` 或 `/w:document/w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]`），彻底解决了同文同名段落的唯一身份识别问题。

### 4.2 安全限额与资源防御 (`DocumentInspectionLimits`)
- 限制文档部件总数 $\le 10,000$；
- 限制单 XML 部件大小 $\le 32\text{ MiB}$，总解压大小 $\le 256\text{ MiB}$；
- 限制解压压缩比 $\le 200.0$，防范 ZIP 炸弹；
- 校验部件路径，严格禁止绝对路径或 `../` 路径逃逸；
- 严格拒绝 `.docm` 宏文档与二进制 `.doc` 文档进入分析流程。

### 4.3 OOXML 有效属性层叠求值算法 (`EffectiveStyleEvaluator`)
- **层叠拓扑**：
  $$\text{Effective} = \text{docDefaults} \oplus \text{Theme} \oplus \text{Style basedOn 链} \oplus \text{Direct Formatting}$$
- **多脚本字体解析**：精确提取 `w:rFonts` 的 `eastAsia`、`ascii`、`hAnsi`、`cs`，并将 `minorEastAsia`/`majorHAnsi` 等主题引用转换为具体字体名；
- **Toggle 属性精确识别**：
  - 成功区分属性缺失（未定义）与显式关闭；
  - `<w:b w:val="0"/>` 或 `<w:b w:val="false"/>` 能够正确覆盖父层样式的 `bold=True`，得到 `bold=False`；
- **段落度量与对齐求值**：解析字符缩进 (`w:firstLineChars="200"` $\to$ 2.0 char)、磅值行距 (`w:lineRule="exact"`) 与倍数行距 (`w:lineRule="auto"` $\to$ 1.5 倍)；
- **大纲层级 (`outline_level`)**：优先从 `pPr/w:outlineLvl` 解析（0-8 转为 1-9），若未定义则回溯样式链解析；
- **受保护对象标记**：自动识别 `field_complex` (`w:fldChar`/`w:fldSimple`)、`omml_math` (`m:oMath`)、`footnote_ref` (`w:footnoteReference`)、`image` (`w:drawing`/`w:pict`) 与 `hyperlink`。

### 4.4 历史警告清理
- 修复了 `lib/styles.py` 中直接对 `lxml.etree._Element` 求 bool 值产生的 `FutureWarning`，运行 `python3 -W error::FutureWarning` 零警告通过。

### 4.5 测试覆盖
- 新增 `tests/test_docx_inspector.py`（5 项测试全部通过）；
- 新增 `tests/test_effective_styles.py`（4 项测试全部通过）；
- 全套测试套件运行：**89 项测试全部通过 (OK, skipped=1)**；
- `git diff --check` 检查通过。

---

## 5. B3 实施详细记录 (统一样式应用与旧规则隔离)

### 5.1 模块架构与核心接口 (`lib/style_applier.py`)

1. **具名 Word 样式安装 (`install_styles`)**：
   - 为目标文档安装确定性的内部样式 ID：`SynthBody`、`SynthHeading1`..`SynthHeading9`、`SynthToc1`..`SynthToc9`、`SynthTitle`、`SynthSubtitle`、`SynthQuote`、`SynthCaption`、`SynthHeader`、`SynthFooter`；
   - 具有本地化显示名称（如 `SynthHeading1` 显示为 `Synth Heading 1`，并在 UI 样式库快速显示）；
   - **大纲级别绑定**：为 `SynthHeading1`..`SynthHeading9` 分别写入 `<w:outlineLvl w:val="0"/>` 到 `<w:outlineLvl w:val="8"/>`，保证 Word 原生导航窗格与 TOC 域代码精确识别；
   - **样式继承与段后样式**：`basedOn` 明确指向基准样式（如 `Normal`），`w:next` 指向下一个段落推荐样式（如正文）；
   - **XML 顺序规范化**：严格依照 ECMA-376 `PPR_ORDER` 与 `RPR_ORDER` 排布子元素，杜绝 Word 打开时弹出修复提示。

2. **角色驱动的样式应用 (`apply_roles`)**：
   - 目标段落设置 `w:pStyle w:val="Synth..."` 与大纲级别；
   - **属性所有权清理 (Property Ownership Clearing)**：清理目标样式所托管的直接格式属性（如段落直接设置的行距 `w:spacing`、缩进 `w:ind`、对齐 `w:jc`，以及 run 上的字号 `w:sz`、字体 `w:rFonts`、前景色 `w:color`），确保目标样式的统一性生效，而不被遗留的直接格式遮蔽；
   - **行内语义严格保护**：
     - 保留显式字符样式；
     - 保留局部斜体强调、上下标（`<w:vertAlign w:val="subscript"/>` / `superscript`）、超链接（`w:hyperlink`）、公式（`m:oMath`）、字段（`w:fldChar` / `w:fldSimple`）；
     - **严禁整段删除 run 覆写纯文本**，完整保持原 XML 树拓扑与内联对象。

3. **输出适配器 (`lib/style_applier.py`)**：
   - `add_styled_heading(document, text, level, resolved_format, bookmark_name)`：基于格式包样式添加规范标题，写入书签与大纲属性；
   - `add_styled_toc_entry(document, title, page_num, level, resolved_format, bookmark_name)`：根据格式包 `toc` 规格设置制表位、前导点并添加超链接目录行；
   - `setup_styled_footer(section, resolved_format, current_page, total_pages)`：基于格式包 `footer` 规则配置页脚（居中、两侧分布或右对齐，阿拉伯数字或页码域）。

4. **幂等性与样式切换**：
   - 格式应用完全幂等，同一文档连续多次应用等价；
   - 切换不同格式包（如从 `academic-basic` 切换至 `report-basic`）能够完全清洗旧格式所有权并应用新样式，无交叉残留。

### 5.2 测试覆盖与验收结果

- 新增 `tests/test_style_applier.py`，包含 5 项核心集成测试：
  - `test_install_styles_creates_synth_styles_and_outline_levels`：验证样式安装与大纲层级正确写入；
  - `test_apply_roles_clears_direct_formatting_and_sets_pstyle`：验证直接格式冲突清理与样式托管生效；
  - `test_preserve_inline_semantics_subscript_and_emphasis`：验证下标、斜体强调、公式在应用样式后 100% 保留；
  - `test_switching_between_formats_is_idempotent`：验证多格式包切换更新与幂等性；
  - `test_output_adapters`：验证标题、目录条目与页脚输出适配器功能。
- **全套测试套件运行**：**94 项测试全部通过 (OK, skipped=1)**；
- `git diff --check` 零违规。

---

## 6. B4 实施详细记录 (渲染上下文与包导入)

### 6.1 模块架构与核心交付 (`lib/layout.py` & `lib/package_importer.py`)

1. **版心几何度量与自适应排版 (`lib/layout.py`)**：
   - **`ContentBox` 数据结构**：以 twip 为基准单位封装页面与页边距几何，动态计算可用版心宽与高，提供 `cm`、`pt`、`mm` 等属性与 `orientation`（`portrait` / `landscape`）；
   - **`compute_content_box(spec_or_section)`**：支持传入 `docx.section.Section`、`PageSpec`（来自 `ResolvedFormat`）、`RenderContext`，彻底摆脱固定纸张硬编码；
   - **`constrain_image_box(image_width_px, image_height_px, content_box, max_height_ratio)`**：根据图像纵横比进行等比约束，优先贴合可用宽度，且高度不超过版心最大比例限制；
   - **`adapt_table_to_content_box(table, content_box)`**：动态检测 `w:tblGrid` 中各列总宽，当超出当前节可用版心时，按列宽比例等比缩小列宽并更新各行单元格 `w:tcW` 与表格 `w:tblW`，杜绝冲出右边距；
   - **`isolate_section_boundaries(section)`**：断开节的页眉页脚关联（`is_linked_to_previous = False`），防止保留区或前置部分修改后向主文或后继节泄漏。

2. **统一渲染上下文 (`RenderContext`)**：
   - 维护当前排版流程中的 `resolved_format`、`formatting_policy`、`current_section`、`exact_pages` 与 `resolved_files`；
   - 内部封装 `track_landscape` 与 `consume_landscape`，彻底替代了全局可变列表 `LAST_RENDERED_LANDSCAPE = [False]`，实现多任务与不同构建过程的状态隔离。

3. **通用受控部件导入器 (`lib/package_importer.py`)**：
   - **`PackageImporter` 核心类**：
     - **关系与媒体重映射**：扫描 `r:embed` / `r:id`，将图片字节从源文档安全注册到目标文档 package，分配全局唯一不冲突的 `rId`；将外部超链接安全挂载到目标关系表中；
     - **书签与绘图 ID 防冲突**：动态扫描目标文档中现有的 `w:bookmarkStart` 与 `wp:docPr`，为导入的对象统一顺延分配单调递增的非冲突 ID，杜绝 Word 打开时书签损坏；
     - **原生列表与编号保留 (`preserve_native_numbering`)**：
       - 将源文档 `numbering_part` 中的 `abstractNum` 与 `num` 定义克隆合并至目标文档；
       - 重映射 `abstractNumId` 与 `numId`，并在导入段落的 `w:numPr` 中更新映射，淘汰了直接扁平化为纯文本的旧有行为；
     - **样式隔离与命名空间保护**：
       - 源文档的未知自定义样式以 `SynthImport_{index}`（或通过 `style_prefix` 配置）安全重命名并登记到目标 `styles.xml`，绝不污染或冲刷目标文档的原有样式或 `Synth*` 格式包样式；
     - **OLE 与绝对定位清理**：
       - 将嵌入式 `w:object` 中的图片提取为静态 `w:pict`，移除绝对定位样式（`position:absolute` 等），防止重排后对象相互遮挡；
     - **向下兼容适配**：`lib/composition.py` 中的 `FrontImporter` 直接平滑继承 `PackageImporter`，现有 8 项合成装配测试 100% 通过。

4. **PDF 渲染缓存健壮性升级 (`lib/renderers.py`)**：
   - `get_cached_pdf_page_image` 的缓存键中纳入了清洗策略标志（`sanitize_footer`、`crop_whitespace`）与版本标识；
   - 确保修改去页码策略或裁切配置后，自动生成全新独立缓存，彻底根除脏缓存问题。

### 6.2 测试覆盖与验收结果

- 新增 `tests/test_layout_geometry.py`（7 项测试全部通过）；
- 新增 `tests/test_package_importer.py`（5 项测试全部通过）；
- 新增 `tests/test_formatting_modes.py`（3 项测试全部通过）；
- **全套测试套件运行**：**109 项测试全部通过 (OK, skipped=1)**；
- `git diff --check` 零违规。

---

## 7. B5 实施详细记录 (Universal Single DOCX Source & Role Mapping)

### 7.1 核心成果与改动概览
1. **语义角色映射器 (`lib/role_mapper.py`)**：
   - 实现了 `RoleAssignment` 数据模型与 `RoleMapper` 映射引擎，支持显式规则优先匹配与启发式规则自动推导（大纲级别 1-9、标准中英文标题样式名、表格、正文、副标题等）；
   - **`NodeRef` 绑定与确定性书签**：为标题段落生成包含位置哈希的唯一书签 `_Toc_{seq:04d}_{hash}`，基于确定的 `NodeRef` 节点路径对齐，彻底消除传统正则匹配因“重复同名标题”依次猜定位导致的书签错绑与目录跳转混乱；
   - 实现了 `validate_role_map`：校验角色映射集合的书签全局唯一性以及 `NodeRef` 指向源文档节点的有效性。

2. **单 DOCX 来源策略 (`lib/source_strategies.py`)**：
   - 新增 `docx_document` 来源策略，支持直接指定单个已有 `.docx` 文件或包含单个主文档的目录；
   - 自动调用 `RoleMapper` 进行只读检查与大纲抽取，不再依赖黄色高亮标记；
   - 支持纯正文无标题文档（body-only）：`heading_count == 0` 但 `content_block_count > 0` 仍为合法输入。

3. **构建计划增强与 SHA-256 源文件防篡改 (`lib/engine.py` & `synthesize.py`)**：
   - `UnifiedSynthesizer.plan()` 统一记录所有源材料文件的 SHA-256 哈希表 `source_hashes`、`content_block_count` 与 `heading_count`；
   - 若文档无标题且请求生成目录（`toc`），在 plan 中报告明确警告；在 `synthesize()` 中遇到无标题却请求 `toc` 时报错退出并提示移除 `toc` 或添加标题；
   - `UnifiedSynthesizer.synthesize()` 开工前调用 `verify_build_plan_hashes`，严格检查所有材料哈希，若材料在计划后被外部修改或删除则立即拒绝构建，杜绝输出过期或损坏的产物；
   - `run_doctor` 与 `print_plan` 适配 `content_block_count` 与 `heading_count`，不再对 `node_count > 0` 作生硬限制。

### 7.2 测试与验证结果
- **新增单元测试**：
  - `tests/test_role_mapper.py`（4 项测试）：验证同名标题通过 NodeRef 确定性消歧、大纲级别与表格角色映射、显式规则覆写、非法角色与重复书签诊断拦截；
  - `tests/test_docx_document_strategy.py`（4 项测试）：验证 docx_document 策略提取标准大纲标题（无黄色高亮）、纯正文 body-only 文档在 plan 中的合法性、请求目录时的提示警告、直接单文件路径解析；
  - `tests/test_build_plan.py`（3 项测试）：验证 SHA-256 哈希计算一致性、plan 中源文件哈希与内容块统计、verify_build_plan_hashes 对篡改和删除的拦截。
- **全套测试回归结果**：
  - 执行命令：`python3 -W error::FutureWarning -m unittest discover -s tests -v`
  - 结果：**120 项测试全部通过（120 passed, 0 failed, 1 skipped, 0 warnings）**。
- **CLI 命令验证**：
  - `python3 synthesize.py --project examples/minimal-demo --plan` 正常输出顶级目录、内容块与标题计数；
  - `python3 synthesize.py --doctor` 与 `--project examples/minimal-demo --doctor` 环境和材料检查全部通过。

---

## 8. B6 实施详细记录 (样例分析、校正报告、格式包复用)

### 8.1 样本文档样式分析与角色候选推导 (`lib/format_analysis.py`)
1. **视觉样式聚类 (`analyze_format_sample`)**：
   - 调用 `inspect_docx` 获取文档全部段落与表格的版面几何与有效属性；
   - 提取段落样式特征向量（中西文字体、字号、加粗、斜体、对齐方式、首行缩进、行距、前后段距）；
   - 按特征一致性自动合并为 `StyleCluster`，统计出现频次、关联节点与代表性样本片段。
2. **角色候选推导 (`FormatCandidate`) 与字段级证据链**：
   - 启发式匹配标准语义角色（`title`, `heading.1`..`heading.9`, `body`, `table.body` 等）；
   - 为每个推导角色生成字段级证据链（如属性来源、支持段落数、代表性节点 ID 与文本摘要）；
   - 识别样例中缺失的必需角色（`missing_roles`），并给出默认继承建议。
3. **严格防泄露保护**：
   - 样本分析产生的格式属性纯粹为样式参数，绝不保留或泄漏样例正文；
   - 样本片段仅用于人工审查证据，不进入编译包。

### 8.2 本地交互 HTML 校正报告与格式包编译器 (`lib/format_review.py`)
1. **自包含单文件 HTML 校正报告 (`generate_html_review`)**：
   - 纯原生内嵌 CSS 与 JS，零外部网络依赖，零 CDN；
   - 所有文本与属性严格进行 HTML 实体转义，防止 XSS；
   - 提供深浅质感卡片布局，清晰展示样本元数据、候选角色证据链、中性虚构文本预览以及交互决策面板；
   - 提供浏览器端一键导出 `decisions.json`。
2. **格式包编译器 (`compile_format_package`)**：
   - 读取分析报告与用户决策，校验 `report_id` 与 `source_sha256` 防篡改绑定；
   - 与基准预设（`base_format_ref`）安全继承合并；
   - 经 `schemas/format-v1.schema.json` 模式校验，输出版本化标准 `format.json`；
   - 内置零正文泄露断言，杜绝样本文本流入编译包。
3. **目标角色映射编译器 (`compile_role_mapping`)**：
   - 将分析推导出的节点角色与用户决策编译为绑定源文件哈希的 `roles.json`（`RoleMap` 格式）。

### 8.3 CLI 入口分发与端到端示例验证
1. **`synthesize.py` 增强**：
   - 新增操作互斥检查：分析/编译子命令与构建/计划/诊断命令严格互斥；
   - 新增子命令：`--analyze-format`、`--compile-format`、`--analyze-source`、`--compile-mapping`、`--replace-output`。
2. **端到端示例 (`examples/custom-format-demo/`)**：
   - 包含 `sample.docx`（格式样本）、`manuscript.docx`（未排版稿件）、`manifest.json`（v3 标准清单）与 `README.md`；
   - 完整验证：样例分析 $\to$ 格式包编译 $\to$ 目标映射 $\to$ v3 规范排版 $\to$ 成果输出。

### 8.4 测试覆盖与验收结果
- 新增 `tests/test_format_analysis.py`（2 项测试全通过）；
- 新增 `tests/test_format_review.py`（4 项测试全通过）；
- 新增 `tests/test_format_cli.py`（3 项测试全通过）；
- **全套测试回归**：**129 项测试全部通过（129 passed, 0 failed, 1 skipped, 0 warnings）**；
- `git diff --check` 保持零违规。

---

## 9. B7 实施详细记录 (M1 验收、内容完整性校验、有效格式核验与构建元数据)

B7 作为 Milestone M1 的收官批次与发布门禁，建立了完整的内容语义审计体系、交付格式有效性核验机制与事务式原子发布回滚通道。

### 9.1 内容语义完整性提取与校验 (`lib/content_integrity.py`)
1. **语义清单提取 (`extract_semantic_inventory`)**：
   - 提取可见文本段落序列与多重集频次（`Counter`），防范文本被丢弃或错误倍增；
   - 提取表格拓扑结构（行数、列数、所有单元格文本矩阵）；
   - 提取嵌入媒体哈希（SHA-256）；
   - 提取 OMML 数学公式规范化表示（使用 XML C14N 规范化并计算 SHA-256）；
   - 提取脚注引用、超链接目标及字段标记。
2. **内容完整性校验 (`verify_content_integrity`)**：
   - 支持显式声明的标题文本替换（`replacements` 映射）；
   - 采用多重集计数法（而非容易受子串包含干扰的单向搜索），严格比对段落频次；
   - 严格比对表格拓扑结构与单元格内容；
   - 校验图片、公式与链接对象的完整存在；
   - 发现任何遗漏、重叠或篡改立即抛出诊断或返回校验失败。

### 9.2 交付物最终有效格式校验 (`lib/content_integrity.py`)
1. **`verify_delivery_format`**：
   - 使用 `EffectiveStyleEvaluator` 深度解析装配与分节完成后的交付 DOCX 文件；
   - 对标配置的 `ResolvedFormat`，核验中西文字体（`eastAsia`/`ascii`）、字号（允许 0.25 pt 浮动公差）、行距类型与值、首行缩进（字符/磅值）、段前段后距、加粗倾斜等属性；
   - 返回结构化 `FormatVerificationReport`（包含通过项计数、验证角色、违规详情及未验证属性清单）。

### 9.3 构建审计元数据与事务原子发布 (`lib/engine.py`)
1. **`generate_build_metadata`**：
   - 自动收集 `generator` ("document-synthesis")、`engine_version` ("3.0.0")、`schema_version`、`project_name`、`build_timestamp` (ISO 8601 UTC)；
   - 记录生效格式包 ID、版本、请求能力列表及内容哈希；
   - 记录所有源文件输入时的 SHA-256 快照；
   - 记录所有交付 DOCX 的 SHA-256；
   - 留存 QA 与内容完整性、格式核验汇总；
   - 记录运行环境（Word 自动化可用性与平台信息）。
2. **事务原子发布与回滚保证 (`_publish_deliveries`)**：
   - 仅对 v3 清单生成并发布 `build-metadata.json`，确保 v1/v2 历史工程目录脚印完全不变；
   - `build-metadata.json` 在临时运行目录 `.work/run-*/` 中就绪，与交付文档一同暂存到 `staged`；
   - 在发布事务中若发生任何写入异常或操作系统错误，触发统一原子回滚，清理已拷贝文件，恢复原有状态。

### 9.4 Smoke 测试升级与 M1 端到端验收
1. **`smoke_test.py` 增强**：
   - 交付检查前优先比对 `build-metadata.json` 中记录的 `source_hashes`，若源文件在构建后被修改则输出告警；
   - 接入 `verify_content_integrity` 与 `verify_delivery_format`；
   - 完美保持 `structure_only` 与历史 v1/v2 运行兼容性。
2. **`tests/test_content_integrity.py`**：6 项单元测试覆盖多重集校验、表格防篡改、合法替换与有效格式容差；
3. **`tests/test_build_metadata.py`**：3 项单元测试覆盖元数据字段完整性、事务原子移动与发布失败原子回滚；
4. **`tests/test_m1_acceptance.py`**：3 项端到端验收测试，覆盖：
   - `test_m1_custom_format_demo_full_loop`：样例分析 $\to$ 格式编译 $\to$ 目标映射 $\to$ v3 构建 $\to$ 完整性校验 $\to$ 元数据审计 $\to$ Smoke 检查闭环；
   - `test_m1_dual_format_switching_on_same_source`：同一稿件在学术与报告预设格式间一键切换，样式差异生效且正文内容零丢失；
   - `test_m1_body_only_document_acceptance`：纯正文无大纲标题文档在 v3 中的排版与元数据审计。

### 9.5 最终测试套件指标
- **当前测试总数**：**141 项测试（140 passed, 0 failed, 1 skipped, 0 warnings）**；
- **回归与兼容性**：v1/v2 历史清单工程与测试用例 100% 通过；
- **代码整洁度**：`git diff --check` 零告警，无 lxml / python-docx FutureWarning。

---

## 10. Milestone M1 完成总结与 M2 展望

### 10.1 M1 达成成果汇总
Milestone M1（通用格式与样例识别闭环）已彻底落地：
1. **契约标准**：建立了标准 `schemas/format-v1.schema.json` 与 `schemas/project-v3.schema.json`；
2. **分析与识别**：具备只读 OOXML 层叠求值、样式聚类、启发式角色识别与证据链追溯；
3. **人机协作**：提供无网络依赖的本地单文件 HTML 校正报告与浏览器端决策导出；
4. **统一应用**：具备具名 Word 样式安装、大纲级别绑定与行内复杂语义（上下标、局部强调、公式）保护；
5. **灵活来源**：支持标准单 DOCX 来源、无黄色高亮标题识别与纯正文无标题文档；
6. **质量门禁**：通过段落多重集与对象拓扑实现 100% 内容无损审计，输出版本化构建元数据。

### 10.2 后续 Milestone M2 路线图预告
在 M1 稳固的内核基础上，M2 聚焦高级排版、复杂对象与企业级协同能力：
- **文档部件与源选区 (`layout.parts.v1`)**：已在 B8 完成；
- **编号区间与真实 Word 测量 (`pagination.roman.v1`)**：将在 B9 实施；
- **高级复杂字段与交叉引用 (`fields.managed_update.v1`)**：将在 B10 实施；
- **高级脚注与尾注跨文档合并 (`notes.merge.v1`)**：将在 B10 实施。

---

## 11. B8 实施详细记录 (文档抽象部件与源选区 / Document Parts & Source Regions)

B8 作为 Milestone M2（多部件文档与高级排版）的首个核心实施批次，打破了历史版本对 `cover / toc / body` 固定部件三元组的硬编码依赖，建立了通用文档抽象部件注册表（`layout.parts`）与基于稳定 `NodeRef` 的源选区切片及内容完整性守护机制。

### 11.1 架构设计与契约模式 (`schemas/project-v3.schema.json`)
1. **`layout.parts` 注册表**：
   - 允许在清单中声明具名部件字典，每个部件包含 `id`、`kind`（`cover` | `generated_toc` | `content`）、可选的 `source_region`（关联源选区）、`section_ref`（分节版式引用）、`page_sequence`（页码序列绑定）及 `include_in_toc`（是否被目录索引）。
   - `output.documents[].parts` 由原先的枚举字符串数组演变为按顺序引用 `layout.parts` 注册表中的合法部件 ID。
2. **`layout.page_sequences`**：
   - 为后续 B9 预留页码序列规则字典，支持 `format`（`decimal`, `upperRoman`, `lowerRoman` 等）与 `start` 属性。
3. **`source.regions` 源选区规范**：
   - 允许为源文档定义具名顶层块切片选区，定义 `start` 与 `end` 路径（如 `/w:document/w:body/w:p[1]`，支持 `end_of_document`），以及 `exclude: true` 显式合法排除标记。

### 11.2 文档抽象部件与选区验证核心 (`lib/document_parts.py`)
1. **数据模型**：
   - `DocumentPart`：具名文档部件（自动将兼容别名 `toc` 规范化为 `generated_toc`）；
   - `SourceRegion`：源选区定义，承载 NodeRef 块路径；
   - 专用异常层级：`SelectionError`、`SelectionOverlapError`、`UnassignedContentError`、`InvalidBoundaryError`。
2. **顶层块扫描与边界合法性约束 (`SelectionValidator`)**：
   - `get_top_level_blocks(doc)`：遍历文档主体的直接子元素（`w:p` 与 `w:tbl`），严格按照 OOXML 标准只接受顶层块作为选区边界；
   - **表格内部防破坏检查**：若 `start` 或 `end` 路径包含单元格或行内部路径（如 `w:tc`），立即抛出 `InvalidBoundaryError`，坚决禁止破坏表格完整性；
   - **选区非重叠与非空检查**：多选区间严格按前开后闭区间校验，若有交叉重叠抛出 `SelectionOverlapError`；
   - **零静默内容丢失校验 (`UnassignedContentError`)**：算法扫描文档中所有未分配给任何 `source_region` 的顶层块；若存在包含文本、图片、公式等实质内容的未分配块且未标记 `exclude: true`，立即抛出 `UnassignedContentError`，从根本上防止静默吞内容。
3. **安全物理切片与深拷贝重建 (`slice_document_by_region`)**：
   - 依据已校验的有效块索引区间，利用 `copy.deepcopy` 精确复制顶层块并填充到新文档中，保留所有的段落格式、表格格式及关联的嵌入关系。
4. **动态边界书签 (`get_part_boundary_bookmark`)**：
   - 针对历史默认部件保留原生兼容书签：`_Synth_cover`、`_Synth_toc`、`_Synth_body`（起点）；
   - 针对自定义部件生成规范边界书签：`_Synth_part_<id>_start` 与 `_Synth_part_<id>_end`，结束边界统一采用 `_Synth_part_<id>_end` 避免书签重名碰撞。

### 11.3 核心流程适配与兼容隔离
1. **清单加载与校验适配 (`lib/config.py`)**：
   - `_get_parts_registry`：若项目未声明 `layout.parts`（包括所有 v1/v2 项目及简单 v3 项目），通过适配器无缝注入默认的三部件注册表（`cover`, `toc`, `body`），实现 100% 向下兼容；
   - `_normalize_documents`：严格校验 `documents[].parts` 引用的每个部件 ID 是否在注册表中存在；
   - `ProjectConfig`：安全暴露 `self.parts`、`self.regions`、`self.layout` 属性供引擎与交付流水线消费。
2. **多部件合成与组装适配 (`lib/composition.py`)**：
   - 采用多态字典 `_BoundariesDict` 替代原有写死字典，既支持索引访问动态书签，又保证现有通过 `set(BOUNDARIES.values())` 做枚举断言的旧测试代码零破坏；
   - `assemble_document`：支持任意序列的自定义部件组合。针对多个部件，通过 OOXML 合法的节属性包装段落（嵌入在 `w:p/w:pPr` 内）建立分节，最后统一进行书签去重与编号更新。
3. **交付检查与流水线升级 (`lib/delivery.py` & `smoke_test.py`)**：
   - `required_bookmarks` 与 `validate_delivery_structure`：接受部件注册表，动态判定所需书签；
   - 停止以固定 `"body"` 字符串判断正文：改由 `part.kind == "content"` 动态判定内容部件，确保前言、附录、主体均可纳入内容完整性审计；
   - `validate_measured_delivery`：以首个内容部件所在物理页为基准核验连续正文编号；
   - `build_deliveries`：支持按内容部件就绪优先级自动对拆分交付物进行拓扑排序；
   - `smoke_test.py`：传入 `parts_registry`，对声明的所有正文部件执行语义清单完整性校验。

### 11.4 测试覆盖与验收结果
1. **新增测试套件**：
   - `tests/test_source_regions.py`：7 项测试，全方位覆盖顶层块扫描、选区非重叠、表格内部非法边界拦截、未分配实质内容拦截（零静默丢失）、显式排除及切片重建；
   - `tests/test_document_parts.py`：10 项测试，覆盖默认部件适配、自定义部件注册表解析、非法部件类型/重复/未声明拦截、自定义部件装配、动态边界书签、缺失书签拦截及 Word 测量多轮分页收敛。
2. **全仓回归指标**：
   - **全套测试回归：158 项测试全部通过（157 passed, 1 skipped, 0 failed, 0 errors, 0 warnings）**；
   - 历史 v1/v2 示例与测试用例 100% 保持原有行为，零兼容退化；
   - `git diff --check` 保持零违规。

---

## 12. B9 实施详细记录 (编号区间与真实 Word 测量 / Page Sequences & Real Word Measurement)

B9 作为 Milestone M2 的核心测量与版式批次，彻底解决了正式学位论文和技术报告中“多节独立页码、罗马/阿拉伯多格式混排、奇数页起章留白、真实 Word 标签核验与双趟目录收敛”的技术难题。

### 12.1 业务背景与技术目标
学位论文与大型正式报告具有复杂且严格的分节排版要求：
1. **多重页码序列**：封面无页码；前置部分（声明、摘要、目录）采用大写/小写罗马数字（如 `i`, `ii`, `iii`）并从 1 重新起号；正文主体采用阿拉伯数字（如 `1`, `2`, `3`）并重新起号；
2. **奇数页起章（`oddPage` 分节符）**：正文或关键章节必须在奇数页起始。若上一部分以奇数页结束，Word 排版会自动生成一个空白偶数页；
3. **真实标签核验与目录双趟收敛**：目录必须精准填入各部分的真实页码字符串（如 `"ii"`, `"1"`），且该页码必须与 Word/PDF 导出的实际排版完全一致；
4. **视觉 QA 白名单机制**：避免将奇数页起章所造成的有意留白页误报为“空白页故障”。

### 12.2 数据模型与数字转换工具 (`lib/pagination_types.py`)
1. **`PageSequence` 数据类**：
   - `id`: 序列标识（如 `front_seq`, `main_seq`）；
   - `format`: 格式类型（`decimal`, `lowerRoman`, `upperRoman`, `lowerLetter`, `upperLetter`）；
   - `restart`: 布尔值，标识是否在该序列首个绑定的分节重新开始编号；
   - `start`: 起始编号数值（默认为 1，校验必须 `>= 1`）。
2. **`PageRecord` 数据模型**：
   - 标准化记录每页与每个书签的测量数据：`physical_page`（全局物理页号）、`printed_page`（节内打印整数值）、`expected_label`（格式化标签字符串，如 `"ii"`, `"1"`）、`observed_label`（从 PDF 页眉页脚提取的实际标签）、`section_id` / `sequence_id`，以及 `label_verified`（实际标签与期望标签一致性验证结果）。
3. **通用编号转换工具**：
   - `int_to_roman(n, lower=True)` 与 `roman_to_int(s)`：支持 1~3999 罗马数字双向转换；
   - `int_to_letters(n, lower=True)`：支持 Excel/Word 样式的 26 进制字母编号（`a...z, aa...`）；
   - `format_page_number(page_num, fmt)`：统一格式化入口，支持 `decimal`, `lowerRoman`, `upperRoman`, `lowerLetter`, `upperLetter`。

### 12.3 契约规范与能力注册 (`schemas/project-v3.schema.json` & `lib/format_resolver.py`)
1. **Schema 扩展**：
   - 在 `layout.parts` 中增加 `section_type` 属性，支持 `nextPage`、`oddPage`、`evenPage`、`continuous`（默认为 `nextPage`）；
   - 完善 `layout.page_sequences` 格式与数值模式校验。
2. **能力注册激活**：
   - 将 `"layout.parts.v1"` 和 `"pagination.roman.v1"` 从 `PLANNED_M2_CAPABILITIES` 正式提升为 `SUPPORTED_ENGINE_CAPABILITIES`；
   - 格式包声明上述能力可正常编译与应用，同时保持未就绪能力（`notes.merge.v1`, `fields.managed_update.v1`）的安全拦截。

### 12.4 Word 装配与分节排版升级 (`lib/composition.py`)
1. **分节页码类型注入 (`normalize_body_sections`)**：
   - 接收 `page_sequence` 与 `is_seq_start`；
   - 精确注入 OOXML `<w:pgNumType>` 元素：设置 `w:fmt`（如 `lowerRoman`）以及 `w:start`（当重新编号时设定起始数值）。
2. **跨部件合并与分节类型保持 (`assemble_document`)**：
   - 追踪序列初次出现标志（`is_seq_start`）；
   - 当将后续部件追加进文档时，通过分节属性段落（嵌入在 `w:p/w:pPr/w:sectPr` 内）保留前一部件的分节属性（包含 `w:type` 的分节类型如 `oddPage`、页眉页脚绑定以及页码格式）；文档末尾更新为最终部件的 `sectPr`。
3. **格式化目录页码写入 (`append_toc`)**：
   - 目录引导符后直接支持写入格式化字符串标签（如 `"ii"`、`"1"`）；兼容旧版整数传入。

### 12.5 真实 Word 测量与 QA 白名单 (`lib/pagination.py`, `lib/delivery.py`, `lib/qa.py`)
1. **PDF 页眉页脚真实标签提取 (`extract_pdf_page_labels`)**：
   - 基于 PyMuPDF 真实扫描导出 PDF 页面顶部与底部的页眉页脚文本块；
   - 提取识别罗马数字与阿拉伯数字，填充至 `page_map` 的 `observed_label` 并校验 `label_verified`。
2. **奇数起章偶数留白白名单 (`validate_measured_delivery` & `run_qa_assertions`)**：
   - 自动根据声明了 `section_type: "oddPage"` 的部件及其测量物理起始页，计算合法的偶数留白页集合 `allowed_blank_pages`；
   - 视觉 QA 模块 `run_qa_assertions` 支持接收 `allowed_blank_pages`；检测页面实质内容时，页眉页脚的孤立页码不会误判为正文内容，且白名单内的空白页判定为合法版式表现。
3. **双趟收敛自适应比较 (`build_deliveries`)**：
   - 比较逻辑采用字符串对齐（`str(guess) == str(exact)`），既支持整数也支持格式化标签，消除了因类型差异导致的无谓二次测量开销。

### 12.6 测试套件与验收结果
1. **新增专项目测试 (`tests/test_page_sequences.py`)**：共 10 项测试：
   - 罗马数字与字母序号转换、双向解析与非法值校验；
   - 清单序列合法性校验与未声明序列引用拦截；
   - 罗马数字与阿拉伯数字多节装配及 OOXML `w:pgNumType` 生成；
   - `oddPage` 奇数页起章与偶数空白页白名单视觉 QA 豁免；
   - 真实测量交付多节连续性核验。
2. **全仓回归指标**：
   - **全套测试回归：168 项测试全部通过（167 passed, 1 skipped, 0 failed, 0 errors, 0 warnings）**；
   - 历史 v1/v2 工程与默认 v3 工程 100% 保持原有行为，零兼容退化；
   - `git diff --check` 保持零违规。

---

## 13. B10 历史实施记录（论文对象、字段与 M2 验收）

以下内容保留当时的模块实现记录。R0 复核已确认其中若干结论不能作为当前 M2 放行依据，尤其是字段、页码和真实 Word 发布链路；最终状态以 R8–R10 的新证据为准。

### 13.1 跨文档脚注与尾注合并 (`lib/notes_merger.py`)
1. **系统保留条目保护**：
   - 严格识别并保留 OOXML 规范的系统分隔符（`w:id="-1"`, `w:type="separator"`）与延续分隔符（`w:id="0"`, `w:type="continuationSeparator"`），禁止将其重编号或覆盖；
2. **非冲突正整数 ID 自动重分配**：
   - 遍历目标文档已有的常规脚注 ID，自增计算可用正整数起点；
   - 为源文档导入的每个普通脚注分配新 ID，生成 `old_id -> new_id` 映射字典；
3. **关系递归挂载**：
   - 扫描脚注内部元素，若包含嵌入图片（`r:embed`）或外部超链接（`r:id`），在目标脚注部件的关系表（`_rels/footnotes.xml.rels`）中完成克隆与重新绑定；
4. **全链路集成**：
   - 集成至 `PackageImporter`：在 `import_element` 时自动扫描并更新 `w:footnoteReference` 与 `w:endnoteReference` 的 `w:id`；
   - 集成至 `SelectionValidator.slice_document_by_region`：源选区切片生成子文档时，自动复制切片所关联的脚注与尾注部件，杜绝切片导致的注脚遗失。

### 13.2 受控字段状态机更新与公式保留 (`lib/field_updater.py`)
1. **跨 run / 跨段复杂字段栈式解析**：
   - 基于 `_ComplexFieldState` 状态帧管理 `w:fldChar`（begin $\to$ instrText $\to$ separate $\to$ result_runs $\to$ end）；
   - 统一支持简单字段（`w:fldSimple`）与跨段复杂字段，禁止逐段截断破坏字段闭合；
2. **题注序号自动递增**：
   - 针对 `SEQ Figure` 与 `SEQ Table` 自动维护计数器，更新字段结果 run 中的显示文本（如 `"1"`, `"2"`, `"3"`）；
   - 支持 `\c` 重复当前序号与 `\r <n>` 显式重置；
   - 支持 `\* ROMAN` 格式化为大写罗马数字；
3. **交叉引用与页码引用**：
   - `REF <bookmark>`：关联书签对应段落文本或题注内容；
   - `PAGEREF <bookmark>`：结合测量 `page_map` 将目标书签的实际页码（如 `"8"` 或 `"ii"`）精准写回字段显示值；
4. **原生数学公式结构保留**：
   - 遍历与字段更新过程完全绕过 OMML 节点，严格保持 `m:oMath` 与 `m:oMathPara` 结构对象，绝不退化为图片或纯文本。

### 13.3 交付物强一致性与结构门禁 (`lib/delivery.py` & `lib/composition.py`)
1. **Staging DOCX 强一致写入**：
   - 在 `assemble_document` 阶段完成初步装配与 `SEQ` 编号更新；
   - 在 `build_deliveries` 多趟收敛过程中，获得测量页码后更新 `PAGEREF` 并写回 staging DOCX，确保发布前的 DOCX 内部显示值与导出的 PDF 完全一致；
2. **悬空未定义脚注引用门禁**：
   - 在 `validate_delivery_structure` 中增加完整性核验：若正文存在 `w:footnoteReference`，必须在 `footnotes.xml` 中存在对应条目，否则抛出 `OfficeExportError` 阻止发布。

### 13.4 引擎能力状态（R0 复核后）
- `notes.merge.v1` 与 `fields.managed_update.v1` 的模块代码仍存在，但不能仅凭历史注册或离线测试宣称已经完成生产验收；
- F06、F08、W06 作为当前负向证据保留，相关能力在能力矩阵中分别标为 `preserve_only` 或 `unsupported`；
- 只有 R8–R10 完成真实 Word/PDF 和发布门禁回归后，才可重新提升能力状态。

### 13.5 测试套件与综合验收指标
1. **专项目测试套件 (`tests/test_notes_and_fields.py`)**：共 6 项测试：
   - 多文档跨文档脚注合并、系统分隔符保留与 ID 映射；
   - 脚注内部超链接关系克隆与挂载；
   - `SEQ Figure` 与 `SEQ Table` 独立递增；
   - `PAGEREF` 引用页面解析；
   - OMML 原生公式零转图保留；
   - 悬空未定义脚注引用拦截。
2. **全要素论文综合验收套件 (`tests/test_thesis_acceptance.py`)**：共 1 项大型端到端测试：
   - 完整模拟具有双语摘要、前置罗马编号、正文阿拉伯编号、奇数页起章留白、多来源跨章节脚注、图 1/图 2/表 1 题注字段递增、OMML 损失函数公式及目录超链接的学术论文项目，全部断言一次性 PASS。
3. **历史回归指标（不作为当前放行条件）**：
   - 当时记录的 175 项自动化测试主要是模块/装配层测试，不能覆盖 R0 固化的 F01–F10 与 W04–W06 反例；
   - 真实 Word/PDF 测试与失败事务仍需按 R10 矩阵重新执行；
   - `git diff --check` 只说明文本差异格式正确，不替代功能验收。

---

## 14. Milestone M2 历史总结与当前支持边界声明

历史记录曾将 B8–B10 描述为完成，但 R0 复核发现完整生产链路仍有明确反例。因此当前不发布 M2 已完成或全部 supported 的结论。

### 14.1 M2 达成成果汇总
1. **部件与选区**：打破了历史写死 `cover/toc/body` 的限制，支持任意具名抽象部件（`layout.parts.v1`）与基于 NodeRef 顶层块切片的源选区机制（`source.regions`），实现零静默内容丢失校验；
2. **高级分页与测量**：支持前言罗马数字（`lowerRoman`/`upperRoman`）与正文阿拉伯数字（`decimal`）的多序列混排与重编（`pagination.roman.v1`），支持 `oddPage` 奇数起章留白白名单与 PyMuPDF 真实标签提取；
3. **复杂对象与引用**：支持多文档跨来源脚注与尾注合并（`notes.merge.v1`），支持题注序号（`SEQ Figure/Table`）与书签交叉引用（`fields.managed_update.v1`），原生保留 OMML 数学公式；
4. **强一致发布闭环**：实现双趟目录与字段收敛，确保 staging DOCX 与导出的 PDF 完全一致，元数据审计与事务回滚健全。

### 14.2 严格支持能力范围与免责边界 (Scope & Boundary)
根据执行计划 §12 与 §13 要求，当前仅确认模块实现，不提前发布 `supported` 能力：
- **模块已有但待端到端验证**：格式包继承、抽象部件、罗马/阿拉伯页码、奇数页起章、脚注/尾注、SEQ/REF/PAGEREF、OMML 保留和样例复用，详见 [能力矩阵](formatting-capabilities.md)。
- **当前明确不支持或仅保留**：需要结构改写的复杂字段、未完成关系闭包的选区切片、未完成真实页面标签验证的分页能力，以及未完成门禁覆盖的自定义交付部分。
- **明确不在本轮范围**：OCR/PDF 样例学习、双栏期刊、参考文献著录转换、多样例融合和全部学校规范认证。
