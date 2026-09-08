# B1–B10 验收后补全执行计划

日期：2026-09-06。状态：**R0–R11 已实施并完成各自当前验收范围**。本轮补齐样例校正/预览/复用、独立匿名识别评估、真实 Word 分节与页面证据、字段/注释依赖更新、M1/W04–W07 集成验收及 R11 发布收尾；未清理工作区中的既有用户改动。R10 已完成一份完整匿名论文的真实 Word/PDF 矩阵，R11 已完成文档/示例/能力矩阵同步，但不据此发布无条件的 M2 全能力声明。

依据：[初次验收 F01–F10](b1-b10-acceptance-report.md)、[真实 Word 补测 W01–W06](b1-b10-word-followup-report.md)。本计划在已有模块上补齐实现，替代“按 B1–B10 标记已完成后直接发布”的推进方式，不重复从零建设。

## 1. 交付目标与执行原则

最终用户流程必须实际跑通：选择内置或自定义格式包；或选择 DOCX 样例、查看格式证据并校正；确认目标稿件角色和选区；预检；用确定配置构建；通过内容、有效格式和真实 Word 检查后发布。

M1 完成通用单栏格式、样例学习、保留/重排及单/多来源；M2 在此基础上完成论文部分、双语摘要、前置/正文编号、页眉页脚、脚注/尾注和受控字段。具体学校规范仍应作为独立且有版本的格式包，不能把 academic-basic 称为全部学校通用的合规模板。

- 继续用 `synthesize.py` 作为唯一构建入口；新增操作由它分发，禁止新增绕过门禁的快捷构建脚本。
- 不改原稿、样例和已发布格式包；格式包修改产生新版本。保留当前未提交改动，不能清空工作区重建。
- 每批先加入能失败的回归，再修改生产代码；不能只写与当前实现一致的正向测试。
- `--plan` 不写文件、不调用 Word/模型/网络；`--doctor` 检查环境；最终布局结论来自 Word 和同次导出 PDF。
- 生产构建使用冻结后的配置、输入、映射和选区；识别结果不确定时等待校正或按明确保留规则执行。
- 不使用大范围放宽 QA 阈值、吞错、默认 body、默认 passed 来消除反例。
- 发布失败必须保留上次成果；错误详情写入独立诊断文件，不覆盖已有成功元数据。

## 2. 批次依赖及原计划映射

每个 R 批次可以拆成若干小提交，涉及同一契约的改动应保持可测试的整体。下表是执行顺序，不要求多代理并行。

| 批次 | 内容 | 前置 | 原批次/问题 | 完成标志 |
| --- | --- | --- | --- | --- |
| R0 | 固定反例、能力状态与证据 | 无 | B7/B10、F10 | 已确认缺陷都有稳定失败测试和追踪 ID |
| R1 | 先封住错误发布与虚假 passed | R0 | B7、F02 | 已验证：失败注入拒绝发布且旧成果/元数据保持不变 |
| R2 | 补齐 Schema、路径来源与依赖 | R0 | B1/B6、F09 | 干净环境与跨目录引用测试通过 |
| R3 | 冻结 BuildPlan，消费映射与策略 | R1/R2 | B2/B5、F03 | 有效映射生效，过期映射提前失败 |
| R4 | 统一样式、页面与多来源上下文 | R3 | B3/B4、F03/F04 | 已验证离线：preserve/mixed、页面规格、生成样式和版心执行接通；真实 Word/PDF 矩阵仍待环境补测 |
| R5 | 原始选区与 OPC 无损导入 | R3/R4 | B4/B8、F01/F07 | 已验证离线：选区、关系闭包、ID 映射和失败回退已封堵；复杂 story 与完整交付范围仍待后续批次 |
| R6 | 按交付范围验证内容和有效格式 | R4/R5 | B2/B7/B8、F02/F08/F09 | 已验证：未测、失败、允许变化可准确区分；真实 Word/PDF 由 R8/R10 补齐 |
| R7 | 完成样例校正、预览与复用流程 | R2/R3/R4/R6 | B6/B7、F09 | `verified`：6 项 R7 回归及 20 份独立匿名评估通过，完整样例→决策→格式包/RoleMap→另一目标闭环 |
| R8 | 分节、序列、页眉页脚和物理页面 QA | R5/R6 | B9、F05、W05 | `verified`：54 项离线回归与 W05 真实 Word/PDF 通过 |
| R9 | 脚注/尾注和字段依赖更新 | R5/R6/R8 | B10、F06/F08、W06 | `verified`：15 项字段/注释回归与 W06 真实 Word/PDF 通过 |
| R10 | 完整 Word 回归、PDF 一致性与故障证据 | R7/R8/R9 | B7/B9/B10、F10 | `verified` 当前范围：246 项离线全套、M1 3/3、W04–W07 4/4、M2 装配级 1/1 及完整匿名论文真实 Word/PDF 1/1 |
| R11 | 收尾、示例、迁移和能力发布 | R10 | B1–B10 | `verified`：文档、能力表、示例、迁移输出、实际 CLI 行为与逐交付物元数据一致 |

执行状态统一记录为 `todo / in_progress / blocked / verified`。`verified` 必须附受检代码哈希、命令、断言结果与相关证据，不以“文件已新增”或“测试总数增加”代替。

## 3. R0：把反例固化为正式回归

**改动位置**：`tests/fixtures/formatting/`、现有相关 test 文件；新增 `tests/test_formatting_pipeline_contracts.py`、`tests/test_publication_gates.py`；`docs/formatting-progress.md`、`docs/formatting-capabilities.md`。

1. 从 `docs/acceptance/reproduce_b1_b10.py` 拆出夹具和期望断言，不直接把“错误行为为 true”的观测写成通过条件。可替换 Word 的测试须在名称/说明中标为 integration without Word。
2. 固定 F01–F10 与 W04–W06 的最小样本，增加输入哈希和预期契约 JSON。人工构造的页码记录只验证校验器，不能计入 Word 验收。
3. 将 `test_thesis_acceptance.py` 中现有测试定位为装配级测试；建立独立真实 Word 测试组，不破坏其已有脚注/公式回归价值。
4. 调整进度与能力表，增加“模块实现”和“端到端验证”两列；要求高风险能力有证据才标 supported。规划、未接入、preserve_only 和 unsupported 分开。

**退出条件**：CI 的期望状态明确；已知缺陷回归能复现预期失败，不能长期用无理由 skip 掩盖。独立 Word 任务环境不可用时结果为 blocked，不计作功能通过。

**R0 执行结果（2026-09-05）**：

- 新增 `tests/fixtures/formatting/r0_contracts.json`，固定 F01–F10、W04–W06 的契约、测试层、预期状态和输入哈希；W04–W06 复用的 DOCX 哈希与现有验收材料一致。
- 新增 `tests/test_formatting_pipeline_contracts.py` 与 `tests/test_publication_gates.py`。F01–F09 当前稳定呈现为 15 个 `expected failure`，F10 已验证装配级与真实 Word 测试层分离；新增元数据写失败事务测试通过。
- 新增 `tests/test_word_formatting.py`，真实 Word 环境必须显式设置 `DOCUMENT_SYNTHESIS_WORD_TEST=1`；当前环境未启用，因此记录为 `BLOCKED`，没有计入功能通过。
- `python3 -m unittest discover -s tests -v`：194 项运行，15 个预期失败、1 个 Word 组 BLOCKED、1 个既有条件跳过，0 个非预期失败、0 个错误。
- `git diff --check`：通过。详细命令、受检文件哈希和断言结果见 [`r0-status.json`](acceptance/r0-status.json)。

## 4. R1：先修发布门禁

**改动位置**：`lib/engine.py:540–578`、`lib/delivery.py`、`lib/content_integrity.py`、元数据和发布回滚测试。

当前发现字号不符只写 warnings，且未执行检查也记 passed。先进行最小封堵，再由 R6 增加完整覆盖。

1. 将检查初始状态改成 `not_run`，每份交付物分别记录状态；保留现有元数据兼容字段，但总体 passed 只能由全部必需检查通过计算。
2. `verify_delivery_format().passed == False` 转为 `FormatVerificationError`；执行不到必需检查时抛有代码的诊断，不允许默认成功。
3. 根据解析后的 part.kind 判断交付内容，不使用字符串 `body` 作为唯一条件。对于尚未实现的选区级完整性验证，明确阻止该能力发布，不能拿整篇检查或默认 passed 顶替。
4. 将内容、格式、字段和页级校验统一汇总为交付级报告。只有全部满足才能调用 `_publish_deliveries()`。
5. 失败日志列出 delivery、源节点、属性路径、预期和实测；未验证的非必需属性单独列出，不能和“已检测到不符合”混为 warnings。

**建议接口**（目标接口，可在现有类上扩展）：

```python
@dataclass(frozen=True)
class CheckResult:
    name: str
    status: str  # not_run / passed / failed / unsupported
    required: bool
    diagnostics: tuple

def assert_publishable(delivery_reports) -> None:
    # required 检查必须全部 passed；否则抛 BuildError。
    ...
```

**必测**：12→70 pt、页面尺寸错误、多文件缺图、自定义部分未验证；第二份成果/元数据写入失败。断言退出失败、旧 DOCX/元数据哈希不变、无部分成功发布。W04 在修复前应成为发布门禁的负向测试。

**R1 执行结果（2026-09-05）**：

- 在 `lib/delivery.py` 增加 `CheckResult`、`DeliveryReport`、`assert_publishable()` 和稳定的 QA 汇总；必需检查的 `not_run`、`unsupported`、`failed` 均不能发布，逐交付物保留 source nodes、诊断和未验证属性。
- `lib/engine.py` 对每个声明的 content 部件执行 DOCX 源清单完整性核验，对每个交付物执行有效格式核验；多文件源逐一检查，缺少可核验源节点时明确阻断。发布前统一聚合门禁，失败写入独立 `build-diagnostics.json`，不替换旧交付物或成功元数据。
- `lib/content_integrity.py` 将 `roles.style` 解析为实际样式 ID（F09 的别名子例已验证；R6 仍需扩展交付范围覆盖），并把最终页面尺寸与边距纳入格式核验；restyle/target 模式会将格式包页面几何落到正文。
- `tests/test_publication_gates.py` 已覆盖必需状态阻断、12→70 pt、Letter/A4 页面错误、自定义 content 部件漏图、元数据写失败事务和结构化诊断；F02 相关反例已从 expected failure 转为正常通过。
- `python3 -m unittest discover -s tests -v`：195 项运行，0 个非预期失败、0 个错误、2 个条件跳过（含 1 个未启用真实 Word 的 Word 测试组）、10 个后续批次 expected failure；`python3 synthesize.py --project examples/minimal-demo --plan` 和 `git diff --check` 通过。`--doctor` 已执行并明确报告当前 shell 无法连接 macOS 图形自动化服务，因此真实 Word 环境记为 `blocked`，不是功能通过。详细哈希与命令见 [`r1-status.json`](acceptance/r1-status.json)。

**R1 遗留**：真实 Word/PDF 的 W04 仍需在显式启用本机 Word 的环境运行；完整脚注、字段、选区范围和页码标签覆盖按 R5/R6/R8/R9 扩展，不在 R1 伪装成已完成。

## 5. R2：补齐配置与文件契约

**改动位置**：`schemas/`、`lib/config.py`、`lib/format_schema.py`、`lib/format_resolver.py`、`lib/format_review.py`、`requirements.txt`。

1. 添加 `analysis-v1.schema.json`、`decisions-v1.schema.json`、`role-map-v1.schema.json`；同步 Python 校验。校验器不只验证字段形状，还验证引用、版本、枚举、候选 ID、源哈希和必需决策。
2. 明确 RoleMap 格式为 `mapping_schema_version + source_sha256 + assignments`，每项携带完整 NodeRef、role、来源和必要的确认标志。Schema、compile_role_mapping、RoleMapper 与 CLI 只维护这一种序列化格式。
3. 统一结构类型与语义角色：`table` 是结构类型，`table.body` 是格式角色；bibliography、toc.* 与正文/标题共享同一角色注册表，禁止各模块各自维护不一致集合。自定义样式 ID 通过 roles.style 引用，不能要求与角色同名。
4. `deep_merge_with_provenance()` 与配置覆写共用来源记录，至少保存每个路径字段由哪个文件声明。format.ref、role_map、模板和 source.file 相对各自声明文件解析。
5. 固定 null/false/0、空列表、缺省值、未知字段、循环继承与无效几何的行为，错误位置使用 JSON Pointer。
6. 将 jsonschema 加入受支持依赖；在选定最低 Python 版本的干净环境验证安装。版本范围由实际兼容测试确定，不直接宣称支持未经测试的旧 Python。

**必测**：跨目录 manifest/override/ref 链；修改包但不改版本的哈希检测；错误候选/角色/版本；格式包引用非同名样式；干净虚拟环境完成 analyze/compile/plan。配置失败在写缓存和调用 Word 前发生。

**R2 执行结果（2026-09-05）**：

- 新增 `analysis-v1.schema.json`、`decisions-v1.schema.json`、`role-map-v1.schema.json` 与共享 `lib/contracts.py`；报告、决策和映射均有版本、64 位源哈希、枚举和 JSON Pointer 错误定位。旧决策文件只在输入边界兼容读取，新输出统一为 v1。
- `RoleMap` 已统一为 `mapping_schema_version + source_sha256 + assignments`；每项包含完整 NodeRef（源哈希、部件 URI、元素路径、文本哈希）、角色、provenance 和 confirmed。`RoleMapper`、编译器和 CLI 共用同一序列化契约，表格结构统一映射为 `table.body`。
- 共享语义角色注册表覆盖正文、标题、`table.body`、`bibliography`、`toc.*`、页眉和页脚；格式解析仍允许角色引用不同名的样式 ID，并拒绝未知角色、未知候选和过期节点。
- 配置覆写已保留每个值路径的声明文件来源；`format.ref`、`role_map`、模板和 `source.file` 按各自声明文件解析。显式 `null`、`false`、`0`、空数组、未知字段、继承循环和无效几何均在解析阶段处理。
- CLI 分析现在同时生成 `analysis.json`、`review.html` 和 `decisions.example.json`。隔离 Python 3.14 虚拟环境安装 `requirements.txt` 成功（含 `jsonschema 4.26.0`），并完成 analyze → compile-format → compile-mapping → plan 闭环。
- `tests/test_r2_contracts.py` 5 项通过；全量 `python3 -m unittest discover -s tests -v`：200 项运行，0 个非预期失败、0 个错误、2 个条件跳过、10 个后续批次 expected failure；`minimal-demo --plan` 与 `git diff --check` 通过。受检文件哈希和命令见 [`r2-status.json`](acceptance/r2-status.json)。

**R2 遗留**：真实 Word/PDF 仍按 W04–W06 的环境门禁记录为 blocked；显式 role_map 的生产消费和 BuildPlan 冻结转入 R3，完整交付范围的格式/内容核验转入 R6。

## 6. R3：共用只读 BuildPlan，消费显式映射

**改动位置**：`lib/engine.py` 的 `plan()/synthesize()`、`lib/source_strategies.py`、`lib/role_mapper.py`、`lib/docx_inspector.py`。可新增 `lib/build_plan.py` 承载类型，避免继续扩大 engine。

当前 plan 和 synthesize 各自推导，生产路径无参构造 RoleMapper。目标是一次准备、多阶段使用：

```python
def prepare_build(source, config) -> PreparedBuild: ...  # 只读
def verify_inputs_unchanged(plan: PreparedBuild) -> None: ...
def render_body(plan: PreparedBuild, run_dir) -> RenderedBody: ...
```

PreparedBuild 至少包含：解析配置及 provenance、输入/格式/映射哈希、来源顺序、原始节点索引、选区计划、角色分配、文档部分/分节、交付声明和能力诊断。对外 plan 输出可序列化且不包含源正文；运行时文档句柄不进入序列化结构。

1. 先用 inspector 建立原始 NodeRef；载入 RoleMap 并校验源哈希、story、element_path、text_hash。不要仅靠段落序号或标题文字重找节点。
2. 显式映射优先于样式/大纲推断；`on_unmapped=error` 必须阻止未知节点；preserve 作为明确策略，不能自动全部转 body。
3. 标题导航和目录只读取最终 assignments；同名标题的书签保持唯一、稳定，并记录导入后的节点/书签映射。
4. 输入哈希在准备后、渲染前核验；映射在材料改动后失效。不同来源的节点身份必须包含源文件哈希，不能只靠相同 XPath。
5. 将 `config.formatting` 明确传入执行上下文，不再从 ResolvedFormat 上读取不存在的 policy。

**必测**：有效 heading.2 显式映射改变样式与导航；过期映射在创建 run_dir 前拒绝；同名标题、表格内段落、源目录标题与正文重复标题不串绑；plan 前后目录快照不变；渲染前篡改输入被拒绝。

**R3 执行结果（2026-09-05）**：

- 新增 `lib/build_plan.py`，`PreparedBuild` 统一保存配置 provenance、输入/格式/RoleMap 哈希、来源顺序、NodeRef 索引、选区计划、角色分配、分节、交付声明、诊断和运行时配置；`RenderedBody` 只返回正文文件与节点元数据，不保存文档句柄。
- `UnifiedSynthesizer.prepare_build()` 成为 `plan()` 与 `synthesize()` 的共同准备入口；构建在 `.work/run-*` 创建前校验源文件、RoleMap 和继承格式包哈希；`render_body()` 再次校验后才打开/写入正文。
- `RoleMap` 生产消费已按 source SHA-256、part URI、element path、story type 和 text hash 绑定；显式映射优先于自动推断，`on_unmapped=error` 拒绝缺失节点，`preserve` 只保留未映射节点的原状；表格内路径可直接写入唯一书签。
- `plan()` 的公开结果改为 JSON 可序列化结构并剥离源正文，仅保留必要结构元数据、哈希和角色绑定；smoke 检查通过适配层消费序列化部件注册表。
- 新增 `tests/test_r3_build_plan.py`，覆盖共享准备、正文不泄露、heading.2 样式/导航、重复标题、过期 RoleMap、run 目录前拒绝与渲染前源篡改；F03 由 expected failure 转为正常通过。
- `python3 -m unittest discover -s tests -v`：203 项运行，0 个非预期失败、0 个错误、2 个条件跳过、9 个后续批次 expected failure；R3 定向测试 3 项全部通过。`git diff --check` 与 `minimal-demo --plan` 通过。

**R3 遗留**：真实 Word/PDF 的 preserve/mixed、多来源格式执行和完整交付范围验证仍按 R4–R6 处理；表格/页眉等非正文 story 的真实 Word 导航证据及复杂选区关系闭包仍不提前宣称完成。

## 7. R4：统一样式与页面执行路径

**改动位置**：`lib/layout.py:RenderContext`、`lib/engine.py:render_tree_node_recursive`、`lib/renderers.py`、`lib/style_applier.py`、`lib/styles.py`、`lib/composition.py`、`lib/highlighted_docx.py`。

1. v3 的单文件、多文件、目录与生成部分都接受同一个 RenderContext；将 fonts 字典和 LAST_RENDERED_LANDSCAPE 限定在旧适配路径。递归调用、PDF/图片导入和转横版恢复均传递上下文。
2. preserve：不清洗源样式或直接格式、不改页眉页脚；必要的 ID 重映射允许发生，但有效格式及语义保持。page_policy=source 使用源节几何。
3. restyle：只清理目标托管属性；mixed：先按 NodeRef scope 计算托管属性掩码，保护区不得进入旧全局标准化。保留粗斜体、上下标、超链接、公式等行内语义，除非有明确覆盖配置。
4. target 页面通过统一的 `apply_section_spec()` 应用纸张、边距、方向、页眉页脚距离和网格。不能只把 PageSpec 存进 config。全部节及后续恢复节都验证目标值。
5. 生成标题用 add_styled_heading，目录用 add_styled_toc_entry，页脚走样式与页码策略；移除 v3 目录 22 pt 和公文 `- n -` 的隐式默认覆盖。
6. 原生编号保留及跨来源 ID 重映射；图片尺寸和表格几何从当前版心计算。缓存键包含页面、DPI、裁切及格式/清洗策略。
7. 检查 w:pPr/sectPr 顺序与空段落分节；应用两次之后应保持有效属性和结构幂等。

**必测**：F03 preserve、F04 37 pt 目录来源标题、W04 A4；report/academic/自定义 180×240 mm 三套；横转竖、图片/表格、目录点线和页脚。Word 修改 SynthBody 后相关正文统一变化；完整性清单无非预期变动。

**R4 执行结果（2026-09-05）**：

- `RenderContext` 已贯穿 v3 单 DOCX、多来源递归、DOCX/PDF/图片/PPTX 导入和横竖版恢复；`fonts` 与 `LAST_RENDERED_LANDSCAPE` 仅作为旧适配路径的兼容参数保留。
- 新增统一 `apply_section_spec()`，实际应用目标/源节的纸张尺寸、方向、边距、页眉页脚距离与 `docGrid`；目标规格不再只停留在 config。
- `preserve/source` 跳过源样式清洗、直接格式清理、页脚接管和几何标准化；`mixed` 按文件与 `node_range` 过滤 RoleMap，未托管段落不再隐式降级为 body。
- 生成标题、目录条目和页脚接入格式适配器；目录制表位、PDF 缓存键、图片和表格几何均使用当前版心/格式策略。
- F03 stale RoleMap、F04 多来源 37 pt 标题与 preserve/source 几何页脚回归均通过；新增页面规格、preserve 无变更与 mixed 隔离测试。
- 全量验证：206 项运行，197 passed、7 expected failures、2 skipped、0 非预期失败、0 错误；`git diff --check` 通过。

**R4 遗留**：真实 Word/PDF 的 W04 页面/目录/页脚矩阵仍需在专用环境执行；选区关系闭包、复杂 story 和跨来源脚注按 R5/R9 处理，不提前宣称无损支持。

## 8. R5：正确的选区和关系闭包

**改动位置**：`lib/document_parts.py`、`lib/package_importer.py`、`lib/notes_merger.py`、`lib/composition.py`、`lib/pagination.py`。

1. 选区对原始文档解析，使用 `[start, end)`，选区之间不得重叠；exclude 也计入覆盖。结构节点索引包含 p、tbl、sdt 及可支持对象，图片/公式/字段专用段落均为实质内容。
2. 跨表格、跨字段、跨书签或未支持对象的切分必须提前诊断。不能制造悬空引用后等 Word 修复。
3. slice_document_by_region 改为调用 PackageImporter 导入选中元素及其关系闭包。统一样式、编号、媒体、超链接、书签、脚注/尾注及关联资源的 ID 映射；处理共享部件复用和不同来源 ID 冲突。
4. 删除 `except Exception: pass` 以及校验失败后复制全文的回退。空选区、缺失选区、未知部分都必须有确定行为和错误码。
5. 导入返回 ImportResult：源→目标 NodeRef、bookmark、relationship、style、numbering、note ID 映射。由它支撑后续目录、字段、完整性验证。
6. 输出包校验不仅验证 rId 存在，还验证关系类型、目标 part、内容类型和引用闭合。内部图片关系指向 footer 即使 rId 存在也必须拒绝。

**必测**：F01 图片 rId 误指页脚；无文字图片漏选；坏 p[99]；两个来源同 styleId/numId/footnoteId；跨段书签/字段切分；源目录排除；单独导出摘要、重排附录、多部分组合都无漏失或重复。

**R5 执行结果（2026-09-05）**：

- `SelectionValidator` 固化 `[start, end)`，覆盖 `p`、`tbl`、`sdt` 和无文字图片/公式/字段/脚注引用等实质对象；重叠、空/非法边界、表格内部边界、跨字段/书签边界和未知顶层对象均产生确定错误码。
- `slice_document_by_region()` 改由 `PackageImporter` 执行，目标 `sectPr` 保持末尾合法顺序；样式/`basedOn`、原生编号、媒体、超链接、书签 ID/名称、脚注/尾注引用和关系目标均做映射，并返回不含运行时句柄的 `ImportResult`。
- 选区只合并实际引用的脚注/尾注；脚注内部媒体/XML 部件递归克隆到目标 OPC 包。输出写入前验证关系存在性、关系类型、目标部件和关键内容类型。
- composition 移除非法选区的宽泛吞错和自定义部件复制全文回退；新增 F01/F07、跨边界、未知对象、ImportResult、关系类型/内容类型和选择性脚注回归。
- `python3 -m unittest discover -s tests -v`：217 项运行，213 passed、4 expected failures、2 skipped、0 非预期失败、0 错误；F01/F07 已转为正常回归。
- `git diff --check` 与 `minimal-demo --plan` 通过。详细哈希、命令和剩余边界见 [`r5-status.json`](acceptance/r5-status.json)。

## 9. R6：内容清单与格式验证覆盖实际交付范围

**改动位置**：`lib/content_integrity.py`、`lib/docx_inspector.py:EffectiveStyleEvaluator`、`lib/engine.py`、`lib/delivery.py`、`smoke_test.py`。

1. 每份交付物从 PreparedBuild 计算 ExpectedInventory，包含源选区、来源顺序、允许的生成内容和合法字段/ID 变化；多文件来源合并清单，不按整篇或字面 body 推断。
2. 文本按节点与多重集验证，去掉用任意子串吞掉缺段的宽松匹配；表格验证单元格位置、拓扑和重复数量。媒体/公式检查实例与引用，不仅比较 set 是否含有哈希。
3. 增加脚注/尾注内容、引用、超链接目标、字段指令及依赖清单。允许 ID 重映射但不允许语义变化；受控字段缓存的合法更新不能作为源文字丢失误报。
4. 有效格式验证按 roles→style 解析，并使用原始节点映射定位输出。覆盖页面、页边距、段前后、行距、首行/悬挂缩进、字体、字号、粗斜体、大纲、keep/widow、目录和页眉页脚的托管属性。
5. 有效属性求值补测文档默认→样式继承→段落/字符样式→直接格式、主题字体和显式 false；混合 run 不能用首个 run 代表整段。属性缺失须继续求有效值或报告未验证，不能直接跳过。
6. 每项报告包含 expected/actual/provenance/coverage；托管但无法验证的属性阻止相应支持声明。smoke 与 engine 使用同一契约，防止两个工具互相矛盾。

**必测**：F08 删除脚注/关系；重复公式删一个；表格重复单元格错位；链接目标改写；合法 SEQ 值更新；学习包 styles 使用 style_body 等名称仍被检查；第二个 run 字号错误；混合附件只验证应托管的部分。

**R6 执行结果（2026-09-05）**：

- 新增 `ExpectedInventory` 与 `build_expected_inventory()`，从 `PreparedBuild` 的来源顺序、`layout.parts`/`source_region` 和实际 DOCX 源计算期望清单；多 DOCX 来源按声明顺序合并，清单携带允许的生成文字、替换映射、合法 ID 重映射说明、provenance 与 coverage。
- `verify_content_integrity()` 改为严格段落多重集，不再用任意子串补足缺段；表格按行列、网格宽度、单元格位置和重复次数核验；媒体、公式、脚注/尾注引用与定义、字段指令/依赖、超链接真实目标和书签均进入门禁。媒体与公式按实例计数，rId/noteId 等仅允许语义不变的重映射；字段缓存更新不再被当作源文字丢失。
- `EffectiveStyleEvaluator` 补齐文档默认段落属性、keep/widow、首行/悬挂/左右缩进和主题/显式 toggle 路径；`inspect_docx()` 保留每个 run 的最终有效值。`verify_delivery_format()` 按 role→style 与原始 `NodeRef` 逐项生成 expected/actual/provenance/coverage，混合 run 的第二个字号错误会阻止发布。
- engine 与 smoke 共用同一 `ExpectedInventory`/严格核验契约；交付元数据的单项门禁附带清单和格式证据，内容只在实际内容部件边界内核验，目录/封面等生成部分不被误判为源正文。
- 新增重复媒体、重复公式、脚注删除、链接目标改写、合法字段缓存更新、PreparedBuild 多来源/选区和混合 run 回归；F08 从 expected failure 转为正常通过回归。
- `python3 -W error::FutureWarning -m unittest discover -s tests -v`：225 项运行，222 passed、3 expected failures（F05、F06×2）、2 skipped（真实 Word/分页环境）、0 非预期失败、0 错误。
- `python3 synthesize.py --project examples/minimal-demo --plan`、`python3 -m py_compile ...` 与 `git diff --check` 通过。详细受检文件哈希和剩余边界见 [`r6-status.json`](acceptance/r6-status.json)。

## 10. R7：样例学习、集中校正与预览闭环

**改动位置**：`lib/format_analysis.py`、`lib/format_review.py`、`lib/role_mapper.py`、`synthesize.py`、`examples/custom-format-demo/`。

1. 分析输出保留字段级证据、样本数量、使用中的样式与仅定义样式区别、候选置信度、缺失项及不支持特性。不能用“最大字号”唯一判断标题，也不能把稀有直接格式自动当成规则。
2. HTML 校正界面展示角色→候选→属性；允许修改候选/数值、选择缺失角色继承、确认目标节点角色。导出的 decisions 绑定 report_id/source_sha256，编译时拒绝未知或未决候选。
3. compile-format 接受符合决策 Schema 的明确选择；不要把原始 analysis 原样传回视为完成用户校正。非交互模式允许明确的继承/覆盖决策，不能暗中替用户确认。
4. 增加计划中的 `preview-format`、`render-preview`、`migrate-manifest`；预览默认使用虚构文本，覆盖标题1–9、正文、列表、表格、图片占位、目录及页脚；近似预览标明引擎，Word 预览记录环境。
5. 本地 HTML 用于集中校正和下载 JSON，不额外引入后端服务；所有错误回到具体字段/节点。界面应说明“识别来源格式”和“映射目标内容”是两步。
6. 保存包删除分析缓存后，应用于另一份目标仍能复用；包中不得含样例正文。导出角色映射后必须实际被 R3 消费。

**必测**：全部 Normal 的手工格式文档、空白模板、缺失层级、重复标题、正文内编号、混合中英字体；HTML 转义、哈希错配、未知候选、未决项、输出冲突；另一个目标稿件得到预期有效格式。

**识别评估**：维护独立于开发夹具的匿名评估集，至少覆盖公文、报告、论文及手工格式四类，每类不少于 5 份，合计不少于 100 个标题正例和 500 个正文负例。报告各级 precision/recall、自动接受覆盖率、待确认率和人工校正次数；样本量仍小时明确区间和局限。继承原计划的自动接受标题精确率 ≥98% 目标，同时报告覆盖率，不能只接受极少样本来制造高准确率。未达到时收缩自动接受范围，保留校正流程。

**R7 执行结果（2026-09-05）**：

- `lib/format_analysis.py` 输出样本计数、实际使用/仅定义样式、字段证据、不支持特性和候选置信度；`lib/format_review.py` 的 HTML 决策允许候选/数值校正，`pending` 角色不能编译。
- 新增 `lib/preview.py` 与 `lib/manifest_migration.py`，分别提供虚构内容近似预览和非破坏 v1/v2 manifest 迁移；CLI 接入 `--preview-format`、`--render-preview`、`--migrate-manifest`。
- 格式包删除分析缓存后仍可复用；`tests.test_r7_workflows` 6 项通过，包含另一份目标对格式包和 RoleMap 的实际消费。
- 独立匿名评估集已实际生成并评估：公文、报告、论文、手工格式各 5 份，共 20 份；100 个标题正例、500 个正文负例。标题各级 precision/recall 均为 1.0，正文误报为 0，自动接受覆盖率/精确率均为 1.0，候选复核率为 28.57%，待决率为 12.5%，人工校正次数为 0。完整指标见 [`r7-evaluation/metrics.json`](acceptance/r7-evaluation/metrics.json)。
- 受限范围：评估集为程序生成的虚构中文样例，不能外推到 OCR/PDF 学习、双栏期刊、参考文献著录转换、多样例融合或全部学校规范。

## 11. R8：分节、页码与真实页面验证

**改动位置**：`lib/pagination_types.py`、`lib/pagination.py`、`lib/composition.py`、`lib/layout.py`、`lib/delivery.py`、`lib/qa.py`。

1. 统一测量契约：保留旧 physical_page/printed_page 的兼容转换层，v3 内部使用 PageRecord，携带 section_id、sequence_id、number_value、expected_label、observed_label、verification_status。未测量值为 None，不能伪造标签。
2. Word 测量采集节点所属节及物理页、调整页码；预期标签依据配置计算，实测标签只来自 PDF/实际页脚。标签格式按所属节推导，禁止所有标题套第一 content 部分。
3. 分别定义“重新编号”“继续编号”“隐藏但计数”“无页码”以及物理页奇偶；若要求装订右页起章，必须按物理页验证，不能只设置 oddPage 或检查逻辑编号奇偶。
4. 首页/奇偶页眉页脚有独立部件和继承规则；preserve 不覆盖，managed 按对应 HeaderFooterSpec 写入。不能把同一旧页脚复制到所有变体。
5. PDF 标签提取使用配置/实测版心与页眉页脚区域；匹配文本、装饰字符、罗马/阿拉伯/字母及无显示状态。模糊或未找到返回 unverified，并触发必要检查失败。
6. 校验开始值、连续、重启、节点实际归属、目录标签和内部跳转；F05 的 `label_verified=false` 必须阻止发布。
7. 空白页只允许由明确边界产生的单页留白，结合前后部分、section 和物理位置判定；禁止把两个部分之间全部偶数页豁免。
8. W05 专项诊断：先验证最终段落 widow/keep、分页控制、纸张/边距，再核对“少于几行”是产品需求还是误判。保留现有拒绝结果；任何规则调整需有适用范围及正负样例，禁止为了单例放宽全局门槛。

**必测**：i→1、I→1、字母序列、多个 restart、隐藏页码、横版继续、同名标题、拆分交付、右页起章、错误标签/多余空白页。既有重复标题真实 Word 测试保持通过；新增样本不能手填页码代替实测。

**R8 执行结果（2026-09-05）**：

- 新增类型化 `PageRecord`，交付门禁使用同次 Word 导出 PDF 的实际页脚标签填充 `observed_label`，缺失/不匹配标签不能发布；`page_records` 写入构建元数据。
- 修复 imported section 的 OOXML `w:type` 位置与 Word `oddPage` 实际行为；对实测偶数边界采用局部双分页符校正，使下一部件真实落在物理奇数页，空白页仅进入对应白名单。
- `tests.test_page_sequences tests.test_document_parts tests.test_delivery_pipeline tests.test_formatting_pipeline_contracts`：53 项通过；W05 真实 Word/PDF 通过，验证前置 `i`、正文 `1`、oddPage 页级 QA 与错误发布拦截。

## 12. R9：脚注、题注和字段依赖

**改动位置**：`lib/notes_merger.py`、`lib/field_updater.py`、`lib/package_importer.py`、`lib/delivery.py`、`lib/composition.py`。

1. 先修 W06：不要把标签字典作为 PageRecord 传递。目录显示标签与字段测量上下文是两个显式类型；FieldUpdater 接收统一页面索引，类型不符合时给结构化错误，不能在 `.get()` 处崩溃。
2. 修正 SEQ Roman/Alphabetic 参数及大小写开关；只接受声明支持的字段/开关。区分简单字段、复杂字段及嵌套字段，按文档真实顺序建立 FieldIndex。
3. 用 bookmarkStart/bookmarkEnd 精确提取范围，支持同段与跨段；不能取整段。建立依赖图，按 SEQ→REF→分页相关字段的顺序求值；前向引用、嵌套依赖和循环均有诊断。
4. 受控字段与保留字段分开：preserve 不擅自重算未知字段；外部链接、未知指令在不支持变换前明确拒绝或按能力表保留。
5. 脚注/尾注合并导入正文和部件中的关系闭包；唯一 ID、separator/continuation、样式及跨来源冲突都覆盖。不得只统计 notes.xml 内节点数量。
6. 字段和目录共享稳定的测量迭代：初次装配→非分页字段更新→Word 导出测量→更新目录/PAGEREF→再次导出；直到页码、标签、字段结果和节点物理页共同收敛。无目录但有 PAGEREF 也必须进入迭代。
7. 收敛上限沿用可配置的有界次数；未收敛时给出每轮差异并拒绝发布。确保最终保存的 DOCX 字段缓存与最后 PDF 一致，不能只让 Word 私有副本更新正确。

**必测**：ROMAN/roman/ALPHABETIC/alphabetic、SEQ 重置/复用、仅序号书签、跨段 REF、前向 REF、带/不带目录的 PAGEREF、未知/循环字段；两来源同 footnoteId、注释含图/链接、OMML 保留。W06 必须完成测量收敛并正确发布，且字典/字符串混用有单独负向测试。

**R9 执行结果（2026-09-05）**：

- `lib/field_updater.py` 已统一处理简单/复杂字段，支持 SEQ 的 ROMAN/ALPHABETIC/ARABIC、大小写、`\r` 重置/继续，REF 精确读取 bookmarkStart/bookmarkEnd 范围，并建立有环检测的依赖图。
- PAGEREF 只接受结构化 `PageRecord`/记录字典，字符串页码会给结构化类型错误；复杂前向 REF、未知字段缓存和尾注 ID/关系闭包均有回归。
- `tests.test_r9_fields_and_notes tests.test_notes_and_fields`：15 项通过；W06 真实 Word/PDF 通过，最终字段缓存由收敛后的 staging DOCX 保留。

## 13. R10：完整集成验收与 PDF 证据

**改动位置**：新增 `tests/test_word_formatting.py`、扩展 `tests/test_m1_acceptance.py`、`tests/test_thesis_acceptance.py`、`tests/fixtures/formatting/expected/`、诊断/元数据代码。

### 13.1 真正的 M1/M2 测试矩阵

| 测试组 | 必须覆盖 | 通过条件 |
| --- | --- | --- |
| v1/v2 回归 | 匿名完整/拆分文档、旧预设 | 交付清单、源内容、目录链接及 Word 基线无未解释变化 |
| M1 格式执行 | academic/report/自定义纸张，同一稿件与多文件 | 最终 DOCX 有效属性与 PDF 页面尺寸满足目标；W04 不再错误发布 |
| M1 学习闭环 | 样例→人工决策→包→另一个目标→映射→Word | 映射被消费；包无样例正文；角色、导航、目录和托管格式全部符合 |
| M1 保留边界 | preserve/mixed、行内强调、已有对象 | 保护区语义和有效格式不变；托管区确定改变 |
| M2 完整匿名论文 | 封面、声明、中英摘要、目录、两章、参考文献、附录 | 源选区完整且无重复；按实际角色/部分验证 |
| M2 页面 | 前置罗马、正文重启、首页/奇偶页、横版、右页起章 | 每页、每节和每个目录节点均可解释；必要标签零不匹配 |
| M2 对象/字段 | 图片、表格、原生列表、脚注/尾注、OMML、SEQ/REF/PAGEREF | 对象关系完整；引用值正确；最终 DOCX/PDF 对应同一收敛版本 |
| 故障与事务 | 坏格式、陈旧映射、缺对象、Word 失败、不收敛、发布失败 | 不改源文件；不替换旧成果；错误诊断完整 |

完整论文夹具从 CLI 调用构建，禁止手工合并正文、填写 pages 或跳过发布门禁后宣称端到端通过。Word 自动化顺序执行，避免多个测试争夺 active document。

### 13.2 PDF 与视觉诊断

1. 保存源哈希、格式包/配置哈希、最终 DOCX/PDF 哈希、节点页码映射、字体/Word/Python/平台版本和每项 QA 结果。
2. 正常流水线 PDF 出现对象偏移警告时保留原文件和解析日志，核对导出完成时机及结构；不能通过静默“修复 PDF”覆盖原证据。
3. 用标准渲染器查看所有关键页面；W05 的 Poppler/MuPDF 差异单列诊断，用相同 PDF 复核并明确哪种能力受影响。未解释差异不得记为“全部视觉通过”。
4. 对纸张尺寸采用 OOXML 整数单位允许的换算误差（例如不超过 0.1 mm）；字段/页码要求精确一致。字体环境不一致应记录缺失与替代，不能用宽松误差掩盖。
5. 正式门禁断言对象、格式、页码和发布事务；视觉抽查补充裁切、重叠、目录点线和图表布局。不能仅靠截图相似度代替语义校验。

### 13.3 执行命令与退出条件

每批先运行受影响的文件测试；基础代码修复后运行全套：

```bash
python3 -m unittest discover -s tests -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p 'test_word_formatting.py' -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p 'test_m1_acceptance.py' -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p 'test_thesis_acceptance.py' -v
git diff --check
```

`test_word_formatting.py` 是本计划新增的独立真实 Word 测试组，已存在并完成 W04–W07 验收。测试计数不作为固定目标；必需用例没有 skip、所有断言通过、所有已确认问题关闭才可通过。Word 条件不满足时标明 blocked，不能将离线替换的结果算入该门槛。

### 13.4 R10 当前执行结果（2026-09-05）

- 代码与测试已按当前工作区执行：`python3 -W error::FutureWarning -m unittest discover -s tests -q` 为 **246 项 OK，2 项条件跳过**；无非预期失败、错误或 FutureWarning。
- 真实 Word：`DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest tests.test_word_formatting -v` 为 **W04/W05/W06/W07 4/4 OK**；同环境 `tests.test_m1_acceptance` 为 **3/3 OK**。W04 验证 A4 DOCX/PDF，W05 验证实际罗马/阿拉伯标签、oddPage 和页级 QA，W06 验证 PAGEREF 收敛，W07 验证完整匿名论文的真实 Word/PDF 交付门禁。
- M2：`python3 -m unittest tests.test_thesis_acceptance -q` 为 **1/1 OK**，并以 W07 真实 Word/PDF 矩阵覆盖封面、声明、中英摘要、目录、两章、参考文献、附录、脚注/尾注、图片、表格、原生列表、OMML 与 SEQ/REF/PAGEREF；装配级测试仍保留其独立回归价值。
- 元数据已记录 DOCX/PDF/configuration/override 哈希、PageRecord、QA 摘要、Word/Python/平台信息；正式证据见 [`r10-status.json`](acceptance/r10-status.json)。
- R10 当前范围关闭 F06/W06 与 F05/W05 的生产回归；R11 负责的文档、示例、能力声明最终同步已完成，不在本节提前宣称全部学校规范支持。

## 14. R11：文档、示例与发布收尾

**改动位置**：`README.md`、`DEVELOPMENT.md`、`formats/README.md`、`docs/formatting-progress.md`、`docs/formatting-capabilities.md`、原执行计划、示例 manifest/决策文件、项目 skill、`docs/acceptance/r11-status.json`。

1. 每个能力声明对应测试 ID、支持级别和最近验收证据；同步移除“175 项全绿即全部达成”等过度结论。
2. 示例命令在干净目录执行，包括分析/校正、编译、映射、预览、plan、doctor、构建和 smoke。不存在的命令不能提前写成可用。
3. 给迁移输出独立文件，v1/v2 默认行为不变；新增格式包与项目清单带版本及迁移说明。
4. README 明确当前覆盖的论文能力和不支持项：本轮不包含 OCR/PDF 样例学习、双栏期刊、参考文献著录转换、多样例融合或全部学校规范认证。
5. 清理本轮 `.work` 和渲染缓存，只保留正式证据。发布元数据包含完整验证覆盖率，不能只写汇总 passed。

**M1 放行**：R0–R7 完成，R10 中 M1/旧版/事务矩阵通过；M2 未验证能力仍不得宣传 supported。

**M2 放行**：R8/R9 完成，完整论文真实 Word/PDF 矩阵通过，W05/W06 和相关反例关闭，R11 文档同步完成；未验证能力仍不得宣传为 `supported`。

### 14.1 R11 执行结果（2026-09-06）

- 自定义格式示例在 `output/custom-format-demo/` 完成一次干净输出闭环：分析样例、最终决策编译、虚构近似预览、目标分析、绑定源哈希的 RoleMap、v1 清单迁移、`--plan`、`--doctor`、真实 Word build 和完整 `smoke_test.py` 均通过。
- 真实 Word build 交付 1 份 `academic_paper.docx`，逐页断言为 1/1，内容完整性、格式核验、页面几何、PDF 页码记录、发布元数据均通过；完整 smoke 再次完成结构、正文文字、同次 Word/PDF 页码和格式核验。
- 修复 `verify_delivery_format()` 在继承格式包存在同名语义角色时错误使用父样式别名的问题，并由 R11 回归测试固定该行为。
- 迁移输出写入独立 v2 文件，输入 v1 文件哈希未变化；格式包为 `academic-demo@1.0.0`，目标 RoleMap 绑定 `manuscript.docx` 的 SHA-256；预览元数据明确 `fictional_content: true`、`approximate: true`、`source_content_included: false`。
- 发布元数据包含逐交付物 `content_integrity`、`format_verification`、DOCX/PDF 哈希、PageRecord、配置与环境信息；清理完成后不保留 `.work`、分页渲染缓存或 Python 缓存。
- 正式状态与哈希见 [`r11-status.json`](acceptance/r11-status.json)。

## 15. 风险、拆分及下一步

| 风险 | 处理方式 |
| --- | --- |
| 门禁修正后大量旧成功用例转为失败 | 接受真实暴露的失败；按缺陷修复，不恢复默认 passed。保留 v1/v2 兼容边界。 |
| 选区与字段/注释跨界 | 预检拒绝不安全切分；先实现整对象导入，再扩大切分能力。 |
| 样式继承复杂导致隐性覆盖 | 以有效属性和来源证据验证，增加混合 run 与显式 false 的边界测试。 |
| Word/字体环境引起分页差异 | 固定验收环境并记录版本；语义和目标格式不变，分页基线变化需解释。 |
| 字段更新导致再次分页 | 有界迭代并保存每轮差异；无法收敛不发布。 |
| 已有全绿测试掩盖集成缺陷 | 独立契约断言、生产 CLI 与真实 Word 三层共同验收。 |

优先执行 **R0 → R1 → R2 → R3**。第一组可交付结果应是“错误成果不能发布、无效映射提前失败、配置来源可追溯”，然后接通样式与无损装配；在这些基础完成前，不优先扩展更多论文格式选项。

本计划不作固定工期承诺。R0 完成反例拆分后，以每批实际失败数和对象复杂度估算资源；每批结束提交变更文件、测试命令、结果、遗留问题和证据链接，才进入下一批。用户已授权匿名验收目录的 Word 访问，后续同类验证无需重复请求该授权。
