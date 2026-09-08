# R0–R11 独立验收报告

日期：2026-09-06。结论：**不通过整体验收，暂不能将 R0–R11 全部标为 verified，也不能据此放行 M1/M2。**

本轮确实补上了大量实现与回归。重新运行现有测试得到 `Ran 252 tests … OK (skipped=2)`，两次独立生产 CLI 构建也都成功发布。然而，检查实际交付物及新增反例后，仍确认 4 项 P1 缺陷，以及内容完整性、迁移、识别评测和集成覆盖方面的 4 项 P2 问题。其中无目录 PAGEREF 和论文封面错误已经通过本机真实 Word 复现，不能仅调整能力声明后忽略。

本报告是新的独立验收结论；原有 `r*-status.json`、进度文档及历史成功日志仍保留，不覆盖历史记录。本轮未修改生产代码、既有测试、格式包或用户输入，仅新增验收脚本、报告和证据。

## 1. 验收范围与证据边界

受检基线为当前未提交工作区，使用 [source-hashes.json](acceptance/r0-r11-review/source-hashes.json) 固定生产代码、测试、Schema、格式包等文件。验收结束核对这些文件未发生变化。

环境：macOS 27.0 arm64，Python 3.14.0，Microsoft Word 16.112.3（16.112.26083020）。Word Automation 探针成功；本轮沿用用户对匿名验收临时目录的访问授权。

| 验证 | 本轮实际结果 | 可以证明什么 |
| --- | --- | --- |
| 全套现有 unittest | 252 项运行；`OK (skipped=2)` | 已有断言通过；不代表新增反例通过 |
| 原有完整论文生成器 → plan → doctor → 生产 CLI → Word | 发布成功，12 页 | 学术预设页面、基础目录/字段/对象链路能够运行；封面实际错误 |
| 无目录、仅正文 PAGEREF → plan → doctor → 生产 CLI → Word | 发布成功，6 页 | 已定位最终 DOCX 与 PDF 字段结果不一致 |
| 实际 `prepare_build` / `render_body` / 最终格式验证函数 | preserve 和行内强调被拒；移除托管样式后反而通过 | 渲染策略与验收契约不一致、检查覆盖可被绕过 |
| 内容清单与迁移函数反例 | 错序/重复正文通过；迁移覆盖自己的输入 | 模块层缺陷，未冒充完整 Word 构建结果 |
| 现有评估集的去结构标记副本 | 100 个标题正确识别数为 0 | 原有 100% 指标不能外推到无大纲标记的手工格式 |
| 论文 PDF 视觉检查 | 逐页查看全部 12 页 | 发现错误封面；其余页未见明显裁切/重叠，不等于所有 M2 版式已验证 |

两个 skip 分别为 opt-in 分页测试和真实 Word 测试类的 `setUpClass`。本轮另跑上述两条真实 CLI 路径，没有把跳过的整个测试组计入通过数，也没有重新宣称历史 W04–W06/M1 全矩阵通过。

## 2. 已确认的问题

### C01 · P1 · 无目录时不更新最终 DOCX 的 PAGEREF，却允许发布

**关联批次：R1、R9、R10。**

触发条件：v3、`academic-basic`、`restyle`，交付物仅含 `body`，来源包含指向第二章的 PAGEREF。

实际结果：生产 CLI 返回成功；最终 DOCX 字段 `PAGEREF _Toc_0002_b71faf` 的缓存为 **1**，同次测量记录的目标打印页为 **4**，PDF 文字也显示 `The second chapter starts on page 4`。用户拿到的 DOCX 与验收 PDF 并非同一字段结果。

代码定位：`lib/delivery.py:464–473` 仅在 `has_toc` 时构造下一轮 `exact`，无目录则为空字典；`all(...)` 对空集合为真，第一次导出即结束。最后没有逐一核对交付 DOCX 字段缓存。R9 原计划第 6、7 条明确要求无目录 PAGEREF 也迭代，并保证最终 DOCX 缓存与 PDF 一致。

**修复验收条件：** 收敛集合由分页依赖字段和目录共同决定，不能由目录是否存在决定。测量后更新将被发布的 DOCX，再次导出；逐字段核对指令、目标、打印标签及缓存。带目录/无目录、非 1 目标页、罗马标签、超过迭代上限必须有断言；不收敛应拒绝发布并保留旧成果。

证据：[观察记录](acceptance/r0-r11-review/body-pageref-observations.json)、[DOCX](acceptance/r0-r11-review/body-pageref.docx)、[PDF](acceptance/r0-r11-review/body-pageref.pdf)、[构建日志](acceptance/r0-r11-review/body-pageref-build.log)。

### C02 · P1 · v3 论文仍自动套用旧业务封面，配置标题消失

**关联批次：R3、R4、R6、R10、R11。**

使用仓库自己的完整论文夹具，清单明确设置 `cover.main_title = "匿名学位论文样例"`，未显式指定模板。`--plan` 自动选中仓库 `templates/封面+目录.docx`；最终 PDF 第一页是旧业务封面，未出现配置的论文标题。元数据却记录内容及格式核验均为 `passed`。

代码定位：`lib/config.py:778–782` 对所有版本沿用模板自动发现；`lib/composition.py:110–136` 载入模板后只替换 `{{MAIN_TITLE}}` 等占位符，模板中的固定业务文字不会被替换。现有检查没有对生成封面的预期内容进行验证。

**修复验收条件：** 明确 v3 封面来源契约，默认通用封面或显式模板；将旧版自动发现限制在兼容路径。显式模板若缺少所需字段，应给出可定位诊断或要求明确采用固定内容。最终生成封面必须验证配置标题等必需字段。增加“仓库存在旧模板、v3 未指定模板”的 CLI 回归，同时保留 v1/v2 兼容测试。仅给这一个夹具加 `template: false` 不能关闭通用路径缺陷。

证据：[计划日志](acceptance/r0-r11-review/thesis-plan.log)、[观察记录](acceptance/r0-r11-review/thesis-observations.json)、[实际 PDF](acceptance/r0-r11-review/thesis-complete.pdf)。该 PDF 是失败证据，不是合格的匿名论文交付示例。

### C03 · P1 · 格式验收忽略保护策略，拒绝正确的 preserve 与行内强调

**关联批次：R4、R6、R10。**

两条不依赖 Word 的实际渲染反例：

- `preserve + page_policy: source` 正确保留来源 180×240 mm 页面与原始样式；验收仍要求目标 A4/25 mm 边距，报告 6 项偏差且 `verified_count=0`。
- 默认保留行内强调的 restyle 正确保留一小段加粗；验收却要求所有 run 都等于正文角色的 `bold=false`，因此拒绝。

代码定位：`lib/engine.py:786` 调用最终验证时仅传入 DOCX 与 `resolved_format`；`lib/content_integrity.py:753` 的接口没有 FormattingPolicy/保护范围，页面检查约 `869–906` 无条件比较目标 PageSpec，`956–965` 对粗斜体逐 run 强制相等。与 `lib/style_applier.py:495–511` 的强调保留逻辑冲突。

**修复验收条件：** 从 PreparedBuild 传入每节点/每节的托管属性、保护属性及来源期望，验证实际允许变化。preserve 应验证与来源一致；mixed 按范围区分；局部粗斜体仅在策略要求改写时按目标覆盖。新增真实 CLI 的 preserve/source、mixed、局部粗体和斜体通过用例，并配对应保护区被破坏的拒绝用例。

证据：[探针脚本](acceptance/review_r0_r11_probes.py)、[观察记录](acceptance/r0-r11-review/probe-observations.json)。此处结论来自实际渲染与实际验证函数，未声称已跑这两条完整 Word 路径。

### C04 · P1 · 托管段落失去样式标记后，错误格式可以漏检通过

**关联批次：R1、R6。**

先渲染一份正常的 5 段正文，确认格式验证通过；再将第一段样式改为 `Normal`、所有 run 字号改为 **70 pt**。验证仍返回 `passed=true`，只是检查数量从 5 降为 4。

代码定位：`lib/content_integrity.py:931–932` 对无法推断角色的节点直接跳过；`986` 只要求 `verified_count > 0`，没有要求实际覆盖全部应托管节点。`lib/engine.py:786` 也未传入来自原始节点的输出映射。用最终样式名判断“该不该查”会把样式被破坏的节点排除出去。

**修复验收条件：** 使用渲染前冻结的 ExpectedFormatInventory 和输入→输出节点映射确定检查对象；丢失角色、映射不完整应是 `unverified` 并阻止发布。增加 5/5 覆盖断言，以及改回 Normal、改成未知样式、导入前缀变化、删除托管节点的负向测试。不能只提高最少检查段落数量。

证据：[观察记录](acceptance/r0-r11-review/probe-observations.json) 的 `removed_managed_style`。

### C05 · P2 · 内容完整性只检查下限，错序和意外重复都通过

**关联批次：R5、R6、R10。**

源正文为 `[A, B]`，交付正文改成 `[B, A, A]`，实际 `verify_content_integrity()` 仍返回 `true`。这既不能保证来源顺序，也不能发现意外多复制了一段。

代码定位：`lib/content_integrity.py:560–567` 仅检查 `actual_count < required_count`。段落多重集没有顺序；允许生成内容也没有被用于严格区分正常新增与重复来源。R6 要求来源顺序、节点清单及允许生成内容，R10 要求源选区无重复，当前尚未满足。

**修复验收条件：** 按交付部件记录有序来源节点及出现次数；生成封面/目录/标题和合法字段变化使用独立白名单。正文顺序颠倒、来源重复、重复文本但不同 NodeRef、跨来源同文段落都应有独立断言。合法多次引用同一选区应由计划显式声明次数。

证据：[观察记录](acceptance/r0-r11-review/probe-observations.json) 的 `reordered_duplicated_content_accepted`。这是验证器故障注入，未推断正常装配必然产生重复。

### C06 · P2 · “非破坏迁移”允许覆盖自己的源文件

**关联批次：R7、R11。**

对匿名 v1 清单调用 `migrate_manifest_file(path, path, replace=True)`，原文件被直接写成迁移结果。虽然 `replace` 明确允许覆盖已有目标，但按本项目“迁移到独立文件、输入只读”的契约，它不应同时解除源文件保护。

代码定位：`lib/manifest_migration.py:118–130` 已 resolve 两个路径，但写入前没有比较源与输出是否为同一文件。

**修复验收条件：** 对同一路径、相对路径别名、符号链接及已有硬链接检查同一文件；这些情况即使有 `replace` 也拒绝。普通已有独立目标仍可按显式 replace 覆盖。断言异常前后源 SHA-256 不变。

证据：[观察记录](acceptance/r0-r11-review/probe-observations.json) 的 `migration_overwrites_own_input`。本轮仅覆盖新建匿名反例文件，没有覆盖用户清单。

### C07 · P2 · 识别评测未覆盖真正无结构标记的手工格式

**关联批次：R7、R10、R11。**

`docs/acceptance/r7-evaluation/generate_dataset.py` 对四类文档的每一个标题都调用 `_set_direct_outline(heading, level)`，包括“手工格式”。因此 20 份文档、100 个标题的 precision/recall=1.0 很大程度上测的是正确读取现成大纲标签。

本轮复制这些样本，仅将标注标题样式设为 Normal 并移除直接 `w:outlineLvl`；文字、直接字体/字号/粗斜体、段落位置及标签保持不变。沿用现有评估器后，各级标题 recall 均为 0，标题准确率与自动接受覆盖率为 0，100 个标题均计入人工校正。没有标题预测，所以 precision 是未定义 `null`，不是 0。该结果也不代表系统会错误自动接受标题，系统选择了低置信度待确认。

**修复验收条件：** 把结构化样本与无结构样本分开报告；补充字体/字号/间距等组合特征形成候选，并保留人工校正与安全弃权。增加正文内编号、局部粗体、重复标题、缺层级、中英混排、空模板等混淆样本。根据实际覆盖率限定“智能识别”声明；不能要求每个不确定标题强行自动通过，也不能用当前结构化 100% 指标概括手工排版能力。

证据：[变换与评估脚本](acceptance/review_r7_no_outline.py)、[完整指标](acceptance/r0-r11-review/no-outline-metrics.json)。这是现有合成集的受控反事实检查，不是独立真实文档质量估计。

### C08 · P2 · W07 的断言与夹具不足以关闭完整 M2 矩阵

**关联批次：R10、R11。**

当前夹具仅声明 `cover/toc/body` 三个部件，声明、中英文摘要、参考文献和附录都作为一个来源正文中的普通一级标题，没有 `source.regions`。标题为 Abstract 的部分仍由同一段中文生成函数填充。夹具没有横版节、区分内容的奇偶页眉，也没有两来源同 noteId 合并场景。

`tests/test_word_formatting.py:108–151` 主要检查成功退出、罗马/阿拉伯标签存在、对象/字段 XML 存在和两章标题存在；没有验证封面配置标题、逐字段缓存、每个部件的范围/页码归属。因此错误封面也能满足当前测试断言。

**修复验收条件：** 根据原计划 13.1 的矩阵逐行补断言，允许使用多份小而独立的真实 CLI 夹具组合覆盖，无须把所有能力塞进一个超大样例。至少补齐独立前置部件与真实中英内容、无重复选区、正文重启、奇偶页眉、横版与右页起章、多来源注释及带/不带目录字段值；让本报告 C01–C06 的反例成为发布回归。覆盖之外的能力明确标记未验证。

证据：[本轮输入清单](acceptance/r0-r11-review/thesis-fixture/manifest.json)、原生成器 `docs/acceptance/create_complete_thesis_fixture.py`、本轮完整论文输出与元数据。

## 3. R0–R11 逐批结论

“现有回归通过”仅指本轮执行了对应现有测试；“部分确认”不等同于该批全部原定退出条件通过。

| 批次 | 本轮结论 | 剩余退出条件 |
| --- | --- | --- |
| R0 | 现有反例与证据机制已建立 | 纳入 C01–C08，保持状态与反例一致 |
| R1 | 部分确认，不能关闭 | C01/C02 错误发布，C04 覆盖漏检 |
| R2 | 现有契约回归通过 | 本轮未重建干净依赖环境，保留历史证据的环境边界 |
| R3 | 基础计划/映射回归通过 | 最终验证尚未完整消费策略和原始节点映射（C03/C04） |
| R4 | 部分确认 | A4 已在实际 DOCX/PDF 生效；C02/C03 尚未关闭 |
| R5 | 基础选区/导入回归通过 | 不能用本轮单来源样例证明复杂选区/多来源无损；见 C05/C08 |
| R6 | 未通过 | C02–C05，尤其按范围核验及完整覆盖 |
| R7 | 工作流部分确认 | C06 非破坏边界、C07 手工格式识别证据 |
| R8 | 基础实测路径确认 | 本轮论文有 i/正文重启和目标 A4；复杂页面矩阵仍须补齐 |
| R9 | 未通过 | 带目录路径的 PAGEREF 已正确；无目录路径仍失败（C01） |
| R10 | 未通过 | C01/C02 实际失败，C08 完整矩阵不足 |
| R11 | 未通过发布验收 | 关闭阻断项后同步状态、示例与能力表，不能保留无条件全批次 verified |

已确认的进展包括：现有测试无失败；学术预设 A4 实际落到 Word/PDF（12 页均约 209.97×297.01 mm）；论文路径可以生成目录罗马 i、正文阿拉伯序列；同一论文的第二章 PAGEREF 缓存为 8，与实测目标打印页 8 一致；图片、表格、列表、脚注/尾注及 OMML 在该单来源样例中存在。以上结论仅限该夹具与环境。

## 4. 建议补全顺序及再次验收标准

1. **先固定反例与正确失败。** 将 C01–C06 转为独立正式回归；保留正常对照样本。R0 的证据索引增加本报告的问题 ID，R11 状态按实际支持边界修正。
2. **先解决错误发布。** 修复 C01 字段收敛与最终缓存核对、C02 封面来源与必需生成内容、C04 预期节点覆盖。每项同时验证拒绝发布不会替换旧 DOCX/元数据。
3. **统一预期清单。** 在 PreparedBuild/RenderResult 中贯通来源节点、输出节点、顺序、保护/托管属性和生成内容许可，供 engine 与 smoke 共用；关闭 C03/C05，避免两套启发式重复判断。
4. **关闭独立的小边界。** 修复 C06 的同文件保护；补齐 C07 的识别评测分层和困难样本，按实测覆盖决定自动接受范围。
5. **补全 Word 矩阵再收尾。** 逐项完成 C08，并重新运行现有完整离线套件及 opt-in Word/M1/论文测试。对每份交付物保存输入/代码/输出哈希、最终 PDF、字段值和逐项 QA。

再次放行的最低要求：所有已确认 P1 反例转为正常断言通过；C05/C06 关闭；C07/C08 要么满足原定覆盖，要么明确收缩里程碑和能力范围；必需的 Word 用例无 skip；最终 DOCX/PDF/配置一致；文档声明与验收证据一致。单纯增加测试数量或让进程 exit 0 不能替代这些条件。

## 5. 复现命令

从仓库根目录执行。下列 scratch 路径可换成新的匿名目录，输出不得复用用户成果目录。

```bash
python3 -m unittest discover -s tests -v
python3 docs/acceptance/review_r0_r11_probes.py output/r0-r11-review/probes
python3 docs/acceptance/review_r7_no_outline.py output/r0-r11-review/no-outline output/r0-r11-review/no-outline-metrics.json

python3 docs/acceptance/create_complete_thesis_fixture.py output/r0-r11-review/thesis-source
python3 synthesize.py --source output/r0-r11-review/thesis-source --manifest output/r0-r11-review/thesis-source/manifest.json --plan
python3 synthesize.py --source output/r0-r11-review/thesis-source --manifest output/r0-r11-review/thesis-source/manifest.json --doctor
python3 synthesize.py --source output/r0-r11-review/thesis-source --manifest output/r0-r11-review/thesis-source/manifest.json --output-dir output/r0-r11-review/thesis-output --keep-work

python3 synthesize.py --source docs/acceptance/word-followup/m2-fixtures/pageref-source.docx --manifest docs/acceptance/r0-r11-review/body-pageref-manifest.json --plan
python3 synthesize.py --source docs/acceptance/word-followup/m2-fixtures/pageref-source.docx --manifest docs/acceptance/r0-r11-review/body-pageref-manifest.json --doctor
python3 synthesize.py --source docs/acceptance/word-followup/m2-fixtures/pageref-source.docx --manifest docs/acceptance/r0-r11-review/body-pageref-manifest.json --output-dir output/r0-r11-review/body-pageref-output --keep-work
```

探针退出码 0 只表示执行完成，实际反例结果必须读取 observations；真实 CLI 退出码 0 同样不能替代字段/封面断言。完整论文生成器会沿用当前模板发现行为，因此可稳定暴露 C02。

## 6. 证据保存与 PDF 限制

正式证据保存在 [acceptance/r0-r11-review/](acceptance/r0-r11-review/)：测试日志、两份实际交付 DOCX、同次流水线原始 PDF、构建元数据、输入夹具、观察 JSON、评测指标及代码哈希。证据文件哈希另见 `evidence-hashes.json`。临时 `.work`、页面 PNG 和受控变换样本在证据固定后清理，可用以上脚本重建。

两个原始 PDF 在 pypdf 解析时仍出现 `Ignoring wrong pointing object … (offset 0)` 警告；没有静默重写或修复原始 PDF。Poppler 的 `pdfinfo`/`pdftotext` 均 exit 0、无警告，与 pypdf 的 12/6 页计数一致，也独立确认错误封面与 PDF 的 PAGEREF=4。详见 [解析日志](acceptance/r0-r11-review/pdf-parser-check.log)。这些交叉检查足以支持本文具体发现，但未解释 PDF 对象偏移警告的根因，因此本报告不授予“所有解析器兼容/所有 PDF QA 完全通过”的结论。
