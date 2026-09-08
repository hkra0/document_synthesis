# N0–N10 独立复验报告

日期：2026-09-06。结论：**部分修复已确认，整体验收仍不通过。**

本轮重新运行现有离线测试，结果为 **286 项，OK（skipped=3）**；随后在本机真实 Word 环境重新运行 NW01–NW12，结果为 **12/12 通过，95.991 秒，无跳过**。新增独立反例仍确认 **2 项 P1、4 项 P2 问题**。现有成功测试不足以支持“优化执行计划全部完成”的结论。

本轮未修改生产代码、既有测试或历史状态文件。新增的复验脚本、报告及证据保存在本报告对应目录。

## 1. 实际执行与证据边界

| 检查 | 本轮结果 | 证据 |
| --- | --- | --- |
| 完整现有离线套件 | `Ran 286 tests in 47.930s`，`OK (skipped=3)` | [unit-tests.log](acceptance/n0-n10-review/unit-tests.log) |
| 真实 Word NW01–NW12 | `Ran 12 tests in 95.991s`，`OK` | [word-tests.log](acceptance/n0-n10-review/word-tests.log) |
| 原无目录 PAGEREF 的独立 CLI 重建 | 成功发布；最终 DOCX、PDF 和实测目标页均为 **4** | [观察记录](acceptance/n0-n10-review/pageref-observations.json)、[DOCX](acceptance/n0-n10-review/pageref.docx)、[PDF](acceptance/n0-n10-review/pageref.pdf) |
| 合法 preserve 混合字号的独立 CLI | Word 成功导出且页级 QA 通过；最终格式门禁错误拒绝，exit 1 | [构建日志](acceptance/n0-n10-review/preserve-build.log)、[拒绝时的 DOCX](acceptance/n0-n10-review/preserve-rejected.docx)、[PDF](acceptance/n0-n10-review/preserve-rejected.pdf) |
| 实际装配/实际验证器反例 | 合法混合字号拒绝；损坏字号通过；页眉页脚及距离损坏通过 | [probe-observations.json](acceptance/n0-n10-review/probe-observations.json) |
| 内容门禁反例 | 来源 1 张图、交付 2 张图，内容验证仍通过 | 同上 |
| 缺层级标题 + 现有评估器 | 三级标题被选为二级，分数 0.985；指标仍报告自动精确率 1.0 | 同上 |
| 现有无大纲冻结集重评 | 原有集仍取得其报告的高指标；不能覆盖新增反例 | [unstructured-metrics.json](acceptance/n0-n10-review/unstructured-metrics.json) |
| 新 PAGEREF PDF 警告诊断 | 重现已知 orphan-xref 模式；三个解析器页数/页面框一致 | [诊断报告](acceptance/n0-n10-review/pdf-diagnostics/report.json) |

离线套件中的 3 项条件跳过包括 opt-in 分页测试和两个 Word 测试类；本轮另行执行了新 Word 矩阵。未重新运行所有历史 opt-in M1/W04–W07 测试组，不能将这些历史组算作本轮独立通过。

Word 使用 `DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access`，沿用用户已有匿名验收目录授权。本轮没有因权限或弹窗留下未完成的 Word 验证。

受检文件由 [source-hashes.json](acceptance/n0-n10-review/source-hashes.json) 固定。正式证据包括新构建的原始 DOCX/PDF、日志、反例输入输出与观察 JSON；没有重写历史 PDF 来消除警告。现有 NW 测试自动清理自己的临时成果，本轮保留其完整测试日志，并另存上述两条独立 CLI 的成果。

## 2. 新增与残留问题

### D01 · P1 · preserve 的不同字号仍被当作同一字号验证，形成误拒绝和漏检

**关联：N3/N4，原 C03 的未关闭边界。**

来源同一段的两个文本区间使用 Arial 12 pt 和 Arial 24 pt，策略为 `preserve + page_policy: source`。实际 render_body 和最终装配正确保留两个字号；验证器却报“期望 12.0，实际 24.0”。这条反例已走生产 CLI 和真实 Word，页级 QA 通过后在格式门禁退出，未发布。

将第二个区间错误地改为 12 pt 后，实际格式验证器反而返回 `passed=true`，覆盖仍显示 expected=1/mapped=1。这说明仅修复粗斜体逻辑区间还不足以保证 preserve 语义。

代码定位：`lib/content_integrity.py:1348–1351` 读取来源首 run 与各 run 快照，但 `1389–1399` 的字体/字号检查始终取 `expected_run`；实际逐 run 比较对象变化了，期望仍固定在首 run。`source_effective_run_spans` 已保存，却没有完整用于字体、字号等保护属性。

**关闭条件：** preserve/mixed 保护区的字体、字号、颜色等逐逻辑文本区间使用来源期望；run 拆分/合并不能改变结论。增加合法 12/24 pt、不同中英字体及损坏后的反向断言。合法样例必须在最终装配及真实 CLI 通过，改坏样例必须失败。

### D02 · P1 · 保护区页眉页脚及其位置仍不在完整验证覆盖内

**关联：N3/N4/N8，原 C03/C04 的覆盖边界。**

来源包含明确页眉 `EXPECTED ANONYMOUS HEADER` 和页脚 `EXPECTED ANONYMOUS FOOTER`。对正确装配结果，把二者改成错误内容，同时将 header_distance/footer_distance 改为 40 mm；实际 `verify_delivery_format(..., verification_context=...)` 仍返回 `passed=true`、零偏差，覆盖显示 expected=1/mapped=1。

代码定位：`lib/docx_inspector.py:1074` 起仅遍历正文及表格段落建立 blocks，未将页眉页脚文本/格式加入相应节点清单。`lib/content_integrity.py:1203–1255` 的页面验证字典只含纸张和四边距，不含页眉页脚距离。即使来源节快照保存了距离，循环也只挑选目标字典中的六项。

这是最终格式验证函数的故障注入证据，没有声称整个发布流水线已经发布了该损坏文件。它已足以证明“保护范围完整验证”的声明不成立。

**关闭条件：** 检查清单覆盖适用的 default/first/even header/footer story、共享关系及继承规则；按节核验距离。增加改内容、改字体、改距离的失败用例，补真实不同首页/奇偶页眉的 PDF 对应页面断言；覆盖分母必须包含这些必需对象与属性。

### D03 · P2 · 内容门禁仍只检查媒体数量下限，额外图片实例被放行

**关联：N5，原 C05 向非文字对象扩展仍未完成。**

新建仅有 1 张图片的匿名来源，在同一段中额外复制图片 drawing，使用独立 docPr ID、复用合法图片关系；正文文字保持不变。对真正的 `ExpectedInventory` 执行 `verify_content_integrity()`，结果为 true。实际 inventory 清楚记录 source_instances=1、actual_instances=2。

代码定位：`lib/content_integrity.py:754–767` 只计算 `required_media - actual_media` 与缺失公式；没有检测额外实例。正文的严格顺序检查无法识别不带新文字的额外 drawing。相近的表格、注释检查也仍使用下限逻辑，应一起补查，本文只将图片多一例列为已复现缺陷。

**关闭条件：** 对每个内容部件精确核对媒体/公式等实例次数与引用位置，生成对象使用范围明确的独立许可；增加“多一个、少一个、位置变化、声明复用”的断言。不能仅把文件资源哈希 set 相同当作实例完整。

### D04 · P2 · 缺失中间层级时按字号连续编号，并给予自动接受高分

**关联：N7，原 C07 尚未满足困难样例要求。**

构造无大纲、Normal 样式样例，包含 `1. Introduction`（18 pt）和 `1.1.1. Detailed protocol`（14 pt），其余均为普通长正文。实际 `analyze_format_sample()` 将第二个标题选为 `heading.2`，分数为 **0.985**，高于 0.98 自动接受阈值。编号已经明确包含三级信息，但字号排名把缺失的二级强行压平了。

代码定位：`lib/role_candidates.py:108–111` 将不同字号依次映射为连续层级；`128–138` 所谓 repeated heading style 只要求同字号出现至少 1 次且全篇有两种字号；`156–161` 直接提升到自动接受分数。编号仅作为布尔信号，没有把编号深度及冲突用于层级判断或降低置信度。

**关闭条件：** 将编号结构、层级缺失及格式冲突作为独立证据；无法确定时给待确认候选，不以连续字号名次补层级。增加缺层级、同层不同字号、不同层同字号、正文编号等冻结验收例。分数应与验证结果相符，不能通过固定提升至阈值之上满足覆盖率。

### D05 · P2 · 自动接受精确率不检查具体层级，错误自动映射仍计为正确

**关联：N7/N10，评测门槛的统计口径不完整。**

沿用现有 `evaluate_unstructured_dataset.evaluate()`，以 D04 的两个标题为基准，一个角色正确、一个角色错误。报告却给出 automatic_acceptance_precision=1.0、coverage=1.0，同时自身 per_level 又显示 heading.3 的 exact recall=0，correction_operations_estimate=1。

代码定位：`docs/acceptance/r7-evaluation/evaluate_unstructured_dataset.py:148–151` 只判断来源是标题且预测也是标题，就增加 `auto_accept_correct`，未要求 `selected == labels[path]`。该值可以描述“标题/正文二分类”精确率，不能用于证明自动角色/层级映射达到 98%。以具体角色计，本反例为 **1/2=50%**。

现有无大纲样本虽已移除 outlineLvl，但生成器仍固定按五个递减字号依次给五级标题，未覆盖 D04 的结构变化。本轮复跑该集成功，并不能消除统计口径问题或证明困难样例已覆盖。基线写入脚本中的 `manually_reviewed` 标记本身也不是本轮独立人工复核证据；本报告不据此推断既往是否发生过人工检查。

**关闭条件：** 分开报告二分类、精确角色/层级的自动接受 precision、coverage 和分组结果；用于放行自动 RoleMap 的指标必须要求角色一致。增加“预测仍是标题但层级错误”的评估器回归，补文档家族隔离与困难样例，重新冻结结果。

### D06 · P2 · NW 测试名称已齐全，但若干原计划场景没有真正被覆盖

**关联：N8/N10，原 C08 未关闭。**

本轮 12 项真实 Word 测试确实全部通过，问题在于夹具和断言仍弱于计划：

| 原计划必需项 | 当前实际实现 | 缺失的验收证据 |
| --- | --- | --- |
| NW06：独立声明、中英摘要、正文等部件与 source.regions | 仍调用 `create_complete_thesis_fixture.py`；清单仅 cover/toc/body，Abstract 内容仍由中文函数生成 | 独立前置选区、真正英文摘要、按部件顺序/重复校验 |
| NW08：不同首页/奇偶页眉及横版节 | 操作的是普通 `section.header`；未启用/填充 first/even 变体，后节未解除与前节链接 | 三种页眉各自适用页面、横版 PDF 页面和实际页眉文字对应关系 |
| NW09：两来源同 noteId、注释含链接/图像、不串来源 | 同一份源复制为 one/two；仅断言脚注 XML 的同一句话至少出现两次 | 不同注释内容与对应引用、图像/链接关系闭包，无法检测串源 |
| NW12：生产故障与发布事务 | 构建成功后另存 corrupted.docx，直接调用验证函数，再检查从未被该调用写入的原文件哈希 | 真正失败的第二次构建、旧 DOCX/元数据原子保护、发布失败/不收敛路径 |

代码定位：`tests/test_post_review_word.py:236` 起的 NW06、`272` 起的 NW08、`294` 起的 NW09，以及 NW12；另见 `docs/acceptance/create_complete_thesis_fixture.py:100` 与 `160–182`。现有旧离线事务测试可以提供部分补充，不能代替这里宣称完成的真实生产失败场景。

**关闭条件：** 按优化计划 NW 表格修正夹具和断言；每个状态必须链接实际场景，而不是仅按测试名称登记完成。缺少的能力应实现并验收，或明确缩小范围并保持对应原计划项 partial，不能继续全批次 verified。

## 3. 已确认的修复与逐批结论

| 批次 | 本轮结论 | 说明 |
| --- | --- | --- |
| N0 | 原反例回归机制已建立 | 应纳入 D01–D06；历史状态不替代当前结论 |
| N1 | 原 C01 基础路径已确认修复 | 独立重建后无目录 DOCX/PDF/实测页均为 4；不外推到全部字段开关/故事流 |
| N2 | 原 C02 基础路径已确认修复 | NW03 及模板/默认/冲突/冻结回归通过，v3 默认封面不再自动吸入旧模板 |
| N3 | 部分完成 | NodeRef/实例/书签映射与配置哈希回归通过，但保护 story 和属性分母仍不完整（D02） |
| N4 | 未通过 | D01/D02；原局部粗体及未知样式反例已改善，但不能据此关闭整个 preserve 契约 |
| N5 | 部分完成 | 原文字错序/重复、生成内容定位回归通过；图片实例仍存在 D03 |
| N6 | 原 C06 已通过本轮回归 | 同路径、相对别名、符号/硬链接保护用例通过；源文件保护已接入 |
| N7 | 未通过自动识别验收 | 候选能力已增加，原固定无结构集可通过；D04/D05 阻止自动角色精确率声明 |
| N8 | 现有 12 项通过，原矩阵未完整验收 | D06；不得混淆“现有断言通过”和“计划范围已覆盖” |
| N9 | 已知警告模式得到补充确认 | 新 PDF 符合已保存诊断的孤立 offset=0 xref 模式，页数/框一致；不声明任意 PDF 兼容 |
| N10 | 不满足整体发布验收 | 当前“已完成/verified”摘要与 D01–D06 不一致，应在修复后重新同步证据 |

本轮里程碑 A 仍受 D01–D03 阻断；里程碑 B 的自动层级映射受 D04/D05 阻断，人工校正流程的存在不消除自动指标问题；里程碑 C 仍需补齐 D06 的实际场景。

N9 限制：不同解析器的规范化整页文字哈希仍可因提取顺序等产生差异，本轮没有将整页文字哈希一致作为已通过项；针对本文结论，PDF 中的 PAGEREF=4 已单独由 Poppler 文本提取确认。现有诊断分类仅适用于可证明未引用的缺失 xref 条目，不能扩展为忽略其他结构警告。

## 4. 修复优先顺序

1. **先修 D01/D02。** 将来源保护属性统一为逻辑文本区间与各 story/节的预期，修复合法保留误拒绝和损坏属性漏检；把本轮两条反向例纳入正式回归。
2. **补 D03 的精确实例检查。** 来源对象和合法生成对象分开计数，按范围和位置核验。
3. **先修 D05 的评估器，再调 D04 的识别策略。** 让错误层级真实降低精确率，加入困难样例后重新评测，避免继续优化错误统计目标。
4. **补 D06 的真实场景，再更新 N8/N10 状态。** 重新生成成果并对最终 DOCX/PDF 断言，不沿用旧测试名称或历史 exit 0。

不需要重做已经独立确认的基本封面、无目录 PAGEREF 和迁移修复；保留其回归，集中修复上述具体边界。

## 5. 复现入口与交付记录

```bash
python3 -m unittest discover -s tests -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access python3 -m unittest discover -s tests -p 'test_post_review_word.py' -v

python3 docs/acceptance/review_n0_n10_probes.py output/n0-n10-review/probes output/n0-n10-review/observations.json
python3 synthesize.py --source output/n0-n10-review/probes/mixed_sizes/source.docx --manifest output/n0-n10-review/probes/mixed_sizes/manifest.json --plan
python3 synthesize.py --source output/n0-n10-review/probes/mixed_sizes/source.docx --manifest output/n0-n10-review/probes/mixed_sizes/manifest.json --doctor
DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access python3 synthesize.py --source output/n0-n10-review/probes/mixed_sizes/source.docx --manifest output/n0-n10-review/probes/mixed_sizes/manifest.json --output-dir output/n0-n10-review/preserve-output --keep-work
```

[复验脚本](acceptance/review_n0_n10_probes.py) 未使用 mock。除明确标记的真实 CLI 外，其他反例均来自实际来源/装配/验证器或实际分析器/评估器调用；不将这些函数级故障注入描述为已发布的错误成果。脚本 exit 0 仅表示观察执行完毕，应读取结果里的 passed/accepted，而非把退出码当作验收通过。

正式证据目录为 [acceptance/n0-n10-review/](acceptance/n0-n10-review/)，含基线哈希、两个测试组日志、独立 Word 构建日志、最终/拒绝成果、反例及评估结果。临时工作目录与渲染缓存在固定证据后清理，可用脚本重建；本轮没有覆盖历史报告、原始输入或既有用户改动。
