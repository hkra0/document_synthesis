# R7 独立匿名识别评估集

这里的评估集独立于 `tests/fixtures/formatting/` 开发夹具。它只包含程序生成的虚构中文材料，不包含业务正文或用户文件。

数据规模固定为四类样例各 5 份：公文、报告、论文、手工格式，共 20 份 DOCX；每份含 5 个已标注标题和 25 个正文负例，因此合计 100 个标题正例与 500 个正文负例。标题使用显式 `w:outlineLvl`，正文不带大纲级别；类别之间改变样式、字体、对齐和直接格式，避免评估只依赖单一样式名。

生成并评估：

```bash
python3 docs/acceptance/r7-evaluation/generate_dataset.py
python3 docs/acceptance/r7-evaluation/evaluate_dataset.py
```

`dataset.json` 保存标签和计数，不保存样例正文；`metrics.json` 保存各类别、各标题级别的 precision/recall，以及指标 v2 的
`heading_detection_precision`、`exact_role_auto_precision`、
`automatic_acceptance_coverage`、`correct_auto_coverage` 和
`candidate_exact_role_coverage`。评估脚本不会调用 Word 或网络。

自动接受阈值仍为 0.98。只有直接结构证据达到阈值的候选才可计入自动接受；其余候选继续要求人工确认。

## N7 无大纲辅助集

无大纲集独立生成 20 份匿名虚构 DOCX，同样覆盖四类、100 个标题正例和 500 个正文负例；标题统一使用 `Normal`，不写入 `w:outlineLvl`，但保留字号、对齐、间距、编号和局部强调等混合信号。它与上面的结构化集分开统计，不把同源副本当作独立样本。

```bash
python3 docs/acceptance/r7-evaluation/generate_unstructured_dataset.py
python3 docs/acceptance/r7-evaluation/evaluate_unstructured_dataset.py
```

当前证据报告前三候选的标题覆盖率、按标题层级的候选覆盖、候选阶段误报/precision、最终选定角色、校正成本估计和未决数。冻结集标签保存在 `unstructured-human-baseline.json`，由独立的可见段落复核固定，候选排序器不参与标签决定。自动接受的 precision 现在指精确角色（含层级）precision；旧字段 `automatic_acceptance_precision` 仅作为带语义说明的兼容别名。`manual_correction_count` 在没有真实操作日志时保持 `null`，理论纠正数另记为 `correction_operations_estimate`。结果仍只适用于该匿名冻结集，不外推为任意真实文档的准确率保证。

困难样例压力集由 `generate_hard_cases.py` 生成，清单和每个实际 DOCX 的 SHA-256 保存在 `unstructured-hard-cases.json`。它覆盖缺失中间层级、同层异字号、日期/版本号、正文编号、短粗体正文和题注；该压力集用于验证弃权与冲突解释，不并入主冻结集指标，也不把没有真实审阅者操作日志的结果写成自动放行证据。

```bash
python3 docs/acceptance/r7-evaluation/generate_hard_cases.py
python3 -m unittest tests.test_role_candidates.RoleCandidateRemediationTest.test_S5_frozen_hard_cases_bind_real_fixtures_and_fail_closed -v
```
