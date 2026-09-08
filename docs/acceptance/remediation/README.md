# N0–N10 remediation evidence

本目录保存 `docs/post-n0-n10-remediation-plan.md` 对应的 S0–S8 证据，不覆盖历史 R/N 批次。审阅页导出的 `review_audit` 只证明操作轨迹与源材料绑定；只有 `origin=browser_ui` 且 `automation=false` 的记录可以声明 `reviewer_attested=true`，`automation=true`、`origin=test_fixture` 或 `origin=imported` 的记录均不能替代真人确认。

真实审阅交接可从匿名困难样例生成：
`python3 synthesize.py --analyze-format docs/acceptance/r7-evaluation/unstructured-hard-samples/公文-hard-cases.docx --analysis-dir <fresh-review-dir> --replace-output`。审阅者需要在 `review.html` 中处理未决项、至少完成一项校正、勾选人工确认并导出 `decisions.json`；随后再用 CLI 编译格式包/RoleMap。当前真人闭环证据见 [`s5-human-review-evidence.json`](s5-human-review-evidence.json)。

S0 先把 D01–D06 的场景、前置事实、命令和结果字段固化为
[`scenario-contracts.json`](scenario-contracts.json)。初始 `todo` 不表示通过；本轮已在逐项核对历史反例、当前正常/损坏对照和正式产物后写入 `verified`，状态不是由测试名称自动推导。

当前正式反例来自 [`review_n0_n10_probes.py`](../review_n0_n10_probes.py) 的匿名输入。旧探针的 exit 0 只表示探针运行完成，不表示缺陷已关闭；本轮测试会直接断言应通过或应拒绝的结果。

当前状态索引见 [`status-summary.json`](status-summary.json) 及各批次 sidecar。真实 Word 测试统一使用固定目录
`<word-access-dir>`，运行时会自动创建该目录；当前 NW01–NW12 为 12/12，remediation_word 为 11/11，均未再次出现授权弹窗。S5 的未完成范围仍保留在对应 sidecar 的 `remaining_limits` 中。
