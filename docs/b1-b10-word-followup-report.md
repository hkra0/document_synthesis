# B1–B10 真实 Word 补充验证

日期：2026-09-05。关联 [初次验收报告](b1-b10-acceptance-report.md) 和 [详细补全计划](b1-b10-completion-plan.md)。

## 结论

**Word 访问授权问题已解除，本轮补测已执行完毕；产品验收仍不通过。** 本轮未修改生产代码。

真实 Word 基础分页与现有 M1 测试通过，但独立检查发现 A4 配置没有生效；新增 M2 编号样本被页级 QA 拦截，PAGEREF 样本在真实测量后的第二轮装配异常退出。不能再把这两个 M2 失败归因于未授权或未执行 Word。

本报告中的“补测完成”指下表列出的检查全部得到结果。完整论文涉及的脚注/尾注跨文件合并、全部选区、首页/奇偶页眉页脚及复杂字段组合，仍须按补全计划修复后重新验收；本轮没有把这些未覆盖项判为通过。

## 实际结果

| 编号 | 验证内容 | 结果 | 证据与解释 |
| --- | --- | --- | --- |
| W01 | Word Automation 与输入预检 | 通过 | [doctor 日志](acceptance/word-followup/doctor.log)、[只读 plan](acceptance/word-followup/plan.log)。 |
| W02 | 原分页测试文件，开启真实 Word | 7/7 通过、无跳过 | [日志](acceptance/word-followup/pagination-tests.log)。其中 1 项为此前跳过的真实 Word 用例，另 6 项是已有离线测试；不能称为 7 个真实 Word 用例。 |
| W03 | 原 M1 验收文件，开启真实 Word | 3/3 通过、无跳过 | [日志](acceptance/word-followup/m1-tests.log)。覆盖纯正文、样例分析/编译/构建、academic/report 切换；测试中的 `lib.delivery.inspect_document` 替换关闭，实际调用 Word。 |
| W04 | 从唯一 CLI 构建 academic 匿名样例，再独立检查 | 构建成功，格式验收失败 | 最终 DOCX 和 Word PDF 均为 Letter 215.9 × 279.4 mm，配置要求 A4 210 × 297 mm。最终元数据仍记录格式 passed。 |
| W05 | 新增罗马目录、正文重启、oddPage 配置的 M2 样本 | 两轮 Word 导出完成，未发布 | 最终共 7 页；可见标签为 i、1、2、3、4、5、6。第 4、7 物理页各剩两行正文，被既有页级 QA 拦截。 |
| W06 | 在 W05 源文档加入指向第二章的 PAGEREF | 第一轮 Word 导出完成，第二轮装配崩溃，未发布 | `lib/field_updater.py:88` 抛 `AttributeError: 'str' object has no attribute 'get'`。与初次验收 F06 的独立探针一致。 |

结构化结果见 [summary.json](acceptance/word-followup/summary.json)。175 项离线基线结果保留在初次报告中；W02/W03 是定向补测，不把重复运行的测试累加为新增测试总数。

### W04：实际发布成功，页面契约没有满足

可追溯文件：

- [最终 DOCX](acceptance/word-followup/academic.docx)、[同次 Word PDF](acceptance/word-followup/academic.pdf)。这些是失败反例证据，不是合格模板。
- [最终元数据](acceptance/word-followup/academic-metadata.json)、[独立测量](acceptance/word-followup/academic-observations.json)、[构建日志](acceptance/word-followup/academic.log)。

检查了整页渲染，中文标题和正文可见，没有发现该页文字重叠或裁切；页脚仍为 `- 1 -`。最终 DOCX 下边距约 24.606 mm，预设要求 25 mm。纸张错误在 DOCX 与 PDF 中一致，不能用渲染近似或字体替换解释。

该结果增强了初次验收 F02/F04：实际 Word 构建成功与格式配置正确是两个必须独立验证的条件。M1 现有测试只核对部分属性，因此全部通过也没有发现这个错误。

### W05：基础标签可见，但整体排版不合格

证据：[未发布 DOCX](acceptance/word-followup/m2-numbering-unpublished.docx)、[最后一轮 Word PDF](acceptance/word-followup/m2-numbering-unpublished.pdf)、[逐页文字和分节记录](acceptance/word-followup/m2-numbering-observations.json)、[失败日志](acceptance/word-followup/m2-numbering.log)。

目录条目指向正文 1、4，正文第一章和第二章分别出现在物理页 2、5；真实标签分别为 1、4。配置的正文 `oddPage` 与实际物理起页的含义还需要在实现契约中明确区分“印刷编号奇偶”与“装订后的物理左右页”；不能仅看到 sectPr 中的 oddPage 就认定右页起章完成。

逐页检查了 7 页渲染。Poppler 对第 4、7 页只显示页脚，而文本提取报告两行正文；额外使用 MuPDF 渲染第 4 页，能看到这两行。日志还存在 PDF 对象偏移警告。这里同时记录渲染器差异，暂不推断其根因。即便采用能显示正文的 MuPDF，页级 QA 所报的两行内容事实仍成立。

不能为让该样例通过而放宽全部项目的孤行门槛。应先验证源段落分页属性、最终托管样式、节几何和 QA 规则是否一致，再用明确预期的长短段落夹具建立回归。

### W06：真实跨页引用链路复现异常

证据：[完整 traceback](acceptance/word-followup/m2-pageref.log)、[第一轮 DOCX](acceptance/word-followup/m2-pageref-first-pass.docx)、[第一轮 Word PDF](acceptance/word-followup/m2-pageref-first-pass.pdf)。它们仅用于定位失败，第一轮产物不代表收敛后的最终交付物。

delivery 将测量值压缩为标签字符串传入下一轮 assemble_document；后者直接交给 FieldUpdater，PAGEREF 按字典调用 `.get()`。本轮没有替换 Word、装配器、字段更新器或发布器，因而确认这是生产链路的参数契约冲突。

## 复现方法

仓库根目录运行。已有测试命令如下，真实 Word 自动化须在本机可用：

```bash
DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p test_pagination.py -v
DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p test_m1_acceptance.py -v
```

W05/W06 的匿名材料由 [生成器](acceptance/create_word_followup_fixtures.py) 创建。生成器只造测试材料；构建仍调用 synthesize.py。示例：

```bash
python3 docs/acceptance/create_word_followup_fixtures.py output/word-followup-fixtures
python3 synthesize.py --source output/word-followup-fixtures/numbering-source.docx --manifest output/word-followup-fixtures/numbering-manifest.json --plan
python3 synthesize.py --source output/word-followup-fixtures/numbering-source.docx --manifest output/word-followup-fixtures/numbering-manifest.json --doctor
python3 synthesize.py --source output/word-followup-fixtures/numbering-source.docx --manifest output/word-followup-fixtures/numbering-manifest.json --output-dir output/word-followup-numbering --keep-work
```

W06 将 source 与 manifest 文件名前缀替换为 `pageref`，并使用独立输出目录。源夹具及预期已留存在 [m2-fixtures](acceptance/word-followup/m2-fixtures/fixture-expectations.json)。

## 对后续工作的影响

1. 将旧结论“真实 Word 验收因权限未完成”更新为“基础/M1 定向检查通过，M2 真实集成失败”。授权问题已经解决。
2. 保持 F01–F10 未关闭。W04、W06 提供更强的真实链路反例，W03 不足以反证这些缺陷。
3. 增加 PDF 渲染差异诊断、分页控制与物理起页契约专项任务。
4. 补全实现后再跑完整 M1/M2 矩阵。具体代码批次、接口和退出条件见 [详细补全计划](b1-b10-completion-plan.md)。

保留的文件是正式验收证据；本次 `.work`、临时渲染图片和探索输出完成记录后清理。用户关于所有匿名验收临时目录的 Word 访问授权持续有效。
