# 开发说明

## 统一构建流程

`synthesize.py` 是唯一入口。构建先在 `output/<项目>/.work/run-*` 生成隔离产物，正文渲染一次，再按规范化交付清单装配。所有声明文档的书签、目录页码与 PDF 页级 QA 通过后才发布，QA 失败不替换既有成果。

配置加载优先级为显式 `--manifest`、项目内 `manifest.json`、`profiles/<项目>.json`、`manifests/<项目>.json` 和默认目录扫描。

R0–R11 的后续优化以 [`docs/post-r0-r11-optimization-plan.md`](docs/post-r0-r11-optimization-plan.md) 为准，当前前置批次状态保存在 [`docs/acceptance/post-review/status.json`](docs/acceptance/post-review/status.json)。N0–N10 复验后的补全以 [`docs/post-n0-n10-remediation-plan.md`](docs/post-n0-n10-remediation-plan.md) 为准，当前 S0–S8 状态保存在 [`docs/acceptance/remediation/status-summary.json`](docs/acceptance/remediation/status-summary.json)。实现完成与真实 Word/PDF 验证分开记录；blocked 或 partial 不得写成 verified。

只选择一个基础配置，随后按顺序应用显式 `--override`。对象递归合并，数组整体替换，标量替换，`null` 不作为删除操作；合并后统一验证未知字段、输出路径和引用。

## 输出契约

- v2 与无配置项目默认一份 `main`，包含 `cover`、`toc`、`body`；正文骨架留在内部工作目录。
- `output.documents` 是完整替换清单。ID 与文件名唯一，部分按 `cover → toc → body` 顺序选择，不能重复。
- 带正文的目录引用自身；独立目录引用另一份含正文的交付物，不写本文件跳转链接。
- v1 规范化为旧三份清单并提示迁移。私有 profiles 不自动改写，不自动删除旧成果。
- 封面与目录不显示页码，正文编号从 1 开始。物理页位置与正文打印编号必须分开，不能通过封面页数推算目录页码。
- 书签 ID 是页码映射键；同名标题不互相覆盖。回填后重新检查装配文档，有上限地迭代直至页码稳定，否则拒绝发布。

目录一致性验证范围仅包含引擎生成并标记的交付目录。源内容内的静态目录、Word 原生 TOC 域和缓存页码不由当前流程统一重建或核验；Word 的 repaginate 不等价于更新 TOC 域。不得用同名标题或模糊文本匹配替这些内部目录推测页码，亦不得把交付目录 QA 通过表述为所有源目录页码均已通过。

## 来源策略

- `directory_tree` 按目录和文件名扫描材料。
- `explicit_tree` 由 profile 声明树、顺序和标题。`docx_outline` 可将 Word 正文标题映射为目录书签。
- `highlighted_docx` 处理单一长篇 DOCX。高亮颜色、标题规则与替换项在 `source.highlight` 配置。

## 模块边界

- `lib/config.py` 提供版本化配置（v1/v2/v3）与安全校验。
- `lib/format_schema.py` 提供格式包规范、不可变数据模型（ResolvedFormat/StyleDefinition/PageSpec）与几何校验。
- `lib/format_resolver.py` 负责格式包引用（`preset:` 或文件路径）、单继承与能力匹配拦截。
- `lib/legacy_format.py` 将 v1/v2 历史配置无缝转译为标准规范格式包，隔离旧公文清洗策略。
- `lib/docx_inspector.py` 只读 OOXML 深度检查与基于有向拓扑图的有效属性层叠求值器（EffectiveStyleEvaluator）。
- `lib/style_applier.py` 统一样式应用引擎，安装具名 Word 样式、绑定大纲级别并保护行内语义（上下标/局部强调/公式）。
- `lib/layout.py` 版面几何几何与 ContentBox 求解器，提供多后端排版度量。
- `lib/package_importer.py` 负责从已排版 Word 文档导入段落，处理媒体部件、书签与原生编号重映射。
- `lib/role_mapper.py` 启发式与显式语义角色映射器，使用稳定 XML 路径 NodeRef 绑定段落与表格。
- `lib/format_analysis.py` 样例分析与样式聚类器，推导格式候选并生成字段级证据链。
- `lib/format_review.py` 单文件自包含 HTML 人机交互校正报告生成器与零正文泄露格式包/角色映射编译器。
- `lib/content_integrity.py` 段落多重集内容无损验证与最终交付有效格式核验器。
- `lib/source_strategies.py` 解析材料来源并分配书签 ID，支持目录树、显式大纲与标准单 DOCX 策略。
- `lib/highlighted_docx.py` 提取高亮标题，标准化正文并处理连续页码。
- `lib/renderers.py` 负责 DOCX、PDF、PPTX 和图片的保真渲染。
- `lib/composition.py` 负责复用正文产物装配声明的部分与边界书签。
- `lib/delivery.py` 提供构建和 smoke 共用的交付结构、目录页码与页级验证。
- `lib/pagination.py` 读取 Word 书签的物理页与打印编号。
- `lib/engine.py` 负责统一来源调度、有限次目录收敛、隔离工作目录、build-metadata.json 审计与事务原子发布/回滚。
- `lib/qa.py` 负责 Word 自动化、页码读取和 PDF 页级断言。

## v3 格式架构与审计契约

1. **版本声明与隔离**：
   - `schema_version: 3` 引入 `format.ref`（支持 `preset:academic-basic@1.0.0`、`preset:report-basic@1.0.0` 等）与 `formatting.mode`（`restyle` 规范重排 vs `preserve` 维持源版式）。
   - 彻底禁止在 v3 顶层混用 `page_setup` 与 `fonts`。
2. **审计元数据 `build-metadata.json`**：
   - 在 v3 项目交付目录中原子发布，记录引擎版本、生效格式包、配置契约 SHA-256、源文件 SHA-256 快照、交付物哈希及 QA 结果。
   - `PreparedBuild` 在装配前及发布前核对同一配置契约哈希；`smoke_test.py` 会重新核对配置、格式包、源文件和每份 DOCX 的哈希，不信任历史 `passed` 字段。
   - 历史 v1/v2 项目严格不生成该文件，确保旧交付目录脚印零变动。
3. **内容无损守恒**：
   - 严禁通过直接格式暴力覆盖抹除作者行内公式、上下标与局部强调。
   - 采用段落多重集（Multiset Counter）频次比对，杜绝同名段落或标题的意外遗漏与倍增。
4. **发布事务原子回滚**：
   - 临时工作目录 `.work/run-*/` 内准备就绪后原子暂存发布；若发生任何 IOError/OS 异常，立即触发全量文件原子回滚，并保留恢复信息。

## 验证

```bash
python3 -m unittest discover -s tests -v
python3 synthesize.py --project <项目> --plan
python3 synthesize.py --doctor --project <项目>
python3 smoke_test.py --project <项目>
```

`--plan` 必须保持只读且不调用 Office；测试覆盖默认完整文档、显式拆分、配置覆写、v1 兼容和拒绝错误输出清单。Word 不可用时，`smoke_test.py --structure-only` 只作结构与文字检查，不能声称打印页码或实际版式正确。

`smoke_test.py` 从 `UnifiedSynthesizer.plan()` 获取实际来源节点和交付清单，不能回退到猜测的旧文件名；每个声明文档均检查，包含正文的先检查以供独立目录引用页码。单文件来源与目录来源遵守相同交付协议。

## R7–R11 格式发布闭环

格式样例工作流必须经过四个明确阶段：分析样例、人工确认 decisions、编译格式包、为目标文档编译绑定源哈希的 RoleMap。`decisions.example.json` 是草稿，含 `pending` 的角色不能传给编译器；最终决策必须保留 `report_id` 和 `source_sha256`，源文件变化后旧决策应被拒绝。

`--preview-format` 与 `--render-preview` 只生成虚构内容的近似预览，并写入引擎标识和 `source_content_included: false` 的元数据。`--migrate-manifest` 必须把 v1/v2 输入写到独立目标文件，不能覆盖输入或静默删除旧交付物。所有动作都通过 `synthesize.py` 分发，不新增项目专项构建入口。

发布前至少核对以下证据层级：

- R7：`tests.test_r7_workflows`、独立匿名评估集和 `r7-status.json`。
- R8：PageRecord/页面序列定向回归与 W05 的 Word/PDF 标签、物理页面证据。
- R9：字段/脚注/尾注定向回归与 W06 的 PAGEREF 收敛证据。
- R10：全量离线回归、M1、W04–W07、M2 装配级与完整匿名论文真实 Word/PDF 证据。
- R11：示例命令在干净输出目录成功执行，状态文件、能力矩阵、README、迁移说明与实际 CLI 参数一致。

当前 R7–R10 的最终数字以 `docs/acceptance/r7-status.json` 至 `r10-status.json` 为准。当前边界不覆盖 OCR/PDF 样例学习、双栏期刊、参考文献著录转换、多样例融合或全部学校规范认证。

临时验证产物、Office 锁文件、`.work` 目录和 Python 缓存均不应提交。默认清理构建工作目录；显式 `--keep-work` 用于诊断保留，诊断完成后清理该次目录。

发布前先备份本次会替换的文件，逐个替换失败时回滚。若文件系统故障导致回滚也失败，工作目录与备份强制保留并报告恢复位置；这不是跨文件的崩溃原子事务。

Word 集成测试需要显式设置 `DOCUMENT_SYNTHESIS_WORD_TEST=1` 后运行。真实 Word 导出统一使用固定目录 `DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access`，避免每次随机 `.work` 子目录都触发 “Grant File Access”；首次运行仍需在 Word 的文件夹授权对话框中选择这个专用目录。离线流水线测试只模拟 Word 测量接口，真实执行装配、结构检查、合成 PDF 页级检查与发布；不能替代真实 Word 导出验收。Automation 权限与 Word 的 “Grant File Access” 文件夹权限分别检查。

## 2026-09-04 本机 Word 验收

在用户授权的项目 `output` 范围内，匿名示例已完成真实 Word 导出、最终目录页码核验、PDF 页级 QA 和逐页 PNG 目视检查。

| 交付组合 | Word 导出页数 | 目录及正文编号 |
| --- | --- | --- |
| 封面＋目录＋正文 | 4 | 两个目录条目指向物理第 3、4 页，正文打印编号为 1、2 |
| 目录＋正文 | 3 | 两个目录条目指向物理第 2、3 页，正文打印编号为 1、2 |
| 封面＋目录 | 2 | 引用上述正文的打印编号 1、2，不含本文件跳转 |

三份导出 PDF 的链接目标与上述位置一致；逐页检查无空白页、孤行、单标题页、裁切或重叠。两个含正文交付物各通过源文档 20 段文字保留检查。真实 Word 集成测试另验证了同名标题和正文重新从 1 编号；原始材料和既有业务交付物未被改写。

验收样本为 `examples/minimal-demo/manifest.json`，拆分使用同目录的 `split.override.json`；样本成果保留在 `output/acceptance-full/` 和 `output/acceptance-split/`。此记录仅证明本次匿名样例与书签集成测试通过，不代表其他项目已重建验收。Word 可能仍会逐次要求授权新建的临时子目录，不应把一次授权描述为永久免提示。
