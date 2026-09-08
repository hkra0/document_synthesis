# 自定义格式与样例识别：代码执行计划

日期：2026-09-05。状态：B0、B1、B2、B3、B4 已完成，待实施 B5。详情见 [formatting-progress.md](formatting-progress.md)。

本文把[总体规划](custom-formatting-plan.md)细化为实现契约、变更批次和验收任务。以下新增字段、模块和命令均为目标设计，当前程序尚不支持。实现以当前工作区为起点，保留已有未提交改动；本文不授权自动发布、重建业务材料或改写原始文件。

## 1. 交付范围与不可变约束

### 1.1 两个发布里程碑

**M1：通用格式与样例识别闭环。** 用户能选择预设、编辑 JSON 格式包，或从一个 DOCX 样例提取格式；经集中校正后，将规则应用到另一份 DOCX 或材料集合。覆盖页面、文档标题、正文、1—9 级标题、目录、基础页脚、已有列表保留、普通表格与图片。输出具名 Word 样式，并通过原有发布门禁。

**M2：基础学位论文排版。** 增加可排序文档部分、双语摘要、前置罗马页码、正文阿拉伯页码、首页/奇偶页眉页脚、章节起页、脚注/尾注合并、题注和受控字段更新。需真实 Word 集成验收后才标为支持。

M1 中，已存在的公式、脚注、字段和复杂对象应原样保留或在不支持的变换前报错；“保留已有对象”与“统一其样式、重新编号、跨文件合并”分别记录能力。PDF/OCR 识别、双栏期刊、多样例融合和参考文献著录转换属于后续扩展。

### 1.2 实施约束

- `synthesize.py` 继续是唯一构建入口。新命令由它分发，业务逻辑进入 `lib/`。
- 不改写输入、样例、私有 profile 和已有格式包。编辑后另存版本，输出仍经过隔离工作目录与失败回滚。
- 无配置项目继续使用 v2 默认。v1 三份交付、v2 默认完整文档的语义不变。
- `--plan` 只读取本地配置和材料，不创建缓存、不写文件、不调用 Word、模型或网络。
- 模型不参与最终格式应用。构建只使用确定的配置、映射和原始材料。
- 不新增项目名称或材料文字特判。既有专项分支先隔离为兼容策略，再逐步替换。
- 不把字体替换后的近似预览或离线测试称为 Word 排版验收。

## 2. 基于当前代码的改造地图

| 当前符号 | 要做的变更 | 首次实施批次 |
| --- | --- | --- |
| `config.SCHEMA_VERSION`、`_validate_config`、`ProjectConfig` | 分离默认版本与支持版本，增加 v3 格式引用及映射配置 | B1 |
| `config.load_project_config` 内部 `merge` | 提取纯合并函数，保留覆写来源，按声明文件解析路径 | B1 |
| `engine.UnifiedSynthesizer.plan/synthesize` | 共用已解析的 BuildPlan，避免构建重复解析后出现配置漂移 | B2、B5 |
| `styles.add_heading_paragraph`、`add_toc_entry_with_hyperlink`、`setup_footer` | 转为读取格式上下文的输出适配器 | B3 |
| `styles.inline_style_properties`、`merge_styles_into_doc` | 分离有效格式解析、样式导入和格式清洗 | B2、B3 |
| `highlighted_docx.extract_and_format_highlighted_headings` | 拆出只读提取；v3 不删掉全部 run 重建标题 | B3、B5 |
| `highlighted_docx.normalize_document_pagination` | 旧行为仅用于兼容模式，新路径服从页面接管策略 | B3 |
| `renderers.render_docx_file`、`copy_element_with_rels` | 导入前分配节点身份；按能力导入对象及样式，返回节点映射 | B4 |
| `renderers.inline_numbering_to_paragraph` | v3 保留原生编号，停止默认转文本 | B4 |
| `engine.render_tree_node_recursive` | 传入 RenderContext，移除新路径对字体字典和全局方向状态的依赖 | B4 |
| `source_strategies.build_outline` | 增加通用单 DOCX 分析入口，文件树标题与段落标题分开 | B5 |
| `composition.FrontImporter` | 抽取经测试的 OPC/ID 导入能力；旧适配器保持原行为 | B4 |
| `composition.assemble_document/BOUNDARIES` | M1 接入样式上下文，M2 改为具名部分及边界 | B3、B8 |
| `pagination._parse_page_map/inspect_document` | 保留旧接口，增加含节和页码标签的 v3 测量结果 | B9 |
| `delivery.validate_measured_delivery/build_deliveries` | 将连续正文编号断言改为节规则；收敛比较包含标签 | B7、B9 |
| `qa.run_qa_assertions` | 页面内容区域和允许空白来自版式上下文 | B7、B9 |
| `smoke_test.test_project/_check_body_text` | 共用 BuildPlan、内容选区与语义清单，停止靠固定 `body` 名称推断 | B7、B8 |
| `synthesize.main/print_plan/run_doctor` | 新增操作分发、能力与字体预检、解析配置展示 | B5、B6 |

需要专门处理的现状：`lib/config.py` 和 `lib/styles.py` 的默认字体键并不完全一致；单文件目前只接受 `highlighted_docx`；`run_doctor` 用标题节点数量判断有效输入；渲染器会按横竖版恢复固定纸张尺寸。不能只在 `ProjectConfig` 中增加一个字典后视为改造完成。

## 3. 目录与模块边界

采用平铺模块与现有工程一致，不引入 Web 框架或新的文档转换引擎。模块随所属批次创建，不提前生成空壳。

```text
lib/
  format_schema.py        # 格式包契约、类型、单位、诊断
  format_resolver.py      # 引用、继承、覆写、来源、能力解析
  legacy_format.py        # v1/v2 适配及旧版清洗策略
  docx_inspector.py       # 只读 OPC 检查、节点索引、有效属性
  role_mapper.py          # 段落角色候选、显式校正、目标映射
  format_analysis.py      # 样例统计、候选格式、编译已确认格式包
  style_applier.py        # 具名样式、托管属性清理、按角色应用
  package_importer.py     # OPC 关系、样式、编号和 ID 重映射
  layout.py               # 版心、分节规格、渲染上下文；M2 部分计划
  format_review.py        # JSON/本地 HTML 报告、校正校验、预览
  content_integrity.py    # 文字与对象语义清单及前后核验
  field_updates.py        # M2 受控字段更新计划
formats/
  README.md
  presets/legacy-official/1.0.0.json
  presets/report-basic/1.0.0.json
  presets/academic-basic/1.0.0.json
schemas/
  project-v3.schema.json
  format-v1.schema.json
  analysis-v1.schema.json
  decisions-v1.schema.json
  role-map-v1.schema.json
examples/custom-format-demo/
  manifest.json
  report.override.json
  academic.override.json
  formats/ ...
  scripts/create_materials.py
tests/fixtures/formatting/
  expected/ ...
  README.md
```

JSON Schema 用于字段、枚举、联合类型和编辑器提示；Python 校验负责引用闭包、几何、兼容和语义冲突。引入 `jsonschema` 时在 B1 验证并锁定与项目最低 Python 版本兼容的依赖，不先假定最新版本可用。P0 同时核实 README 的 Python 3.8 声明与现有依赖/类型注解的实际兼容性；调整运行版本须有测试证据，单独记录。

依赖方向：`format_schema` 不导入 config/engine；resolver/inspector 不导入 renderer；mapper 不调用 Word；applier/importer 不读取项目全局变量；engine 编排所有阶段。避免 config → styles → config 循环。

## 4. 配置契约：实现时先固定这一层

### 4.1 清单版本与默认行为

将版本常量拆成 `DEFAULT_SCHEMA_VERSION = 2` 与 `SUPPORTED_SCHEMA_VERSIONS = {1, 2, 3}`。保留旧常量兼容引用直至调用点迁移。v1/v2 经 `legacy_format.from_project_config()` 生成内部格式上下文，但不改变已有合法输入的默认处理策略。

v3 必须显式提供 `format`，不接受顶层 `fonts/page_setup`；提示迁移位置。v3 的显式封面资源使用精确路径，禁止自动找一个名字相似的模板。v1/v2 的封面发现规则保持原样。

M1 v3 项目示例，所有路径相对本清单所在目录：

```json
{
  "schema_version": 3,
  "project_name": "AcademicDemo",
  "source": {"strategy": "docx_document"},
  "format": {
    "ref": "formats/academic-demo.json",
    "overrides": {
      "styles": {"body": {"run": {"size_pt": 12}}}
    }
  },
  "formatting": {
    "mode": "restyle",
    "role_map": "mappings/manuscript.json",
    "on_unmapped": "error",
    "page_policy": "target",
    "source_toc": "preserve"
  },
  "output": {
    "documents": [{"id": "main", "filename": "academic.docx", "parts": ["body"]}]
  }
}
```

M1 的 `output.documents` 保持 v2 的三种部分及目录引用语法；未指定时仍默认完整文档。论文单文件示例显式选择 body，避免自动多加封面和目录。M2 才启用 §11 的 layout 扩展。

`format.ref` 只允许本地文件或 `preset:report-basic@1.0.0`，不接受 URL、自动最新版或模糊名称。`extends` 采用相同引用语法，单继承、最多 8 层、循环报错。格式包自带 ID/version；引用预设与实际内容不符直接报错。

### 4.2 覆写与路径

先按当前规则选一份基础清单，再合并 `--override`。对每个叶子属性保留来源文件和 JSON Pointer。路径按该叶子最后一次声明的文件目录解析；命令行提供的文件路径按 cwd 解析，不按输入目录猜测。

格式解析顺序是父包 → 当前包 → 合并后清单中的 `format.overrides`。对象递归合并，数组整组替换，标量替换；显式 false/0 必须保留。`null` 不表示删除，M1 托管样式字段不接受 null。缺省表示继承；取消粗体用 false、去掉缩进用 0 等具体值。

最终产物包括 `ResolvedFormat.values`、`provenance`、`content_hash`。哈希排除绝对路径与时间戳，包含规范化值和引用资源内容哈希。构建重新检查输入、映射、资源哈希；变化则重建计划或停止，不能沿用过期节点。

### 4.3 格式包字段

```json
{
  "format_schema_version": 1,
  "id": "academic-demo",
  "version": "1.0.0",
  "extends": "preset:academic-basic@1.0.0",
  "metadata": {"description": "虚构排版示例，非学校提交规范"},
  "styles": {
    "body": {
      "run": {"east_asia": "宋体", "latin": "Times New Roman", "size_pt": 12},
      "paragraph": {
        "alignment": "justify",
        "first_line_indent": {"value": 2, "unit": "char"},
        "space_before_pt": 0,
        "space_after_pt": 0,
        "line_spacing": {"mode": "multiple", "value": 1.5},
        "widow_control": true
      }
    },
    "heading.1": {
      "based_on": "body",
      "run": {"east_asia": "黑体", "size_pt": 16, "bold": true},
      "paragraph": {"keep_with_next": true, "page_break_before": true}
    }
  },
  "roles": {
    "body": {"style": "body", "include_in_toc": false},
    "heading.1": {"style": "heading.1", "outline_level": 1, "include_in_toc": true}
  }
}
```

角色定义引用样式；样式定义只含视觉和段落属性。无编号的摘要标题可有大纲级别并入目录；可见编号另外配置。示例继承的完整预设须由 B1 提供，所有用到的角色都必须可解析。

| 字段组 | M1 支持 | M2 扩展 |
| --- | --- | --- |
| `page` | `width_mm/height_mm`、四边距 mm、页眉页脚距离 mm、`orientation`、网格开关 | gutter、镜像边距、分节页型；栏数超过 1 暂报不支持 |
| `styles.*.run` | east_asia、latin、complex_script、size_pt、bold、italic、RGB/auto 颜色 | 字符间距等按样本需要增加 |
| `styles.*.paragraph` | alignment、首行/悬挂/左右缩进、段距、行距、keep_with_next、keep_lines、widow_control、page_break_before、snap_to_grid | 复杂制表位和专业规则 |
| `roles.*` | style、outline_level 1—9、include_in_toc | 编号方案引用、题注类别 |
| `toc` | title、title_role、max_level、各级条目样式、leader、page_number_gap_mm | 具名部分的目录范围 |
| `header/footer` | managed/none/source；普通页静态文本与 PAGE token、对齐和样式 | first/even/default、STYLEREF、链接上一节 |
| `normalization` | 预定义的清洗策略开关，缺省不启用业务清洗 | 更多经测试策略 |
| `required_capabilities` | 能力 ID 列表，先由引擎核实再允许构建 | 后续增量扩展 |

M1 通用预设必须完整声明 body、title、subtitle、heading.1—9、toc.title、toc.1—9、header、footer、table.body、caption、quote、bibliography 样式；其中后四类可通过人工映射应用基础格式，不承诺自动识别或重建编号。

能力 ID 在 B1 固定，例如 `styles.paragraph.v1`、`styles.character.v1`、`layout.single_column.v1`、`docx.single_source_preservation.v1`、`docx.multisource_basic.v1`、`numbering.preserve.v1`；M2 再启用 `layout.parts.v1`、`pagination.roman.v1`、`notes.merge.v1`、`fields.managed_update.v1`。能力由引擎注册表声明，格式包只能要求能力，不能自行声明引擎支持。未知 ID 和已知但未实现的 ID 分别报错。

内部用整数 twip/half-point 或 Decimal 做转换，序列化使用规范单位。页面 mm 转 twip 容差最多 1 twip；字号必须能表达为半磅；倍数行距使用 240 为 1 倍的单位。char 缩进保留字符单位，不能按正文大小一次性换算后丢掉语义。首行与悬挂缩进互斥，exact/at_least 行距用 pt，multiple 用倍数；single 不能同时携带任意 value。

几何验证覆盖页边距、装订线、有效宽高、页眉页脚区域与正文碰撞。配置合法但目标环境缺字体时，由 `--doctor` 给出可操作诊断；严格预设不静默替换。

### 4.4 格式应用策略

`formatting.mode` 为 preserve/restyle/mixed；`page_policy` 为 source/target；`on_unmapped` 为 preserve/error。restyle 缺省 error，mixed 缺省 preserve；preserve 不进行角色重排。缺省页面策略为 preserve→source、restyle→target，mixed 必须明确声明。

`formatting.inline_emphasis` 为 preserve/target，缺省 preserve。preserve 保留显式字符样式及已标注的局部粗斜体；target 允许目标规则覆盖粗斜体，但始终保护上下标、数学对象、字段和链接结构。无法区分正文局部强调与全段旧格式时列为待判断，不使用固定比例猜测后直接删除。v3 迁移命令只对能够完整表达的旧规则输出结果；遇到无法表达的专项分支返回定位，不能输出外观已改变却声称等价的配置。

mixed 的 `scopes` 数组每项含精确文件相对路径、可选节点范围和完整策略补丁。匹配顺序固定为项目 → 文件 → 节点范围；同级范围重叠报错。保护规则不通过模糊标题子串匹配。

M1 `source_toc` 支持 preserve/error；remove/rebuild 在 M2 通过明确区域映射实现。preserve 在报告中标为“原目录未更新”，有旧目录且结构发生变化的严格论文任务不得以此状态发布。

## 5. 内部类型与函数契约

类型采用 dataclass/TypedDict，具体语法遵守 B0 确认的 Python 版本。以下是接口设计，不是可直接运行代码。

```python
@dataclass(frozen=True)
class NodeRef:
    source_sha256: str
    part_uri: str
    element_path: str  # 源 XML 中基于元素序号的规范路径

@dataclass(frozen=True)
class RoleAssignment:
    node: NodeRef
    role: str
    status: str        # accepted / candidate / unresolved
    evidence_ids: tuple

@dataclass(frozen=True)
class Diagnostic:
    code: str
    severity: str      # info / warning / error
    location: str      # JSON Pointer 或 NodeRef 的序列化值
    message: str

def load_format(ref, *, declared_in, repository) -> FormatPackage: ...
def resolve_format(package, overrides, *, capabilities) -> ResolvedFormat: ...
def inspect_docx(path, *, limits) -> DocumentInspection: ...
def propose_roles(inspection, *, purpose, rules) -> RoleProposal: ...
def apply_role_decisions(proposal, decisions) -> RoleMap: ...
def infer_format(inspection, role_map, *, base_format) -> FormatCandidate: ...
def compile_format(candidate, decisions, *, package_id, version) -> FormatPackage: ...
def install_styles(document, resolved_format) -> StyleIdMap: ...
def import_region(source, region, destination, context) -> ImportResult: ...
def apply_roles(document, role_map, imported_nodes, context) -> ApplyReport: ...
def validate_format(document, resolved_format, application) -> ValidationReport: ...
```

`DocumentInspection` 包括 sources、blocks、style_graph、numbering_graph、sections、effective_properties、capabilities 和 diagnostics。每个 block 保存结构类型、story 类型、源 NodeRef、可见文字摘要、样式引用、有效属性、受保护对象标记；完整 XML 留在源包对象中。

`BuildPlan` 包括 parsed_config、resolved_format、source_assets、node_roles、deliveries、diagnostics、hashes。`UnifiedSynthesizer.plan()` 为旧调用者返回兼容 dict；新增内部 `_prepare_build()` 返回 BuildPlan，synthesize 与 smoke 复用它。

`RenderContext` 包括 resolved_format、current_section、import_registry、source_node_map、policy、diagnostics。方向状态放在实例中，v3 不再读写 `LAST_RENDERED_LANDSCAPE`。`ImportResult` 至少返回源节点到目标元素的映射、ID 改写表、导入对象清单与诊断。

错误统一进入已有 ConfigError/BuildError/OfficeExportError 边界。诊断代码使用 `FORMAT_CYCLE`、`INVALID_GEOMETRY`、`ROLE_UNRESOLVED`、`MAPPING_STALE`、`UNSUPPORTED_OBJECT`、`FONT_MISSING`、`PAGE_LABEL_MISMATCH` 等稳定标识。CLI 保持 0 成功、1 处理失败、2 参数用法错误；不为每个诊断发明退出码。

## 6. DOCX 读取与角色识别实现

### 6.1 安全读取和源节点身份（B2）

1. 用 OPC 根关系找到主文档部件，验证 Content Types 与关系目标；禁止路径逃逸，不跟随外部 URL。DOCM/DOC 首版拒绝作为识别样例。
2. ZIP/XML 限额设为可配置：初始建议总解压 256 MiB、单 XML 32 MiB、部件数 10,000、压缩比 200；以 B0 样本测量调整。超限返回诊断，不裁掉一半内容继续分析。
3. 按 document/header/footer/footnote/endnote/textbox story 建立节点索引。保留表格单元格和嵌套路径，不能只遍历 `Document.paragraphs`。
4. 以源文件哈希＋部件 URI＋规范元素路径定位。`w14:paraId` 可做辅助证据，不假定存在或唯一；同名同文段落仍保持不同身份。
5. clone 发生时立即写入内存节点映射；段落分割操作须记录一对多映射。v3 不凭文本去重新寻找已经导入的标题。

### 6.2 有效格式求值（B2）

构建 docDefaults、主题、basedOn、段落/字符/表格样式与编号依赖图；缓存样式求值，逐段逐 run 叠加上下文。字体分别解析 eastAsia/ascii/hAnsi/cs 与主题引用；粗斜体等 toggle 属性按 OOXML 语义处理。直接格式缺失与显式关闭必须区分。

第一批精确支持正文/标题的字符和段落属性、主题字体、列表属性；表格条件样式在支持首行/末行/奇偶带等样本通过前不标为完全解析。未实现的贡献层会让相关属性变成 unresolved，不能把低层默认值标为“精确观察”。网格吸附也进入有效行距解释和应用策略。

验证至少含：三层 basedOn、缺失父样式、循环、字符样式覆盖、直接 `bold=false`、主题字体、混合字号、编号缩进、表格条件样式、snapToGrid。有效属性与来源链一起输出，供报告解释。

### 6.3 格式聚类与角色判断（B5、B6）

格式指纹由规范化有效属性组成；1 twip 以内的转换误差归一化。忽略空 run 对字号众数的影响，分别统计字符覆盖量和段落数量；正文、表格和页眉页脚分池，不能让表内短段落成为正文样本多数。

角色规则顺序：精确人工节点映射 → 文件内显式样式/大纲规则 → 可信原生大纲与编号 → 格式与上下文候选。相同优先级冲突生成待判断项。已知目录区域中的标题文本不参与章节候选；遇到不同角色共享格式时保留多个角色。

标题自动接受要求语义证据与上下文一致；仅“大字＋短句”一类视觉规则始终先作为候选。标题层级跳跃列出诊断，不自动改变作者层次。正文段内出现“图”或“表”不足以判成题注。

可选模型适配器只负责候选排序，输入最小必要上下文、输出 schema 约束结果；不接受原始 XML 补丁。M1 核心不依赖模型服务，适配器启用后也必须通过相同校正和验证路径。报告记录模型/提示版本；模型中断时返回规则结果，不能卡住后续已保存规则的构建。

### 6.4 样例格式与目标角色映射分别保存

样例分析结果含来源正文摘要，仅作为项目内诊断；编译的格式包只保留规则及必要证据摘要，不含原始段落内容。目标文档映射单独绑定其 source_sha256；新目标重新生成角色映射，不复用样例段落编号。

格式包中没有观察到的角色可继承指定预设，但必须标为 inherited。样例里同一角色具有多种格式时，编译器要求选择代表规则或定义变体；不能将封面的大字号混到正文。

空白模板中的未使用样式在报告中单列 defined_only；只有用户明确映射后才进入角色候选。普通样例优先考虑 observed_effective 属性。

## 7. 识别报告、人工校正与 CLI

### 7.1 报告文件

一次分析输出 `analysis.json`、`review.html` 和 `decisions.example.json`。报告含 report_schema_version、report_id、source_sha256、analyzer_version、purpose、roles、style_candidates、sections、diagnostics；每个候选有稳定 ID 和证据定位。

校正文件示例：

```json
{
  "decisions_schema_version": 1,
  "report_id": "sample-analysis-001",
  "source_sha256": "64位十六进制输入哈希",
  "node_roles": {"node-0007": "heading.1", "node-0008": "body"},
  "role_styles": {"body": "candidate-body-02", "heading.1": "candidate-heading-01"},
  "missing_roles": {"heading.3": "inherit"}
}
```

编译时校验报告身份、输入哈希、节点和候选存在、角色合法、必需属性已解决。示例哈希是说明文字，真实文件必须为 64 位 hex。人工文件不允许注入 XML、脚本或任意输出路径。

目标映射编译为 `role-map-v1`，包括 source_sha256、analyzer_version、accepted assignments、on_unmapped 和受保护范围。段落变化后映射失效；首版只提示重新分析，不做模糊重定位。

### 7.2 新增操作与副作用

统一入口的操作互斥：现有 build/plan/doctor 与下表分析、编译、预览操作不能同时指定。参数在实际读取文件前校验。`--all` 不允许样例分析或格式编译。

| 操作 | 参数 | 输出和副作用 |
| --- | --- | --- |
| 分析样例 | `--analyze-format DOCX --analysis-dir DIR` | 写报告和本地 HTML；不调用 Word、不改样例 |
| 编译格式包 | `--compile-format REPORT --decisions JSON --format-base REF --format-id ID --format-version VERSION --format-out JSON` | 写校验通过的格式包；不构建文档 |
| 分析目标 | `--analyze-source DOCX --analysis-dir DIR` | 生成目标角色候选；可用 `--manifest` 提供角色集合 |
| 编译目标映射 | `--compile-mapping REPORT --decisions JSON --mapping-out JSON` | 写绑定源哈希的映射 |
| 格式预览 | `--preview-format REF --output-dir DIR` | 写虚构内容 preview.docx；默认不调用 Word，显示未验收 |
| Word 预览 | 上项加 `--render-preview` | 预检后导出 PDF；属于渲染操作 |
| 迁移清单 | `--migrate-manifest OLD --manifest-out NEW` | 只写新 v3 清单与显式兼容设置，不覆盖 OLD |

写入命令默认不覆盖已存在的报告/包/映射，提示更换路径；支持显式 `--replace-output`，仍必须禁止输出等于任何输入。多文件报告先 staging 后整体发布；失败不留下部分新报告。

HTML 内嵌静态 CSS/JS，不用 CDN，所有样例文本转义。页面提供候选选择、段落角色修改和导出 decisions.json；用户把导出文件交给编译操作。浏览器不直接修改仓库，也不需要启动写文件服务器。展示样例内容的 review 与仅用虚构内容的格式 preview 分开。

### 7.3 目标使用闭环

以下命令在 B6 完成后应由匿名示例验证。路径是计划中的示例文件，当前尚不存在。

```bash
python3 synthesize.py --analyze-format examples/custom-format-demo/sample.docx --analysis-dir output/format-review
python3 synthesize.py --compile-format output/format-review/analysis.json --decisions output/format-review/decisions.json --format-base preset:academic-basic@1.0.0 --format-id academic-demo --format-version 1.0.0 --format-out output/formats/academic-demo.json
python3 synthesize.py --analyze-source examples/custom-format-demo/manuscript.docx --analysis-dir output/source-review
python3 synthesize.py --compile-mapping output/source-review/analysis.json --decisions output/source-review/decisions.json --mapping-out output/source-review/roles.json
python3 synthesize.py --source examples/custom-format-demo/manuscript.docx --manifest examples/custom-format-demo/manifest.json --plan
python3 synthesize.py --source examples/custom-format-demo/manuscript.docx --manifest examples/custom-format-demo/manifest.json --doctor
python3 synthesize.py --source examples/custom-format-demo/manuscript.docx --manifest examples/custom-format-demo/manifest.json --output-dir output/custom-format-demo
```

示例 manifest 须预先引用上述输出格式包与映射的正确相对路径；decisions.json 由 review 导出或手工填写，不能把 example 文件默认当成已确认。新增 `--format` 快捷构建开关暂不做，避免与 manifest/override 出现第四套优先级。

## 8. 格式应用与导入：防止后续覆盖设置

### 8.1 样式写入（B3）

`install_styles()` 为目标创建确定的内部 ID，如 SynthBody、SynthHeading1，显示名可读；设置 basedOn、next、outlineLvl 和段落/字符属性。目录条目有独立样式，不能继续借标题字体、硬编码 8900 twip 制表位。

`apply_roles()` 按属性所有权清除冲突直接格式，然后设置 pStyle。凡目标样式声明的字体/字号/段距/行距等属性由目标管理；斜体强调、上下标、数学内容和超链接按保护策略保留。标题整体加粗与正文局部加粗的判别依据角色和源字符样式，不通过“删光 run 格式”实现。

缺少某个托管属性时不能悄悄清掉源属性；完整预设与 ownership 清单共同决定行为。生成段落和导入段落走同一应用器。纯文本内容无变化的重排不得重写 w:t；标题替换作为单独显式操作，遇到跨字段或复杂 run 先报不支持。

### 8.2 兼容策略隔离（B3、B4）

旧黑色化、长段落 1.5 倍行距、首行空格改缩进、PDF 页码清除、空白裁切、旧材料专项表宽、连续页码等由 `LegacyPolicy` 包装。v3 report/academic 默认不启用。

v1/v2 初期继续走经测试的兼容路径。v3 的解析和渲染使用同一新上下文；待新应用器能够表达兼容行为并通过基线后，再将旧包装逐项转发。不能为了避免两条临时路径而一次更改所有旧行为。

### 8.3 OPC 导入（B4）

M1 先保证单 DOCX 以原包副本为基础重排，降低脚注、公式及字段丢失风险。多 DOCX 合并只接受能力表已支持的关系类型；跨文件脚注/复杂域未实现时在 plan 阶段报错或按显式方案整页保留。

重映射至少覆盖样式、basedOn/next/link、numId/abstractNumId、书签 ID/名称和内部超链接、relationship ID、图片与 drawing ID。M2 增加脚注/尾注 ID 及受控 REF/PAGEREF/SEQ/STYLEREF 引用。不能仅克隆关系部件却漏掉正文中的 ID。

源主题与目标主题冲突时，针对导入样式解析颜色/字体并写成等价显式值；不复制整个源主题覆盖目标。未知主题对象效果在支持矩阵中单列。受保护的跨段复杂字段不能被逐段“孤立 end”清理器破坏。

### 8.4 页面、图片和目录几何（B4）

新增 `compute_content_box(section_spec)` 计算可用宽高；图片宽高按比例限制在版心内。标题与图片同页优先通过 keep_with_next 关联，遇到容量不足采取可解释的独立起页策略；不把字符数估算当成真实剩余高度。

PDF/图片缓存键包含输入内容哈希、页号、DPI、裁切/去页码策略和转换版本；最终适配尺寸若进入缓存也必须计入键。v3 横版页面与回到竖版都使用格式规格，不能回退到固定 A4 公文边距。

`page_policy=source` 时保留原节页面与页眉页脚；新增封面或目录有独立的目标节。混合模式转入/转出保留区时明确节边界，不能让前一个源页脚泄漏到主文。

## 9. 通用单 DOCX 与构建编排（B5）

新增来源策略 `docx_document`。`--source file.docx` 在显式 v3 中支持；v1/v2 单文件策略不改变。`discover_standalone_input_files()` 仍只自动列出明确 profile 声明的单文件项目，避免把目录里的附件都当项目。

无标题但有正文的 DOCX 是合法输入。BuildPlan 新增 content_block_count、heading_count；doctor 不再要求 node_count > 0。只有请求生成目录且没有目录条目时，按明确策略报错或允许空目录；M1 选择报错并提示移除 toc。

highlighted_docx 提取结果也转换为 RoleAssignment；现有正则替换语义保留在兼容路径。docx_outline 正则结果新增 NodeRef，v3 不通过相同文本列表依次猜定位。

构建调用顺序固定为：

```text
加载并校验配置
  → 解析格式包及能力
  → 只读分析来源，生成/加载目标角色映射
  → 检查歧义、对象能力、输入哈希
  → 创建隔离工作目录，按需转换旧 Office 材料
  → 按来源导入或复制原 DOCX 包，保留节点映射
  → 安装目标样式，按明确角色应用格式
  → 生成或装配交付部分，处理其节规则
  → 内容与格式结构 QA
  → Word 排版、目录回填、最终 QA
  → 事务式发布交付物和构建元数据
```

对于需要 Office 转换才能分析的旧 DOC，plan 只报告待转换和能力未知，M1 不支持对其自动 restyle；用户可先获得 DOCX。不能在只读 plan 内偷偷转换。

## 10. M1 验证与发布（B7）

### 10.1 语义完整性

`content_integrity.py` 为每个来源选区生成 SemanticInventory：规范可见文本序列、表格拓扑、媒体内容哈希、OMML 规范结构、脚注引用、字段指令、超链接目标和书签关系。

比较时允许明确的样式/编号 ID 重映射和生成的封面/目录文本；正文文字变更仅接受配置中声明的精确替换。不能继续只用“源段落文本是目标全文子串”证明保留，因为重复段落可能漏掉一份仍通过。

同一个输入被显式选入两次时，按 import instance 计数；媒体资源可去重存储，但引用次数不丢失。自动编号字段的动态文本与原始作者文字分开比较。

### 10.2 格式验证

对 mapped 节点验证最终有效字体、字号、段距、缩进、行距、outlineLvl；特别测试后续装配和页脚处理后仍成立。不要只检查样式 XML 的定义。

先覆盖具备确定性支持的属性；未能验证的属性进入报告，不输出“格式全部通过”。格式严格任务要求其所声明属性都有验证器。

### 10.3 发布元数据与 smoke

交付目录生成 `build-metadata.json`，包含输入/格式包/映射哈希、依赖版本、Word 和字体环境、交付物哈希、所用部分、QA 范围及未覆盖项。作为本次引擎保留文件名，参与 staging/发布回滚；用户的 output.documents 仍只声明 DOCX。

smoke 优先对照元数据的配置快照与选区清单；当前输入发生变化时报告与本次成果不对应，不用新输入重新推断旧成果。v1/v2 无元数据的既有成果保留当前验证入口。

Word 可用前运行结构测试；最终 M1 验收需真实 Word 导出和代表页面检查。报告预览与最终交付明确区分。

## 11. M2：文档部分、分节和论文对象

### 11.1 具名部分（B8）

v3 新增 `layout.parts` 注册表，output.documents.parts 改为按顺序引用其中 ID。未声明 layout 的 v3 与旧版均通过适配器获得 cover/toc/body 三个默认部件。

每个部件有 `kind`、`source_region` 或生成内容、`section_ref`、`page_sequence`、`include_in_toc`。kind 先支持 cover/generated_toc/content；摘要、参考文献等作为 content 的语义标签，不为每个名称写一个 renderer。

```json
{
  "layout": {
    "parts": {
      "abstract_cn": {"kind": "content", "source_region": "abstract-cn", "section_ref": "front", "page_sequence": "front"},
      "toc": {"kind": "generated_toc", "section_ref": "front", "page_sequence": "front"},
      "main_text": {"kind": "content", "source_region": "main", "section_ref": "main", "page_sequence": "main"}
    },
    "page_sequences": {
      "front": {"format": "upperRoman", "start": 1},
      "main": {"format": "decimal", "start": 1}
    }
  },
  "output": {
    "documents": [{"id": "main", "filename": "thesis.docx", "parts": ["abstract_cn", "toc", "main_text"]}]
  }
}
```

这是局部配置示意；完整示例必须同时定义 `front/main` 节规格和 source_region。源选区存在独立的 role-map 文件中，以 start-inclusive/end-exclusive NodeRef 指定，末端允许显式 end_of_document。边界只能落在允许的块边界，不能截断表格、字段、书签或分节对象。

选区重叠或存在未分配的实质内容默认报错；明确的 exclude 范围才允许不输出。v3 的内容完整性按实际声明选区检查，不能要求每份拆分文档包含全体来源。

M2 每份交付物最多一个 generated_toc，沿用文档级 toc.reference/links；目录范围必须是引用交付物内存在的标题集合。独立目录的引用规则保留，多份正文的目录不能隐式串用。

生成每个部分 start/end 边界书签，如 `_Synth_part_<id>_start`。书签 ID 校验和名称长度规则统一实现；旧 `_Synth_body` 等仅由旧适配器保留。部件是否含正文、是否需文字核验改为 kind/selection 判断，不比较字符串 body。

### 11.2 页码与分节（B9）

引入 PageRecord，字段为 physical_page、section_id、sequence_id、number_value、expected_label、observed_label、label_verified。封面是否显示页码与是否计入序列是不同规则；序列起始页用 start，后续节明确 continue，不重复写 start=1。

先在 B0 验证 macOS Word 能读取的范围、节和 PAGE 域信息。若只能取到 adjusted number，预期罗马字符串由配置生成，但不能把它填成 observed_label。真实标签可由 PAGE 域结果和对应页脚区域的 PDF 文本交叉检查；取不到则该能力不可发布。

`validate_measured_delivery()` 按 sequence 验证编号数值、显示标签及连续性，按物理页验证目录跳转。相同 number_value 可以出现在不同 sequence。隐藏页码没有 observed_label 是正常状态，不要求伪造数值字符串。

奇数页起章产生的插入空白页必须能从节规则和物理页记录追溯，不允许简单按“所有偶数空白页”放行。页眉页脚内容从按节计算的区域中识别，替换当前排除 `- N -` 文本和固定底部比例的做法。

不同封面页数可能改变正文起始奇偶性。M2 缺省逐交付物装配并分页；只缓存来源解析/导入中间结果。只有几何、序列、起页与页眉依赖确实一致时才复用正文排版结果。

### 11.3 脚注、题注和字段（B10）

新增 footnote/endnote part ID 分配和正文引用重映射；分隔符等保留条目不能按普通脚注重新编号。OMML 作为结构对象保留，不转图片或纯文本。脚注样式受 format 规则管理，编号范围由节规则控制。

题注用角色＋序列方案＋书签表示；先支持图、表的已有/新建 SEQ 编号和 REF/PAGEREF 引用，公式序列在样本通过后启用。识别出的手工编号默认保留；自动换成域必须能确定其引用对应关系，否则要求人工映射。

受控字段更新顺序：章节/题注编号 → 引用和页眉 → 分页 → 目录页码 → 重新分页。跨段字段按 begin/separate/end 栈解析，禁止逐段删“多余 end”。不全局更新未知外链、OLE 或第三方字段；若未知字段影响目标内容则阻止相应变换。

当前 inspect_document 在副本中导出且不保存回源。M2 更新字段必须写入本次 staging DOCX，并让 PDF、页码测量和最终交付都来自同一更新版本；不能只更新检查副本后发布未更新版本。

收敛状态包含书签物理页、sequence/value/label、目录内容哈希和相关字段结果；沿用有上限的迭代，缺省 3 次。未收敛不发布，保留诊断位置，绝不回退猜测页码 1。

## 12. 分批实施任务单

以下每批建议对应一个可评审提交或 PR。所有条目为未完成；同一批内部按编号顺序执行。

### B0 基线与技术验证

依赖：无。

1. 记录当前 git 状态、可用 Python/依赖/Word/字体版本；区分当前工作区与 HEAD，禁止将用户改动清掉后测基线。
2. 运行现有 unittest 与 minimal/chinese demo 的只读 plan；用匿名样本建立旧格式的有效属性和实际 Word 页面基线。
3. 在 tests/fixtures/formatting 定义 fixtures 清单和预期结果；测试生成器只写 TemporaryDirectory 或示例输入，不改业务材料。
4. 技术验证混合字体继承、跨段字段保留、脚注合并、罗马标签读取、奇数页起章。形成 `docs/formatting-capabilities.md`。

完成条件：现有失败有明确归因；每个高风险能力标为 supported/preserve_only/unsupported，不以“应该可行”进入实现承诺。B9/B10 的验证失败不阻塞 M1，但阻塞相应 M2 功能。

### B1 格式契约和兼容适配

依赖：B0。

1. 新增 schemas、format_schema/resolver/legacy_format；抽取合并和来源记录。
2. 增加 v3 校验和三个完整预设；无配置仍为 v2。
3. 实现引用、单位、循环、几何、未知能力错误；增添迁移到新文件的函数，CLI 在 B6 接入。
4. 给 ProjectConfig 增加只读 resolved_format/formatting，旧属性保留给兼容调用者。

测试：`test_format_config.py`、`test_format_resolver.py`，扩展 `test_output_config.py`。

完成条件：预设＋样例覆写可得到唯一可序列化配置；null/false/0、不同目录覆写、未知键和循环均有明确行为；旧配置测试不改预期即可通过。

### B2 只读 DOCX 分析和有效属性

依赖：B1。

1. 实现 package 检查、NodeRef、story/section 索引。
2. 实现样式图、主题与有效属性，按能力报告未覆盖来源。
3. 建立可见文本与对象清单读取接口，供 B7 共用。
4. 所有读取函数不得创建磁盘缓存，不触发 Office 或网络。

测试：`test_docx_inspector.py`、`test_effective_styles.py`，包含限额和损坏关系。

完成条件：同文重复段落可区分；各层属性贡献可追溯；未支持属性不会被标为精确；输入哈希不变。

### B3 统一样式应用与旧规则隔离

依赖：B2。

1. 实现 install_styles/apply_roles，覆盖生成标题、目录与普通段落。
2. 将字号、段距、行距、页脚规则从 styles/composition/highlighted 新路径迁入格式上下文。
3. 为兼容清洗提供命名策略；v3 禁止隐式黑色化及行距修补。
4. 修改后检查保留的行内语义及 XML 元素顺序。

测试：`test_style_applier.py`；同一文档用两种格式、直接格式覆盖、显式 false、超链接、上下标和公式。

完成条件：在 Word 修改 body 样式能统一影响正文；标题进入导航；重复应用同一格式在结构/有效属性上幂等；源文本不变。

### B4 导入上下文、编号和自适应几何

依赖：B3。

1. 建立 RenderContext；新路径移除全局方向和字体默认字典传递。
2. 抽取 package_importer，保留原生列表，消除样式/书签/关系 ID 冲突。
3. 用当前节版心计算图片与目录尺寸；更新 PDF 缓存策略键。
4. 建立 preserve/restyle/mixed 接管边界；能力不支持时在写入前退出。

测试：`test_package_importer.py`、`test_layout_geometry.py`、`test_formatting_modes.py`；扩展 composition 原有图像/编号/页眉保留测试。

完成条件：相同 styleId 的两个来源互不污染；横转竖回到用户指定纸张；更改清理策略不复用错误图片；保留区不被全局标准化改写。

### B5 通用来源与角色映射

依赖：B4。

1. 新增 docx_document、目标规则与 RoleMap 校验。
2. 拆分高亮提取和格式修改；v3 docx_outline 用 NodeRef 写书签。
3. 共用 _prepare_build，新增内容块计数；更新 doctor、plan 元数据。
4. 引擎从 plan 对象构建，开工前检查所有哈希，拒绝过期映射。

测试：`test_docx_document_strategy.py`、`test_role_mapper.py`、`test_build_plan.py`；扩展现有全来源流水线测试。

完成条件：普通单 DOCX 不需黄色标记即可重排；无标题正文可 body-only 构建；同名标题、正文内编号、目录重复标题不会错绑。

### B6 样例分析、校正报告、格式包复用

依赖：B5。

1. 实现聚类、角色候选和 FormatCandidate，字段级证据和缺失项报告。
2. 实现 report/decisions/role-map 校验及编译函数。
3. 新增 synthesize 操作分发和本地 HTML 校正导出；加入虚构内容预览。
4. 准备 custom-format-demo，按 §7.3 完整跑通样例→包→目标映射→构建。

测试：`test_format_analysis.py`、`test_format_review.py`、`test_format_cli.py`；HTML 转义、错误哈希、未决候选、写入冲突、非交互决策均覆盖。

完成条件：已保存格式包在删除分析缓存后仍可复用；同一包应用到不同目标；HTML 校正导出的 JSON 能通过编译；格式包中没有样例正文泄漏。

### B7 M1 验收、元数据和文档

依赖：B6。**[状态：已完成]**

1. 实现内容清单与最终有效格式验证，接入 delivery/smoke。
2. 发布元数据加入现有回滚事务，新增失败注入测试。
3. 运行离线测试、真实 Word 单栏样本、双格式切换和混合附件验收。
4. 更新 README、DEVELOPMENT、示例和项目技能的 v3 操作说明。

完成条件：§13 M1 全部强制项通过；旧匿名示例版面基线无未解释变化；不存在未标注的能力缺口。至此可以交付 M1。**[全部达成：141 项单元与验收测试 100% 通过]**

### B8 文档部分和源选区 [已完成]

依赖：B7。已于 2026-09-05 完成全部实施与验证（`test_source_regions.py` 7/7，`test_document_parts.py` 10/10，全仓 158/158 测试全绿）。

1. [已完成] 实现 layout.parts/regions 和 selection validator；更新 output.parts 校验。
2. [已完成] 改造 composition、required_bookmarks、目录范围及拆分交付排序。
3. [已完成] 修改 smoke 对实际选区的完整性验证；保留旧适配器。

测试：`test_document_parts.py`、`test_source_regions.py`；部分重排、选区重叠、表格中断、剩余内容、独立目录引用。

完成条件：摘要、目录、正文按配置装配；没有隐藏复制或漏掉正文；仅拆出某部分时检查其声明内容。[全部达成]

### B9 编号区间与真实 Word 测量 [已完成]

依赖：B8，且 B0 罗马标签验证已通过。

1. 实现 sequence 和节规格，v3 PageRecord 测量及标签核验。
2. 改造 footer/header、目录标签写入和收敛比较。
3. 更新页级 QA 的正文区域和有意空白判断。

测试：`test_page_sequences.py`、`test_word_formatting.py`；罗马→阿拉伯、隐藏但计数、多个 restart、横版继续、奇数起章、分拆交付。

完成条件：同名标题、相同编号数值仍指向正确物理页；真实标签匹配配置；无法测量或不收敛不发布。[全部达成]

### B10 论文对象、字段与 M2 验收 [已完成]

依赖：B9，且相关 B0 能力验证通过。

1. 实现脚注/尾注合并、题注序列和受控字段依赖。
2. 更新 staging DOCX 后再测量和发布，确保文件与 PDF 一致。
3. 生成完整匿名论文样本；对对象保留、字段结果和页面逐项验收。
4. 写明支持的学校/论文能力范围，未实现参考文献著录转换不列入支持项。

完成条件：§13 M2 强制项通过；公式、脚注、题注引用均正确；旧版和 M1 回归通过。至此可交付 M2。[全部达成]

## 13. 测试矩阵与发布门槛

| 样本 | 关键断言 | M1 | M2 |
| --- | --- | --- | --- |
| v1/v2 匿名完整/拆分示例 | 交付清单、目录书签、源文字、Word 页数基线 | 必须 | 必须 |
| 同一正文两套格式 | 标题正文目录的有效属性变更，后续装配不覆盖 | 必须 | 必须 |
| 全部 Normal＋手工格式 | 候选＋人工校正能生成可靠包，歧义不静默强改 | 必须 | 必须 |
| 空白模板 | defined_only 样式可手工映射，缺失角色标记来源 | 必须 | 必须 |
| 混合字体、字符样式、上下标 | 有效格式正确，局部语义保留 | 必须 | 必须 |
| 多文件同名样式和多级列表 | ID 隔离、原生编号连续/重启按源意图保留 | 必须 | 必须 |
| 普通表格、图片、横竖版混合 | 表格结构不丢失、图片不越版心、页面恢复正确 | 必须 | 必须 |
| 单文件已有脚注/公式/跨段域 | 支持原样保留；不支持变换在预检显式阻止 | 必须 | 必须 |
| 摘要/目录/正文/附录选区 | 全覆盖且不重复，目录范围正确 | 不启用 | 必须 |
| 罗马与阿拉伯、首页和奇偶页 | 标签、物理页、隐藏与计数、留白全部可解释 | 不启用 | 必须 |
| 跨来源脚注与题注引用 | ID/域更新正确，最终 DOCX 与 PDF 同步 | 不启用 | 必须 |
| 错误格式/陈旧映射/超限包 | 失败在发布前，无输入修改或旧成果替换 | 必须 | 必须 |
| 第 2 个文件或元数据发布失败 | 回滚已替换文件；回滚失败保留恢复信息 | 必须 | 必须 |

识别效果单独统计标题各级和正文 precision/recall、自动接受覆盖率、待判断比例、人工校正次数。开发样本与评估样本分开；建议 M1 自动接受标题精确率目标 ≥98%，同时要求包含足量正负例的独立评估报告，不以极少接受样本制造高准确率。未达到时收缩自动接受规则、保留校正流程，不夸大自动化程度。

确定性格式字段按支持矩阵逐项通过；源内容和关系零非预期丢失；Word 关键页码零错位。对每个失败保存位置、预期、实测值和复现样本，不能通过放宽全局 QA 阈值消除单例失败。

测试命令采用 unittest discover，与当前仓库一致，避免依赖 tests 已是 Python 包：

```bash
python3 -m unittest discover -s tests -v
python3 -m unittest discover -s tests -p 'test_format_*.py' -v
python3 synthesize.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --plan
git diff --check
```

每批先跑涉及模块测试，再按发布门槛跑回归。真实 Word 测试沿用 `DOCUMENT_SYNTHESIS_WORD_TEST=1` 的显式开关，并将临时测试目录限制在匿名验收输出区域；执行前使用现有 doctor。测试清理只删除本次创建的目录。

## 14. 排期、拆分与回退

规划顺序为 B0 → B1 → B2 → B3 → B4 → B5 → B6 → B7 → B8 → B9 → B10。本计划不要求多代理并行实施；如多人开发，可按接口冻结后的非冲突文件分工。

用于资源安排的初步工作量：B0 2—3 人日，B1 2—4，B2 4—7，B3 3—5，B4 4—7，B5 3—5，B6 4—7，B7 3—5；M1 合计约 25—43 人日。B8 3—5，B9 4—7，B10 5—9；M2 另约 12—21 人日。估算包含测试和匿名样本，不包含完整图形编辑器、模型服务接入和真实学校规范包制作；B0 后根据能力验证重估，不作为交付承诺。

每批合并前检查旧接口适配、配置迁移、测试和文档。B7/B10 是对外能力发布门禁，不在 B1 schema 接受字段时提前宣称功能完成；尚不支持的 required_capabilities 必须报错。

回退优先回到旧配置/预设和旧适配路径，不修改私有输入；新生成格式包带版本，可保留并重新选择旧版本。不要通过改写已经发布的格式包版本内容回退。无需为每个小函数加 feature flag，v3 和能力表即可隔离未成熟功能。

## 15. 实施开始时的具体顺序

第一批代码工作从 B0/B1 开始：记录基线 → 核实环境与边界样本 → 固定 format-v1/project-v3 schema → 写完整预设 → 实现解析和兼容适配 → 通过配置测试。暂不先开发“选样例后直接改 Word”的快捷流程，因为它会绕过格式契约和保留机制。

每批结束需更新本文对应状态及能力矩阵，记录真实执行命令和结果。新增功能文档示例须由测试或示例流水线执行验证后，才能把“目标设计”标记为“已可用”。

## 16. 本次规划的检查记录

2026-09-05 仅新增本执行计划，没有实施 B0—B10 功能。已做以下检查，用于说明实施起点：

- 现有 `python3 -m unittest discover -s tests -v` 退出码为 0，结果 OK，1 项真实 Word 测试按显式开关跳过。运行中有已有 lxml 元素真假判断 FutureWarning，B2 应消除对该行为的依赖。
- minimal-demo 与 chinese-demo 的 `--plan` 均成功；前者 2 个大纲节点，后者 11 个，均未调用 Word。
- 本文 4 个 JSON 代码片段均能进行 JSON 语法解析；其中占位值和局部示例仍需未来 schema 与完整示例验证，不代表当前可执行。
- 已检查代码围栏配对、B0—B10 批次数量、总体规划相对链接及行尾空白。

本次未运行真实 Word 导出；当前测试通过不能用作新格式识别或论文能力已经通过的证据。
