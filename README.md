# 多源公文排版与材料合成引擎

一套用于处理申报、验收材料和多源印证文档的排版合成工具。

系统可将 PDF、多份 Word 文档、图片和演示文稿合并排版，并按公文版式统一字体、字号和间距。它通过本地 Word 导出 PDF，读取各级标题的实际页码，再生成带书签跳转的目录。

当前 post-review 前置基线见 [`docs/acceptance/post-review/status.json`](docs/acceptance/post-review/status.json)；N0–N10 的历史验收证据已完成，N7 独立匿名冻结集达到自动接受门槛；N9 已确认 Quartz PDF 的已知 orphan-xref 警告不影响当前页面语义，并保留原始诊断。当前补全批次以 [`docs/acceptance/remediation/status-summary.json`](docs/acceptance/remediation/status-summary.json) 为准：S0–S8 已在声明范围内验证；离线 325/325、NW01–NW12 12/12、`remediation_word` 11/11。S5 的真人审阅、格式包、目标 RoleMap 和另一份目标稿件的真实 Word build/smoke 证据见 [`s5-human-review-evidence.json`](docs/acceptance/remediation/s5-human-review-evidence.json)。NW01–NW12 使用固定 `DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=/private/tmp/document-synthesis-word-access`；不把 skip 计为通过。所有结论仅适用于声明的匿名夹具、冻结集和 Word 环境，不外推为任意真实文档的准确率保证。

---

## 主要功能

1. **公文排版规范**
   采用 A4 规格，上白边 3.7 厘米、下白边 3.5 厘米、左白边 2.8 厘米、右白边 2.6 厘米。标题使用方正小标宋与黑体，正文使用三号仿宋，页脚居中插入页码。
2. **精准打印页码反查**
   正文只渲染一次，再按交付清单装配文档。通过本地 Word 读取书签所在的物理页与打印编号，回填目录，并核验最终装配文档；目录回填改变分页时进行有上限的重排。
3. **段落属性分页控制**
   使用段落属性 `page_break_before` 控制大章节换页，避免在文档末尾插入独立换页符时出现多余空白页。
4. **页面底部旧页码清除与高分辨率渲染**
   使用 PyMuPDF 定位页面底部旧页码区域并进行小范围覆盖清除，保留正文与表格底线，页面按 300 DPI 渲染插入。
5. **Word 文档结构与样式合并**
   合并源 Word 文档的段落结构与样式定义，清除默认的主题颜色，支持横版表格分节与连续编页。
6. **配置清单驱动**
   材料结构与排版代码分离，通过修改 JSON 清单即可调整材料顺序与排版参数。

---

## 项目结构

```text
document_synthesis/
├── synthesize.py                    # 统一主入口，支持排版合成与格式样例分析/编译
├── smoke_test.py                    # 自动化交付检查与语义完整性核验
├── requirements.txt                 # Python 依赖清单
├── schemas/                         # 规范契约（format-v1, project-v3）
├── formats/presets/                 # 内置预设格式包（公文、学术论文、报告规范）
├── lib/                             # 排版引擎模块库
│   ├── config.py                    # 版本化配置契约（v1/v2/v3）与加载
│   ├── format_schema.py             # 格式包对象模型与单位几何校验
│   ├── format_resolver.py           # 格式包继承、解析与能力匹配
│   ├── docx_inspector.py            # 只读 OOXML 检查与有效属性求值
│   ├── style_applier.py             # 统一样式应用引擎与行内语义保护
│   ├── layout.py                    # 版面几何与 ContentBox 排版度量
│   ├── package_importer.py          # 段落导入、媒体与原生编号重映射
│   ├── role_mapper.py               # 语义角色识别与 NodeRef 映射
│   ├── format_analysis.py           # 样例分析与样式聚类
│   ├── format_review.py             # 本地单文件 HTML 校正与格式包编译
│   ├── content_integrity.py         # 段落多重集内容完整性与有效格式核验
│   ├── source_strategies.py         # 目录扫描、显式大纲与标准单 DOCX 策略
│   ├── highlighted_docx.py          # 高亮标题单文件策略
│   ├── composition.py               # 封面、目录、正文的统一装配
│   ├── delivery.py                  # 交付物结构、目录与页级验证
│   ├── pagination.py                # Word 书签的物理页与打印编号
│   ├── styles.py                    # 公文样式、字体、版心与目录书签
│   ├── sanitizers.py                # 文本清洗、PDF 底部清理与图片裁切
│   ├── renderers.py                 # 多源材料统一渲染
│   ├── qa.py                        # 本地 Word 渲染导出与页码反查
│   └── engine.py                    # 统一合成调度、元数据审计与原子事务回滚
├── templates/                       # 可选的封面与目录模板，默认不提交二进制文件
├── manifests/                       # 通用项目清单（profile 未命中时使用）
├── profiles/                        # 可复用的案例 profile
├── input/                           # 原始材料目录，内容由 .gitignore 忽略
├── output/                          # 合成成果输出目录
├── examples/custom-format-demo/     # 自定义格式提取与规范排版完整示例
├── examples/minimal-demo/           # 可公开查看的匿名演示
├── .codex/skills/document-synthesis/ # 项目内操作 Skill
├── DEVELOPMENT.md                   # 开发者手册与架构说明
├── DEVELOPMENT_LOG.md               # 排版实现与技术说明
└── README.md                        # 项目使用说明
```

---

## 快速开始

### 1. 运行环境

- Python 3.10 以上版本
- macOS + Microsoft Word：通过 AppleScript 完成导出和页级检查，页码已核验
- Windows + Microsoft Word：通过 COM（pywin32，随 `requirements.txt` 在 Windows 上自动安装）完成导出和页级检查。该后端已在 Windows 11 + Microsoft 365 Word（64 位）上跑通三个示例的完整构建；Windows 真实 Word 验收矩阵（P4）完成之前，视为预览支持，页码只对生成它的 Word 环境有效
- 没有 Word 的 Windows 与 Linux 可用于其他处理，但不能把 LibreOffice 的版面当作 Word 的等价结果
- Windows 上把文中的 `python3` 换成 `python`（或 `py`）

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 匿名演示材料

仓库提供可公开查看的最小匿名示例，内容不含个人、单位、地点、联系方式或业务数据。先运行下面的只读验证命令。

```bash
python3 synthesize.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --plan
```

中文展示示例见 [examples/chinese-demo/README.md](examples/chinese-demo/README.md)，包含可编辑的 3∶4 竖版封面、A4 中文正文及自动目录。最小匿名测试示例见 [examples/minimal-demo/README.md](examples/minimal-demo/README.md)。项目内 Skill 位于 [.codex/skills/document-synthesis/SKILL.md](.codex/skills/document-synthesis/SKILL.md)，用于引导后续项目采用统一入口、profile/manifest 和安全验证路径。使用封面模板时，可放置 `{{HEADER_TITLE}}`、`{{SUB_TITLE}}`、`{{MAIN_TITLE}}` 占位符替换内容，也可保留固定文案。

### 3. 检查与预览

先检查环境和输入目录。`--plan` 不调用 Office；`--doctor` 会执行 Word Automation 只读探针。两者均不改写材料。

```bash
# 检查 Python 依赖，并执行一次不读写文件的 Word Automation 探针
python3 synthesize.py --doctor

# 只读预览自动推断的目录、顺序和待转换文件
python3 synthesize.py --project 示例项目 --plan
```

`--doctor` 返回标准退出码：
- `0`：环境完全就绪，支持 `exact` 精确构建（Python 依赖齐全、输入合法、Office 探针通过）。
- `2`：Python 依赖就绪，但无可用精确 Office 自动化（例如 Linux，或没有 Word 的 Windows / macOS）；可运行 `--plan`、格式分析类命令和 `smoke_test.py --structure-only`，不能构建带目录页码的交付物。
- `1`：缺少 Python 核心依赖（`docx`、`pymupdf` 等）、输入无效，或 Office 后端配置无效。

可用 `--office-backend {auto,word,none}` 或环境变量 `DOCUMENT_SYNTHESIS_OFFICE_BACKEND` 指定后端；命令行参数优先。在 Windows 上，`--doctor` 还会报告 Word 版本与位数、安装方式，以及 PowerPoint 是否可用；没有 PowerPoint 时只有 `.pptx` 转换不可用。

Windows 上每次 Office 操作都在独立子进程中启动隐藏的 Word 实例，默认 600 秒超时（可用环境变量 `DOCUMENT_SYNTHESIS_OFFICE_TIMEOUT` 调整）。超时或失败时只结束本次启动的 Word 进程，不影响你自己打开的 Word 文档。

### 4. 合成

通过主入口 `synthesize.py` 调度全部或指定子项目（支持直接使用子项目名称或模糊匹配）。

```bash
# 合成 input/ 下的全部项目
python3 synthesize.py --all

# 合成指定板块材料
python3 synthesize.py --project 示例项目

# 单文件项目也由 profile 自动识别，无需项目名称特判
python3 synthesize.py --project 单文件示例
python3 synthesize.py --source input/单文件示例.docx

# 支持模糊子串匹配
python3 synthesize.py --project 示例
```

构建会在 `output/<项目>/.work/` 中完成旧版 Office 文件转换、Word 页码读取和 PDF 页级检查。所有声明的文档检查通过后，程序才会发布交付物；QA 失败时已有交付物不会被替换。输入目录保持不变。无参数运行只显示帮助。默认清理中间产物，需要排查正文骨架或 PDF 时显式加 `--keep-work`，诊断后自行清理对应工作目录。

### 4.1 自定义格式包制作与样例识别工作流 (v3)

支持从已有 Word 规范样例中逆向推导排版格式，生成无正文泄露的标准格式包，并应用于新稿件排版：

```bash
# 所有生成物放入一个可删除的工作目录；以下命令从仓库根目录执行
export R11_WORK=output/custom-format-demo

# 1. 分析格式样例，生成 analysis.json、review.html 和 decisions.example.json
python3 synthesize.py --analyze-format examples/custom-format-demo/sample.docx \
  --analysis-dir "$R11_WORK/review" --replace-output

# 2. 在浏览器打开 review.html 完成人工校正；处理所有未决项、至少改变一项决定并勾选“我确认已人工审阅并核对上述决定”，再用带 review_audit 的最终 decisions.json 编译格式包
python3 synthesize.py --compile-format "$R11_WORK/review/analysis.json" \
  --decisions examples/custom-format-demo/decisions.final.json \
  --format-base preset:academic-basic@1.0.0 \
  --format-id academic-demo --format-version 1.0.0 \
  --format-out "$R11_WORK/formats/academic-demo.json" --replace-output

# 2.1 生成虚构内容的近似预览；预览不包含样例正文
python3 synthesize.py --render-preview "$R11_WORK/formats/academic-demo.json" \
  --preview-out "$R11_WORK/preview.html" \
  --preview-meta "$R11_WORK/preview-metadata.json" \
  --preview-engine html-css-approximate --replace-output

# 3. 分析待排版稿件，并用同一份最终决策编译绑定源哈希的 RoleMap
python3 synthesize.py --analyze-source examples/custom-format-demo/manuscript.docx \
  --analysis-dir "$R11_WORK/target-review" --replace-output
python3 synthesize.py --compile-mapping "$R11_WORK/target-review/analysis.json" \
  --decisions examples/custom-format-demo/manuscript-decisions.final.json \
  --mapping-out "$R11_WORK/review/manuscript-roles.json" --replace-output

# 3.1 将旧 v1 清单迁移到独立 v2 文件，不覆盖输入
python3 synthesize.py --migrate-manifest examples/custom-format-demo/legacy-manifest-v1.json \
  --manifest-out "$R11_WORK/migrated/manifest-v2.json" --replace-output

# 4. 使用 v3 清单执行 plan、doctor、构建和 smoke 审计
python3 synthesize.py --source examples/custom-format-demo/manuscript.docx \
  --manifest examples/custom-format-demo/manifest.json --plan
python3 synthesize.py --source examples/custom-format-demo/manuscript.docx \
  --manifest examples/custom-format-demo/manifest.json --doctor
python3 synthesize.py --source examples/custom-format-demo/manuscript.docx \
  --manifest examples/custom-format-demo/manifest.json \
  --output-dir "$R11_WORK/deliverables"
python3 smoke_test.py --source examples/custom-format-demo/manuscript.docx \
  --manifest examples/custom-format-demo/manifest.json \
  --output-dir "$R11_WORK/deliverables"
```

完整示例可直接参考 [examples/custom-format-demo/README.md](examples/custom-format-demo/README.md)。

格式包还可以使用 `--render-preview` 生成虚构内容的近似预览，使用 `--migrate-manifest` 将 v1/v2 清单输出到独立的新文件；两者都不会把样例正文写入格式包，也不会覆盖输入文件。所有命令的参数以 `python3 synthesize.py --help` 为准。

若 `--doctor` 报告“未获准控制 Microsoft Word”，请在 macOS“系统设置 → 隐私与安全性 → 自动化”中允许实际运行构建的宿主应用控制 Word。Full Disk Access 只处理文件访问，不能替代 Automation 权限；若报告“受限上下文”，请改从可访问 macOS 图形自动化服务的本机 Codex 或终端运行。

Automation 探针成功不代表 Word 已能读写工作目录。若导出出现 “Grant File Access”，需由用户确认并授予所需输出目录的访问权限；未完成导出与最终 QA 前，不能声称分页验收通过。

### 5. 自动化测试

运行测试脚本验证全部源材料完整性、文本保留情况与 Word 原生排版逐页质检（孤行、单标题、空白页）。

```bash
# 全量自动化质检
python3 smoke_test.py --all

# 测试指定板块
python3 smoke_test.py --project 示例项目

# 无 Office 环境只做结构和文字核验，不代表 Word 页码与版式通过
python3 smoke_test.py --project 示例项目 --structure-only
```

检查器使用与构建相同的来源策略和交付清单，支持 `--source`、`--manifest` 和可重复的 `--override`。检查拆分交付时也应传入构建时使用的覆写文件。单项目的 `--output-dir` 指实际交付目录，`--all` 时指输出根目录。

### 6. 当前能力边界与验收证据

历史 R7–R11 记录保留的证据包括：R7 独立匿名评估集 20 份（100 个标题正例、500 个正文负例）；R8/R9 定向回归 54/15 项；R10 离线全套 246 项（2 项条件跳过）、M1 真实 Word 3/3，以及 W04–W07 真实 Word/PDF 4/4；R11 自定义格式示例的分析、决策、编译、预览、RoleMap、迁移、plan、doctor、真实 Word build 和完整 smoke 均已通过。当前发布结论以 [`docs/acceptance/post-review/status.json`](docs/acceptance/post-review/status.json) 为准；本轮 post-review 的 NW01–NW12 真实 Word 矩阵已完成，N9 的 orphan-xref 诊断规则已通过且不会改写原始 PDF。

当前声明不覆盖 OCR/PDF 样例学习、双栏期刊、参考文献著录转换、多样例融合或全部学校规范认证。保留/重排边界、测试 ID 和最近证据见 [能力矩阵](docs/formatting-capabilities.md)、[R10 验收状态](docs/acceptance/r10-status.json) 与 [R11 验收状态](docs/acceptance/r11-status.json)。

---

## 交付成果

新项目默认只生成 `<项目>_完整文档.docx`，包含封面、目录和正文，保存在 `output/<项目>/`。封面与目录不显示页码，正文从 1 开始。正文骨架是内部构建产物，不再默认交付。

`output.documents` 声明本次完整交付清单。每份文档从 `cover`、`toc`、`body` 按该顺序选择，每个部分最多一次。省略清单才采用完整文档默认；显式清单整体替换默认，空清单、重复 ID、文件名冲突或无效目录引用均报错。是否包含封面由 `parts` 决定，`cover.template: false` 仍表示使用内置封面。

带正文的目录默认引用自身，使用本文件书签跳转。独立目录必须用 `toc.reference` 指向含正文的交付文档，`toc.links` 使用 `none`，其页码以被引用文档的正文打印编号为准，不产生无效的本文件链接。

这里的目录回填与页码 QA 仅针对引擎生成的交付目录。源 Word 文件内部已有的目录、原生 TOC 域及其缓存页码属于源内容，本流程不会统一重建或验证它们的页码；正文重新排版后，不能把这些内部目录也视为已经验收。

旧 `schema_version: 1` 配置暂时保留正文骨架、封面＋目录、目录＋正文三份交付并提示迁移。迁移至 v2 时，删除旧 `output` 字段即可使用完整文档默认，或改为 `output.documents` 明确声明。旧成果不会按通配符自动删除；计划会提示输出目录中不属于本次清单的 DOCX，确认不再需要后再自行归档。

在 `schema_version: 3` 工程中，交付目录会同时输出 `build-metadata.json` 构建元数据文件。该文件记录了生成器版本、构建时间戳、格式包 ID/版本/能力、配置契约 SHA-256、源文件 SHA-256 快照、交付 DOCX 的 SHA-256 以及内容语义完整性校验状态，用于长期归档与防篡改审计。`smoke_test.py` 会重新计算并核对这些哈希，不把历史 `passed` 字段当作当前验证结果。旧版本 v1/v2 工程不输出此文件，保持既有交付目录脚印完全不变。

---

## 配置清单与来源策略

新项目优先复制 [manifests/manifest.example.json](manifests/manifest.example.json)，并命名为 `manifests/<项目名>.json`；项目私有配置也可放在 `input/<项目>/manifest.json`。新配置使用 `schema_version: 2`；版本 1 仍兼容。解析或校验失败会终止构建，绝不会悄悄改用默认规则。

基础配置按以下顺序选择一份，不隐式叠加。显式 `--manifest`、输入项目内 `manifest.json`、`profiles/<项目名>.json`、`manifests/<项目名>.json`、默认目录扫描。`profiles/` 用于收录已验证的案例，不应把一次性的项目判断写回主入口。项目私有的 `input/<项目>/manifest.json` 会整份替代 profile。

```json
{
  "schema_version": 2,
  "project_name": "示例项目",
  "source": { "strategy": "directory_tree" },
  "tree_order": ["1建设方案", "2印证材料"],
  "node_overrides": {
    "2印证材料.docx": { "title": "二、印证材料" }
  }
}
```

上述配置无需声明输出，即生成一份完整文档。需要拆分交付时，可在基础配置中增加以下 `output`，也可将其单独保存为覆写文件：

```json
{
  "output": {
    "documents": [
      {
        "id": "front",
        "filename": "示例项目_封面+目录.docx",
        "parts": ["cover", "toc"],
        "toc": {"reference": "main", "links": "none"}
      },
      {
        "id": "main",
        "filename": "示例项目_目录+正文.docx",
        "parts": ["toc", "body"]
      }
    ]
  }
}
```

```bash
python3 synthesize.py --project 示例项目 --override overrides/split.json --plan
```

`--override` 可重复，按命令行顺序应用于选中的基础配置。对象递归合并，标量替换，数组整体替换；`null` 不表示删除字段，合并后统一校验。切换 v1 到 v2 应直接迁移基础配置，移除旧输出字段，不能只叠加 `schema_version: 2` 后保留旧键。`--plan` 显示基础配置路径、覆写来源、最终交付清单、目录引用与编号规则，不调用 Word。

`directory_tree` 是零配置目录扫描的默认策略。需要精确声明每一层材料、顺序和标题时使用 `explicit_tree`。单一 DOCX 用黄色高亮标记标题时使用 `highlighted_docx`。后者通过 `source.highlight` 配置颜色、标题正则和必要的标题替换，不再依赖任何项目专项脚本。`explicit_tree` 还可用 `docx_outline` 节点扫描某份 DOCX 的正文标题。`outline_rules` 用正则声明标题层级。合成时会在原段落写入书签，不会额外复制或重排正文。输出字段只能是输出目录内的 `.docx` 文件名，防止配置误写到项目外部。提交构建前先运行 `--plan` 审核目录。

项目只保留 `synthesize.py` 作为构建入口。历史专项脚本和仅供旧路径使用的渲染器已移除；新项目通过 profile 或 manifest 扩展，不新增专项脚本。

更多架构设计和排版细节可以参考 [DEVELOPMENT.md](DEVELOPMENT.md)。
