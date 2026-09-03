# 多源公文排版与材料合成引擎

一套用于处理申报、验收材料和多源印证文档的排版合成工具。

系统可将 PDF、多份 Word 文档、图片和演示文稿合并排版，并按公文版式统一字体、字号和间距。它通过本地 Word 导出 PDF，读取各级标题的实际页码，再生成带书签跳转的目录。

---

## 主要功能

1. **公文排版规范**
   采用 A4 规格，上白边 3.7 厘米、下白边 3.5 厘米、左白边 2.8 厘米、右白边 2.6 厘米。标题使用方正小标宋与黑体，正文使用三号仿宋，页脚居中插入页码。
2. **精准打印页码反查**
   先合成全量文档骨架，通过本地 Word 引擎排版并导出 PDF 逐页反查各级标题的物理页码，再把真实页码回填到目录控件中，生成带有点号引线和超链接跳转的目录。
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
├── synthesize.py                    # 统一主入口，支持合成指定子项目或全部项目
├── requirements.txt                 # Python 依赖清单
├── lib/                             # 排版引擎模块库
│   ├── config.py                    # 版本化配置契约与加载
│   ├── source_strategies.py          # 目录扫描、显式目录等来源策略
│   ├── highlighted_docx.py           # 高亮标题单文件策略
│   ├── styles.py                    # 公文样式、字体、版心与目录书签
│   ├── sanitizers.py                # 文本清洗、PDF 底部清理与图片裁切
│   ├── renderers.py                 # 多源材料统一渲染
│   └── qa.py                        # 本地 Word 渲染导出与页码反查
├── templates/                       # 可选的封面与目录模板，默认不提交二进制文件
├── manifests/                       # 通用项目清单（优先于 profile）
├── profiles/                        # 可复用的案例 profile
├── input/                           # 原始材料目录，内容由 .gitignore 忽略
├── output/                          # 合成成果输出目录
├── examples/minimal-demo/           # 可公开查看的匿名演示
├── .codex/skills/document-synthesis/ # 项目内操作 Skill
├── DEVELOPMENT.md                   # 开发者手册与架构说明
├── DEVELOPMENT_LOG.md               # 排版实现与技术说明
└── README.md                        # 项目使用说明
```

---

## 快速开始

### 1. 运行环境

- Python 3.8 以上版本
- 只有在 macOS 上通过 Microsoft Word 完成导出和页级检查，页码才可视为准确
- Windows 与 Linux 可用于其他处理，但不能把 LibreOffice 的版面当作 Word 的等价结果

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 匿名演示材料

仓库提供可公开查看的最小匿名示例，内容不含个人、单位、地点、联系方式或业务数据。先运行下面的只读验证命令。

```bash
python3 synthesize.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --plan
```

详细说明见 [examples/minimal-demo/README.md](examples/minimal-demo/README.md)。项目内 Skill 位于 [.codex/skills/document-synthesis/SKILL.md](.codex/skills/document-synthesis/SKILL.md)，用于引导后续项目采用统一入口、profile/manifest 和安全验证路径。使用封面模板时，请在模板中放置 `{{HEADER_TITLE}}`、`{{SUB_TITLE}}`、`{{MAIN_TITLE}}` 占位符。

### 3. 检查与预览

先检查环境和输入目录。下面两条命令均不会调用 Office 或改写材料。

```bash
# 检查 Python 依赖，并执行一次不读写文件的 Word Automation 探针
python3 synthesize.py --doctor

# 只读预览自动推断的目录、顺序和待转换文件
python3 synthesize.py --project 示例项目 --plan
```

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

构建会在 `output/<项目>/.work/` 中完成旧版 Office 文件转换、Word 页码读取和 PDF 页级检查。所有检查通过后，程序才会替换已有交付物。输入目录保持不变。无参数运行只显示帮助。

若 `--doctor` 报告“未获准控制 Microsoft Word”，请在 macOS“系统设置 → 隐私与安全性 → 自动化”中允许实际运行构建的宿主应用控制 Word。Full Disk Access 只处理文件访问，不能替代 Automation 权限；若报告“受限上下文”，请改从可访问 macOS 图形自动化服务的本机 Codex 或终端运行。

### 5. 自动化测试

运行测试脚本验证全部源材料完整性、文本保留情况与 Word 原生排版逐页质检（孤行、单标题、空白页）。

```bash
# 全量自动化质检
python3 smoke_test.py --all

# 测试指定板块
python3 smoke_test.py --project 示例项目
```

---

## 交付成果

编译完成后，各子项目的交付文档保存在 `output/` 对应子目录下。

| 交付物 | 作用 |
| --- | --- |
| `<项目>_合成材料.docx` | 用于 Word 物理页码反查的正文骨架，同时保留为可复核产物 |
| `<项目>_封面+目录.docx` | 独立红头封面与带跳转书签的精确目录 |
| `<项目>_目录+正文.docx` | 最终交付文档，目录无页码，正文从第 1 页连续编号 |

- 若 Word 无法导出 PDF，或最终 PDF 未通过空白页、单标题页和孤行检查，构建会失败，已有交付物不会被替换。

---

## 配置清单与来源策略

新项目优先复制 [manifests/manifest.example.json](manifests/manifest.example.json)，并命名为 `manifests/<项目名>.json`；项目私有配置也可放在 `input/<项目>/manifest.json`。所有配置都必须带 `schema_version: 1`，解析或校验失败会终止构建，绝不会悄悄改用默认规则。

配置按以下顺序查找。显式 `--manifest`、输入项目内 `manifest.json`、`profiles/<项目名>.json`、`manifests/<项目名>.json`、默认目录扫描。`profiles/` 用于收录已验证的案例，不应把一次性的项目判断写回主入口。项目私有的 `input/<项目>/manifest.json` 可覆盖 profile。

```json
{
  "schema_version": 1,
  "project_name": "示例项目",
  "source": { "strategy": "directory_tree" },
  "output": {
    "compiled_document": "示例项目_合成材料.docx",
    "cover_toc": "示例项目_封面+目录.docx",
    "toc_body": "示例项目_目录+正文.docx"
  },
  "tree_order": ["1建设方案", "2印证材料"],
  "node_overrides": {
    "2印证材料.docx": { "title": "二、印证材料" }
  }
}
```

`directory_tree` 是零配置目录扫描的默认策略。需要精确声明每一层材料、顺序和标题时使用 `explicit_tree`。单一 DOCX 用黄色高亮标记标题时使用 `highlighted_docx`。后者通过 `source.highlight` 配置颜色、标题正则和必要的标题替换，不再依赖任何项目专项脚本。`explicit_tree` 还可用 `docx_outline` 节点扫描某份 DOCX 的正文标题。`outline_rules` 用正则声明标题层级。合成时会在原段落写入书签，不会额外复制或重排正文。输出字段只能是输出目录内的 `.docx` 文件名，防止配置误写到项目外部。提交构建前先运行 `--plan` 审核目录。

项目只保留 `synthesize.py` 作为构建入口。历史专项脚本和仅供旧路径使用的渲染器已移除；新项目通过 profile 或 manifest 扩展，不新增专项脚本。

更多架构设计和排版细节可以参考 [DEVELOPMENT.md](DEVELOPMENT.md)。
