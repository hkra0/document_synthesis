# 开发说明

## 统一构建流程

`synthesize.py` 是唯一入口。构建先在 `output/<项目>/.work/run-*` 生成隔离产物，完成 Word 物理页码反查与 PDF 页级 QA 后，才原子发布三份 DOCX 交付物。

配置加载优先级为显式 `--manifest`、项目内 `manifest.json`、`profiles/<项目>.json`、`manifests/<项目>.json` 和默认目录扫描。

## 来源策略

- `directory_tree` 按目录和文件名扫描材料。
- `explicit_tree` 由 profile 声明树、顺序和标题。`docx_outline` 可将 Word 正文标题映射为目录书签。
- `highlighted_docx` 处理单一长篇 DOCX。高亮颜色、标题规则与替换项在 `source.highlight` 配置。

## 模块边界

- `lib/config.py` 提供版本化配置与安全校验。
- `lib/source_strategies.py` 解析材料来源并分配书签 ID。
- `lib/highlighted_docx.py` 提取高亮标题，标准化正文并处理连续页码。
- `lib/renderers.py` 负责 DOCX、PDF、PPTX 和图片的保真渲染。
- `lib/engine.py` 负责两趟编译、隔离工作目录和原子发布。
- `lib/qa.py` 负责 Word 自动化、页码读取和 PDF 页级断言。

## 验证

```bash
python3 -m unittest discover -s tests -v
python3 synthesize.py --project <项目> --plan
python3 synthesize.py --doctor --project <项目>
```

临时验证产物、Office 锁文件、`.work` 目录和 Python 缓存均不应提交；构建完成后可安全清理。
