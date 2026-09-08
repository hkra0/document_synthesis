# Custom Format Demo（自定义格式提取与重排）

本目录提供一个匿名端到端示例：从参考 DOCX 提取格式候选，经过明确的人工决策生成版本化格式包，再分析另一份目标 DOCX、生成绑定源哈希的 RoleMap，最后由统一入口执行计划、环境检查、构建和 smoke 检查。样式分析、编译、映射、预览、迁移和 `--plan` 的离线复核记录在 [`docs/acceptance/post-review/offline-verification.json`](../../docs/acceptance/post-review/offline-verification.json)；构建与 smoke 仍需要 Microsoft Word Automation。

`sample.docx` 与 `manuscript.docx` 只含虚构文本。`decisions.final.json` 和 `manuscript-decisions.final.json` 是已经完成的示例决策，分别绑定两份 DOCX 的 SHA-256；如果替换源文件，编译器会拒绝使用旧决策，必须重新分析并校正。

## 从零运行

以下命令在仓库根目录执行。示例自带不含正文的格式包和角色映射，因此 `--plan`、`--doctor` 和离线构建不依赖历史 `output/` 目录。分析、预览、迁移和 Word 交付仍写入匿名的 `output/custom-format-demo/`，验证后可以清理。

```bash
export R11_WORK=output/custom-format-demo
```

### 1. 分析样例并完成校正

第一条命令生成 `analysis.json`、自包含的 `review.html` 和带 `pending` 标记的 `decisions.example.json`。在浏览器打开 `review.html`，检查候选角色、字号和对齐方式，导出或保存最终决策。仓库提供的 `decisions.final.json` 是本示例的已确认版本。

```bash
python3 synthesize.py --analyze-format examples/custom-format-demo/sample.docx \
  --analysis-dir "$R11_WORK/review" --replace-output

python3 synthesize.py --compile-format "$R11_WORK/review/analysis.json" \
  --decisions examples/custom-format-demo/decisions.final.json \
  --format-base preset:academic-basic@1.0.0 \
  --format-id academic-demo --format-version 1.0.0 \
  --format-out "$R11_WORK/formats/academic-demo.json" --replace-output
```

`decisions.example.json` 仅是草稿，仍含未决角色，不能直接编译。编译输出不包含样例正文，只包含格式、角色和页面规则。

仓库内的 `academic-demo.format.json` 与 `manuscript-roles.json` 是本示例的冻结、可复用结果，便于直接运行第 4 步；若重新分析样例或目标稿件，应把新结果写入 `$R11_WORK`，并在确认后替换这两个文件。

### 2. 生成虚构预览和迁移输出

预览不会读取样例正文，输出会标明近似 HTML 引擎并写入可审计元数据。迁移命令把旧 v1 清单写到独立的 v2 文件，不覆盖输入文件。

```bash
python3 synthesize.py --render-preview "$R11_WORK/formats/academic-demo.json" \
  --preview-out "$R11_WORK/preview.html" \
  --preview-meta "$R11_WORK/preview-metadata.json" \
  --preview-engine html-css-approximate --replace-output

python3 synthesize.py --migrate-manifest \
  examples/custom-format-demo/legacy-manifest-v1.json \
  --manifest-out "$R11_WORK/migrated/manifest-v2.json" --replace-output
```

### 3. 分析目标并编译 RoleMap

目标分析有自己独立的报告和源哈希。最终 RoleMap 只绑定 `manuscript.docx` 的稳定 NodeRef；`on_unmapped: error` 让构建在出现未确认节点时停止。

```bash
python3 synthesize.py --analyze-source examples/custom-format-demo/manuscript.docx \
  --analysis-dir "$R11_WORK/target-review" --replace-output

python3 synthesize.py --compile-mapping "$R11_WORK/target-review/analysis.json" \
  --decisions examples/custom-format-demo/manuscript-decisions.final.json \
  --mapping-out "$R11_WORK/review/manuscript-roles.json" --replace-output
```

### 4. 计划、环境检查、构建和 smoke

`--plan` 只读且不调用 Office；正式构建前必须先执行 `--doctor`。构建使用本目录的 v3 manifest，其中的格式包和 RoleMap 路径指向本目录冻结的无正文结果。

```bash
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

构建需要本机 Microsoft Word Automation；`--doctor` 或构建提示受限上下文时，应改在获 Automation 权限的本机桌面终端运行。`Grant File Access` 只授权当前 Word 所需的临时目录，不能把一次授权描述为永久权限。
