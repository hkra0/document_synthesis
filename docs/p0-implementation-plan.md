# P0 批次实施计划与执行记录：发布前清理与基线固定

- **编制日期**：2026-10-09
- **所属计划**：[Windows 支持与公开发布执行计划](windows-and-public-release-plan.md)
- **当前状态**：**已完成 (Verified)**
- **执行原则**：每项动作均以“契约断言 → 实现/改写 → 对照回归 → 证据固定”执行；不把 `blocked` / `partial` 提前写成 `verified`；源文件 SHA-256 绝不篡改。

---

## 1. 批次目标与退出条件

### 1.1 核心目标
1. **彻底消除路径与隐私敏感信息（G1）**：将被跟踪证据文件与状态脚本中的开发机绝对路径规范化改写为 `<repo>/...`，将临时访问路径改写为 `<word-access-dir>/...`，消除个人邮箱及用户目录痕迹。
2. **源头防回归与标准化（G1/G2）**：统一抽取 `relativize()` 与 `sanitize_text()` 规范化函数，供现有及后续证据脚本统一调用；编写 `tests/test_repo_hygiene.py` 自动化检测防护。
3. **跨平台换行与文件分类（R6）**：引入 `.gitattributes`，统一文本文件换行规则（`eol=lf`）与二进制文件声明，消除跨平台 CRLF/LF 漂移。
4. **工作区环境清理与忽略规则加固（G9）**：清除本地未跟踪的大型临时文档产物，在 `.gitignore` 中补充模式规则。
5. **冻结当前基线能力**：在干净的工作树下运行完整测试套件与 macOS Microsoft Word 验收矩阵，将基线结果固化至 `docs/acceptance/public-release/p0-baseline.json`。

### 1.2 退出条件对照表

| 序号 | 退出条件 | 验收方法 | 状态 | 验证结果 |
| :--- | :--- | :--- | :--- | :--- |
| C1 | `test_repo_hygiene` 全部通过 | `python3 -m unittest tests.test_repo_hygiene -v` | **已通过** | 4 项用例全部通过（0.295s） |
| C2 | `git grep -nE '/(?:Users\|home)/'` 无任何匹配 | 运行 git grep 命令返回空（exit code 1） | **已通过** | 0 项匹配，返回码 1 |
| C3 | 证据脱敏前后源文件 SHA-256 与业务数据无变更，仅路径字段规范化 | 比对 git diff 摘要与 metadata 文件 | **已通过** | 仅路径前缀改为 `<repo>` / `<word-access-dir>`，哈希值严格不变 |
| C4 | `evidence-hashes.json` 与脱敏后文件哈希严格一致 | 自动化校验 3 份哈希索引条目 | **已通过** | 全部文件 SHA-256 重新计算并精确吻合 |
| C5 | 脱敏与清理后，离线测试维持基线水准（新增卫生检查） | `python3 -m unittest discover -s tests -v` | **已通过** | 329 项通过（含 4 项新测验）、6 项条件跳过（38.264s） |
| C6 | 固化真实 Word 运行结果与离线测试基线至 `p0-baseline.json` | 检查基线 JSON 完整性与可读性 | **已通过** | 写入 `docs/acceptance/public-release/p0-baseline.json` |

---

## 2. 任务分解与执行步骤

### Task 1: 建立脱敏工具模块 (`docs/acceptance/tools/sanitize_paths.py`)
- **目标**：提供可复用的路径规范化函数 `relativize(path, repo_root)`、文本脱敏函数 `sanitize_text(text, repo_root)` 以及批量文件脱敏与哈希重算 CLI。
- **规范化规则**：
  - 仓库根绝对路径 -> `<repo>/<relpath>`
  - `/private/tmp/document-synthesis-word-access` -> `<word-access-dir>`
  - `/private/tmp/...` / `/tmp/...` -> `<word-access-dir>/...`
  - 用户下载目录（如 `.../Downloads/...`）-> `<downloads>/...`
  - 其他个人主目录 -> `<user>/...`
- **CLI 能力**：支持 `--check`（只读巡检）与 `--write`（执行替换并重算 `evidence-hashes.json`）。

### Task 2: 被跟踪证据文件路径脱敏改写
- **范围**：
  - `docs/acceptance/n0-n10-review/`（metadata、diagnostics、stdout/stderr、report）
  - `docs/acceptance/post-review/`（pdf-diagnostics、word-evidence、n*-status）
  - `docs/acceptance/r0-r11-review/`（metadata）
  - `docs/acceptance/r8-status.json` ~ `r11-status.json`
  - `docs/acceptance/remediation/`（s*-status、s5 评估证据）
  - `docs/acceptance/word-followup/`（academic-metadata）
- **约束**：仅替换路径字段，源文件 SHA-256、测试度量指标保持原样。

### Task 3: 证据生成脚本接入统一 `relativize()`
- **涉及脚本**：
  - `docs/acceptance/post-review/pdf-diagnostics/diagnose_pdfs.py`
  - `docs/acceptance/post-review/write_status.py`
  - `docs/acceptance/remediation/write_status.py`
- **改动**：引入 `from docs.acceptance.tools.sanitize_paths import relativize, sanitize_text`，生成报告和 sidecar 时杜绝再次写入绝对路径。

### Task 4: 证据哈希索引文件同步更新
- **涉及索引**：
  - `docs/acceptance/n0-n10-review/evidence-hashes.json`（更新 9 项）
  - `docs/acceptance/r0-r11-review/evidence-hashes.json`（更新 6 项）
  - `docs/acceptance/word-followup/evidence-hashes.json`（更新 8 项）
- **操作**：对改写后的证据文件重新计算 SHA-256 并更新索引，确保审计链条闭环。

### Task 5: 编写仓储防回归检查 (`tests/test_repo_hygiene.py`)
- **检查项**：
  1. `git ls-files` 追踪的所有文本文件中不存在绝对用户路径（`/(?:Users|home)/`）、Windows 用户路径（`[A-Z]:\\Users\\`）及个人邮箱（`LICENSE` 声明除外）。
  2. 验证 `.gitattributes` 配置完整性。
  3. 验证 `.gitignore` 包含指定忽略模式。
  4. 验证所有 `evidence-hashes.json` 中记录的哈希值与磁盘文件一致。

### Task 6: 配置 `.gitattributes`
- **配置内容**：
  ```gitattributes
  * text=auto eol=lf
  *.json text eol=lf
  *.md text eol=lf
  *.py text eol=lf
  *.docx binary
  *.png binary
  *.pdf binary
  ```

### Task 7: 清理本地未跟踪大文件并更新 `.gitignore`
- **清理文件**：`.test_export_clean.pdf`、`tmp_test_full_qa.pdf`、`修改/`。
- **更新 `.gitignore`**：增加 `修改/`、`tmp_*.pdf`、`.test_*` 显式忽略规则。

### Task 8: 运行全量测试套件并固化 `p0-baseline.json`
- **离线全套测试**：329 passed, 6 skipped（包含 325 项历史测试 + 4 项新仓储检查）。
- **macOS Word 验收**：NW01–NW12 (12/12), SW01–SW07 (11/11), W04–W07 (4/4), M1 (3/3), Thesis (1/1), Pagination (1/1), Gates (2/2) 全部通过（合计 34/34 用例）。
- **写入产物**：`docs/acceptance/public-release/p0-baseline.json`。

---

## 3. 实施进度与变更记录

| 步骤 | 任务名称 | 执行状态 | 变更文件 / 产物 |
| :--- | :--- | :--- | :--- |
| Step 1 | 编制 P0 详细实施计划文档 | **已完成** | `docs/p0-implementation-plan.md` |
| Step 2 | 编写 `sanitize_paths.py` 及测试支持 | **已完成** | `docs/acceptance/tools/sanitize_paths.py` |
| Step 3 | 执行证据脱敏与哈希更新 | **已完成** | 53 处被跟踪证据文件及 3 份 `evidence-hashes.json` |
| Step 4 | 改造证据生成脚本接入 `relativize()` | **已完成** | `diagnose_pdfs.py`, `write_status.py`（两处） |
| Step 5 | 编写 `tests/test_repo_hygiene.py` | **已完成** | `tests/test_repo_hygiene.py` |
| Step 6 | 创建 `.gitattributes` 与更新 `.gitignore` | **已完成** | `.gitattributes`, `.gitignore` |
| Step 7 | 清理本地临时文件 | **已完成** | 移除根目录 `.test_export_clean.pdf`、`tmp_test_full_qa.pdf`、`修改/` |
| Step 8 | 运行回归验证并记录 `p0-baseline.json` | **已完成** | `docs/acceptance/public-release/p0-baseline.json` |

---

## 4. 关键证据与差异摘要

### 4.1 证据脱敏 Diff 示例（源文件 SHA-256 保持不变）
```diff
--- a/docs/acceptance/n0-n10-review/pageref-metadata.json
+++ b/docs/acceptance/n0-n10-review/pageref-metadata.json
@@ -35,3 +35,3 @@
   "source_hashes": {
-    "<developer-home>/.../docs/acceptance/word-followup/m2-fixtures/pageref-source.docx": "16c3645603c1c52ee0d4bca8a21a4250c76ddf3c137b479f2656bb281e33d50d"
+    "<repo>/docs/acceptance/word-followup/m2-fixtures/pageref-source.docx": "16c3645603c1c52ee0d4bca8a21a4250c76ddf3c137b479f2656bb281e33d50d"
   },
```

```diff
--- a/docs/acceptance/word-followup/academic-metadata.json
+++ b/docs/acceptance/word-followup/academic-metadata.json
@@ -18,3 +18,3 @@
   "source_hashes": {
-    "<developer-home>/.../examples/custom-format-demo/manuscript.docx": "46505dcd122c513102836eae0ef5666f77c5ae06bf34883c5b523bd6628e6e34"
+    "<repo>/examples/custom-format-demo/manuscript.docx": "46505dcd122c513102836eae0ef5666f77c5ae06bf34883c5b523bd6628e6e34"
   },
```

### 4.2 离线测试套件回归结果
```
Ran 329 tests in 38.264s
OK (skipped=6)
```
- 说明：包含原有 325 项业务与流水线测试，加上 4 项新增仓储卫生测试，全部通过；6 项有条件跳过均为离线环境下跳过的 Word 集成用例。

### 4.3 真实 Word 验收用例结果（macOS + Microsoft Word）
1. `tests/test_post_review_word.py` (NW01–NW12): 12/12 passed (69.966s)
2. `tests/test_remediation_word.py` (SW01–SW07): 11/11 passed (55.951s)
3. `tests/test_word_formatting.py` (W04–W07): 4/4 passed (25.053s)
4. `tests/test_m1_acceptance.py`: 3/3 passed (8.071s)
5. `tests/test_thesis_acceptance.py`: 1/1 passed (0.497s)
6. `tests/test_pagination.py`: 1/1 passed (1.385s)
7. `tests/test_publication_gates.py`: 2/2 passed (4.423s)
- **合计**：34 项真实 Word 集成用例 100% 通过，0 失败，0 错误。

---

## 5. 验收结论
P0 批次的所有执行动作已全面落地，6 项退出条件全部达到验收标准。工作树已处于干净、脱敏、跨平台配置对齐的健康状态，为后续 P1（跨平台加固与三平台 CI）奠定了基线。
