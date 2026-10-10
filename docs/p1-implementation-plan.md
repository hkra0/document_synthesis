# P1 实施计划：无 Office 部分的跨平台加固与 CI

## 1. 目标与背景

根据 [windows-and-public-release-plan.md](windows-and-public-release-plan.md) 的规划，在推进 Office 后端抽象（P2）与 Windows COM 原生对接（P3）之前，必须先将整个仓库中**不依赖 Microsoft Word / Office 的核心逻辑与基础设施**完成跨平台加固，并建立跨平台自动化持续集成（CI）。

### 核心收益
1. 确保在 Windows 控制台（CP936/GBK/UTF-8）及无终端重定向环境下，中文公文标题和输出信息不发生崩溃（R1）。
2. 在所有平台上统一阻止 Windows 非法字符、保留设备名与大小写重名，避免 macOS 导出的配置在 Windows 上无法创建文件（R3）。
3. 增加文件被其他进程（如 Word、防病毒软件）占用时的退避重试与清晰提示（R2）。
4. 大小写无关匹配文件扩展名（`.DOC`、`.Pptx`、`.Docx`）（R5）。
5. 缩短内部临时目录标识并增加长路径检查（R4）。
6. 清理测试可移植性障碍（符号链接条件跳过 `can_symlink()`、消除平台特定硬编码）（R8）。
7. 明确 Python 运行版本要求（`>=3.10`），配置 `pyproject.toml`，更新文档（G8）。
8. 建立 GitHub Actions 三平台（Linux、macOS、Windows）CI 矩阵。

---

## 2. 详细实施任务与清单

### 任务 1：控制台编码与子进程编码加固（R1）
- **目标文件**：`synthesize.py`、`smoke_test.py`、`lib/qa.py`、`lib/engine.py`、`lib/pagination.py`
- **改造内容**：
  - 在 `synthesize.py` 和 `smoke_test.py` 的入口处检测 `sys.stdout.encoding` 和 `sys.stderr.encoding`，非 UTF-8 时通过 `stream.reconfigure(encoding="utf-8", errors="replace")` 重新配置。
  - 生产代码及测试中所有 `subprocess.run` / `subprocess.Popen` / `subprocess.check_output` 调用如果使用 `text=True` 或 `capture_output=True`，统一明确声明 `encoding="utf-8", errors="replace"`，阻断平台默认内码解码异常。

### 任务 2：Windows 兼容文件名与项目名校验（R3）
- **目标文件**：`lib/config.py`、`tests/test_output_config.py`
- **改造内容**：
  - 在 `_validate_filename` 中增加：
    1. Windows 保留设备名称检测（不论大小写、不论是否有扩展名）：`CON`、`PRN`、`AUX`、`NUL`、`COM1-9`、`LPT1-9`。
    2. Windows 非法字符检测：`<`、`>`、`:`、`"`、`/`、`\`、`|`、`?`、`*` 以及 ASCII 控制字符（`ord(c) < 32`）。
    3. 尾部空格与点检测：主文件名（stem）或整体文件名不可包含尾随空格或尾随点。
  - 在 `_validate_project_name` 中增加相同的保留设备名、非法字符和尾部点/空格校验。
  - 在 `tests/test_output_config.py` 中补充成对的正反例单元测试。

### 任务 3：文件占用退避重试与友好错误提示（R2）
- **目标文件**：`lib/engine.py`、`lib/pagination.py`、`tests/test_delivery_pipeline.py`
- **改造内容**：
  - 提取可复用的原子文件替换/移动重试函数 `replace_with_retry(src, dst, max_retries=5, initial_delay=0.1)`。
  - 当捕获 `PermissionError`（Windows 下目标文件被 Word 独占打开时的典型异常）时进行有限次退避重试（总耗时约 2 秒）。
  - 若重试后依然失败，抛出包含明确业务指引的错误提示（“目标文件可能正在 Word 中打开: ...”），并保持原子事务回滚保障。
  - 编写模拟测试验证 `PermissionError` 重试成功与重试耗尽后的提示。

### 任务 4：统一大小写无关文件扩展名匹配（R5）
- **目标文件**：`lib/engine.py`、`lib/source_strategies.py`、`lib/config.py`、`synthesize.py`
- **改造内容**：
  - 替换直接在大小写敏感文件系统（如 Linux CI）上失效的 `glob("*.docx")`、`rglob("*.doc")`、`rglob("*.pptx")`。
  - 统一使用基于 `path.is_file() and path.suffix.lower() == target_suffix` 的扫描逻辑。
  - 补充对 `.DOC`、`.Pptx`、`.DOCX` 大写扩展名输入材料的测试。

### 任务 5：短工作目录与长路径预警（R4）
- **目标文件**：`lib/engine.py`、`synthesize.py`
- **改造内容**：
  - 将 `run_dir` 的 uuid 长度从 32 位缩减为 8 位：`run-{uuid.uuid4().hex[:8]}`，降低 Windows 路径超限风险。
  - 在 `--doctor` 中增加 Windows 专属的 `LongPathsEnabled` 注册表检查，并在当前工作路径/预计路径超长时输出预警。

### 任务 6：测试可移植性清理（R8）
- **目标文件**：`tests/test_post_review_contracts.py` 等相关测试
- **改造内容**：
  - 封装 `can_symlink()` 运行时探测函数。
  - 将符号链接特异性契约测试使用 `@unittest.skipUnless(can_symlink(), ...)` 明确标记跳过，在 Windows 无管理员/无开发者模式权限时规范计数。

### 任务 7：Python 版本界定与工程元数据（G8）
- **目标文件**：`README.md`、`requirements.txt`、`pyproject.toml`、`synthesize.py`
- **改造内容**：
  - 正式将最低支持 Python 版本界定为 3.10。
  - 创建 `pyproject.toml`，配置 `requires-python = ">=3.10"` 与项目基本构建元数据。
  - 更新 `README.md` 与 `synthesize.py` 中的最低版本声明。

### 任务 8：跨平台 GitHub Actions CI 矩阵配置
- **目标文件**：`.github/workflows/ci.yml`
- **改造内容**：
  - 配置 `ubuntu-latest`、`macos-latest`、`windows-latest`。
  - 矩阵运行 Python 3.10 与 3.12/3.13。
  - CI 执行步骤：
    1. 依赖安装；
    2. 全量单元测试 `python -m unittest discover -s tests -v`；
    3. 仓库卫生检查 `python -m unittest tests.test_repo_hygiene -v`；
    4. 结构检查与最小示例执行：`python smoke_test.py --structure-only`、`python synthesize.py --plan --project minimal-demo`。
    5. Windows 平台专门配置 `chcp 936` 与清空 `PYTHONIOENCODING` 测试控制台编码弹性。

---

## 3. 验收标准与退出条件
1. 本地所有测试（包括新加的 R1/R2/R3/R5/R8 测试）全量通过。
2. `test_repo_hygiene.py` 与 `sanitize_paths.py --check` 保持 100% 通过。
3. Windows 规范名与非法字符在 macOS 本地即可被严格拦截并报错。
4. CI 配置文件语法与步骤完整，已就绪供远端运行。

---

## 4. 实施完成情况与核验记录

### 4.1 任务执行清单

| 任务编号 | 任务名称 | 变更文件 | 关键改进与成果 | 状态 |
|---|---|---|---|---|
| **R1** | 控制台与子进程编码加固 | `synthesize.py`<br>`smoke_test.py`<br>`lib/qa.py`<br>`lib/engine.py`<br>`lib/pagination.py`<br>`tests/*` | 入口 `configure_console_encoding()` 重配非 UTF-8 终端；所有文本模式子进程显式声明 `encoding="utf-8", errors="replace"`。新增 `tests/test_console_encoding.py`（3 项测试）。 | **已完成** |
| **R3** | Windows 兼容文件名与项目名校验 | `lib/config.py`<br>`tests/test_output_config.py` | 拦截 Windows 保留设备名（`CON`、`PRN`、`AUX`、`NUL`、`COM1-9`、`LPT1-9`）、非法字符（`<>:"/\\|?*`）与尾部空格/点。`test_output_config.py` 正反例扩展通过（16 项）。 | **已完成** |
| **R2** | 文件占用退避重试与提示 | `lib/engine.py`<br>`lib/pagination.py`<br>`tests/test_platform_hardening.py`<br>`tests/test_build_metadata.py` | 实现 `_replace_with_retry` 与 `_replace_pdf_with_retry`，捕获 `PermissionError` 指数退避重试 5 次；耗尽抛出 `DeliveryLockError(PermissionError, BuildError)` 并包含“可能正在 Word 中打开”提示。 | **已完成** |
| **R5** | 大小写无关扩展名匹配 | `lib/engine.py`<br>`lib/source_strategies.py`<br>`lib/config.py`<br>`synthesize.py` | 废除直接 `glob("*.docx")`，采用 `p.is_file() and p.suffix.lower() == target_suffix`，适配 `.DOC`、`.Pptx`、`.DOCX` 等在 Linux/Windows 上的发现。 | **已完成** |
| **R4** | 短工作目录与长路径预警 | `lib/engine.py`<br>`synthesize.py`<br>`tests/test_platform_hardening.py` | 临时工作目录名缩短为 `run-<8位>`；`run_doctor` 支持 Windows `LongPathsEnabled` 注册表检查与 >240 字符路径风险预警。 | **已完成** |
| **R8** | 测试可移植性清理 | `tests/test_post_review_contracts.py` | 封装 `can_symlink()` 运行时探测，将符号链接特异性契约测试使用 `@unittest.skipUnless(can_symlink(), ...)` 明确隔离跳过。 | **已完成** |
| **G8** | Python 版本与工程元数据 | `README.md`<br>`pyproject.toml`<br>`.gitattributes`<br>`tests/test_repo_hygiene.py` | 最低 Python 版本设定为 3.10；创建 `pyproject.toml`（2026-10-10 复验时发现扁平目录自动发现导致构建失败，已显式声明 `packages`/`py-modules`；当前只支持仓库内 `pip install -e .`，可分发 wheel 留到 P8）；`.gitattributes` 增加 `*.toml` 并在仓库卫生测试中断言。 | **已完成** |
| **CI** | 跨平台 GitHub Actions CI 矩阵 | `.github/workflows/ci.yml` | 覆盖 Ubuntu、macOS、Windows 三平台与 Python 3.10、3.12 矩阵；针对 Windows 增加 `chcp 936` 无环境变量控制台弹性校验；新增 wheel 构建步骤。 | **已编写，待首次实跑**（工作流尚未推送，三平台全绿是 P1 退出条件） |

### 4.2 自动化测试与合规性验证结果

1. **离线与跨平台加固测试**：
   - 运行：`python3 -m unittest discover -s tests -p 'test_*.py'`
   - 结果：**338 passed, 6 skipped**（耗时约 39 秒），全部通过。
2. **仓库卫生与防泄露回归测试**：
   - 运行：`python3 -m unittest tests.test_repo_hygiene -v`
   - 结果：**4/4 passed**（路径/邮箱扫描、.gitattributes 配置完整、.gitignore 规则、证据哈希一致性）。
3. **脱敏工具自检**：
   - 运行：`python3 docs/acceptance/tools/sanitize_paths.py --check`
   - 结果：`Hygiene check passed: no user paths or emails found in tracked files.`。
4. **全局敏感路径扫描**：
   - 运行：`git grep -nE '/Users/|/home/'`
   - 结果：**0 matches**。

