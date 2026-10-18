# P6 实施计划：字体可移植性

> **文档状态**：已完成；F0–F6 全部完成（第 8–12、14–15 节）  
> **基线依赖**：P5 与封面模板修复（`main` = `0227f37`）  
> **核心目标**：按[执行计划](windows-and-public-release-plan.md)第 11 节，让构建在开始前就知道格式要求的字体本机是否具备；缺字体时明确报告，并允许在格式包中声明、经显式启用后使用的回退字体。默认不替换任何字体，不悄悄改变公文格式。  
> **环境约束**：开发用的 macOS 与 Windows 都已安装公文字体（`仿宋_GB2312`、`楷体_GB2312`、`方正小标宋简体` 等）。“缺字体”场景不靠卸载字体获得，验证方法见第 5 节。

---

## 1. 退出条件

沿用执行计划，并按本机已装字体的实际情况细化：

1. `--plan` 与 `--doctor` 列出格式需要的每个字体及其状态（已安装 / 缺失 / 无法判断），缺失时提示 Word 会替换字体、分页可能变化，`exact` 结果只对当前字体环境有效。
2. 已安装的公文字体在 macOS 与 Windows 上都被识别为“已安装”，没有误报缺失（以 F0 探针记录的实际情况为准）。
3. 格式要求一个本机确实没有的字体时，`--doctor` 报告缺失；不启用回退时构建照常进行，但终端和 `build-diagnostics.json` 记录缺失字体。
4. 启用回退后，Windows + Word 的精确构建与精确 smoke 通过，生成的文档中实际写入回退字体，`build-diagnostics.json` 与构建元数据记录 `请求字体 → 实际字体` 的映射。
5. 回退字体本身也缺失时，构建失败并说明原因，不继续往下尝试未声明的字体。
6. 离线全量、macOS 真实 Word 回归、Windows 真实 Word 回归与 WW01–WW12 均无回退（WW 矩阵的页码与 P4 入库结果一致）。

不在本批范围：来源文档自身带来的字体（只在 `--plan` 中列出，不替换）；在仓库中分发任何字体文件；推算字体替换后的分页。

## 2. 关键问题：怎样判断“字体已安装”

`w:eastAsia="仿宋_GB2312"` 写的是字体的本地化家族名。各平台能直接拿到的名字不一定是这个：

| 平台 | 现成的来源 | 问题 |
| --- | --- | --- |
| Windows | 注册表 `HKLM`/`HKCU\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts` | 值名是安装程序随意写的显示名，可能是英文（如 `FangSong_GB2312 (TrueType)`）、可能是中文、TTC 会把多个家族写在一个值名里；按用户安装的字体只在 `HKCU`，文件在 `%LOCALAPPDATA%\Microsoft\Windows\Fonts` |
| macOS | 无 `fc-list`（Homebrew 才有） | Word 自带字体在 `Microsoft Word.app/Contents/Resources/DFonts`，系统字体工具看不到，但 Word 能用 |
| Linux | `fc-list` | 未必安装；无 Word，只影响 `--plan`/`--doctor` 的提示 |

因此计划是：各平台只负责列出字体文件，**统一读取字体文件 `name` 表**中所有语言的家族名（nameID 1、4、16）来匹配，不依赖注册表值名。解析器只用标准库（`struct`），支持 TTF、OTF、TTC，不新增依赖。

另一个候选来源是 Word 自己的字体列表（COM `Application.FontNames`），它最接近“Word 排版时能不能用”。但它需要启动 Word，只能用于 `--doctor` 的探针，不能用于不启动 Office 的 `--plan`。两者是否一致由 F0 探针确认。

## 3. 设计

### 3.1 `lib/fonts.py`

| 接口 | 作用 |
| --- | --- |
| `read_font_names(path)` | 解析一个字体文件（TTC 返回每个子字体），给出各语言的家族名、全名、PostScript 名；损坏或不支持的文件返回错误而不抛出 |
| `installed_fonts()` | 按平台列出字体文件并汇总为“名字 → 文件”的索引；Windows 读 `HKLM` 与 `HKCU` 注册表并补扫两个字体目录；macOS 扫系统、用户与 Word `DFonts` 目录；Linux 优先 `fc-list`，没有时扫常见目录。结果在进程内缓存 |
| `required_fonts(config, format)` | 收集一次构建会写入文档的字体：格式包各样式的 `east_asia`/`latin`、旧配置 `fonts`、渲染器中写死的字体（目录标题 `黑体`、表格 `宋体` 等） |
| `check_fonts(required, installed)` | 返回每个字体的状态 `installed` / `missing` / `unknown`（字体索引建立失败时为 `unknown`，不能写成已安装） |

名字比较时去掉首尾空白并忽略大小写，不做其他模糊匹配：`仿宋` 与 `仿宋_GB2312` 是不同的字体。

按 F0 结果补充（第 8 节）：

- `.fon` 位图字体是可执行文件格式，没有 `name` 表，记为“不支持”而不是解析失败；它们不会出现在格式要求里。
- 系统 `FontSubstitutes` 中有 `FangSong_GB2312 → FangSong`、`KaiTi_GB2312 → KaiTi`。缺少公文字体时，`--doctor` 可以提示“系统登记了替代字体”，但 Word 排版时是否采用这条替代未经验证，只能写成提示，不能把字体状态改为已安装。替代项的键是英文名，要和格式中的中文名对应，需要一张很小的已知别名表（`仿宋_GB2312`↔`FangSong_GB2312`、`楷体_GB2312`↔`KaiTi_GB2312`），只用于这条提示。

### 3.2 `--plan` 与 `--doctor`

- `--plan`：在构建计划后增加“字体”一节，逐个列出状态和来源（格式包样式名 / 渲染器）。只读，不启动 Office。
- `--doctor`：增加同样的检查；缺失字体计为警告，不改变现有退出码含义（0/2/1）。Windows + Word 时额外用 `FontNames` 交叉核对，两者不一致时如实列出，不自行裁决。
- 按 F0 结果：Word 对每个字体只列一个名字（中文界面下中文字体列中文名，不列英文别名），所以交叉核对用 `name` 表各语言的名字去对 Word 列表，不能反过来。本机没有文件、但 Word 列表中有的字体记为 `word_only`，提示“Word 可用，本机未安装，换一台电脑结果可能不同”，不计为缺失。

### 3.3 回退字体

格式包增加可选字段：

```json
"font_fallbacks": {
  "仿宋_GB2312": ["仿宋"],
  "楷体_GB2312": ["楷体"],
  "方正小标宋简体": ["小标宋", "宋体"]
}
```

- 只在命令行加 `--font-fallbacks`（环境变量 `DOCUMENT_SYNTHESIS_FONT_FALLBACKS=1`）时生效；只替换本机确实缺失的字体，按列表顺序取第一个已安装的。
- 替换发生在写入样式时（`lib/style_applier.py` 与 `lib/legacy_format.py` 取字体名的地方统一经过一个映射函数），不改格式包文件。
- 回退列表中没有已安装的字体时，构建失败。
- `format_schema` 校验字段类型；回退目标不能与原字体相同。内置 `legacy-official` 预设附带上面的默认列表，但同样需要 `--font-fallbacks` 才生效。

### 3.4 记录

`build-diagnostics.json` 与 v3 构建元数据新增 `fonts`：

```json
"fonts": {
  "fallbacks_enabled": true,
  "checked": [{"name": "仿宋_GB2312", "status": "installed"}],
  "substitutions": [{"requested": "P6缺失测试字体", "used": "仿宋", "reason": "missing"}],
  "missing_unreplaced": []
}
```

终端在构建结束前打印缺失与替换摘要。精确构建的结论仍是“页码已核验”，但只对记录中的字体环境有效。

## 4. 步骤与执行位置

| 步骤 | 内容 | 在哪里做 | 退出条件 |
| --- | --- | --- | --- |
| F0 | 字体探针：记录 Windows 上字体的实际注册方式、`name` 表中的名字、Word `FontNames` | 脚本在 macOS 写好；**在 Windows 上运行** | 得到探针结果，回答第 6 节的问题 |
| F1 | `lib/fonts.py` 与单元测试（用合成的最小 TTF/TTC 字节测解析，用注入的文件列表测索引） | macOS | 离线全量通过；macOS 本机公文字体全部识别为已安装 |
| F2 | `--plan`/`--doctor` 字体检查 | macOS | 离线测试覆盖已安装、缺失、索引失败三种状态 |
| F3 | `font_fallbacks` 字段、`--font-fallbacks`、写样式时的映射、记录 | macOS | 合成夹具：缺失字体被替换、回退也缺失时失败、未启用时不替换 |
| F4 | macOS 真实 Word 验证与回归 | macOS | 第 1 节第 3、4、6 条在 macOS 上满足 |
| F5 | Windows 真实 Word 验证与回归 | **Windows** | 第 1 节第 2–6 条在 Windows 上满足，证据入库 |
| F6 | README 说明公文字体的合法获取途径与 `--font-fallbacks` 用法；更新执行计划状态 | macOS | 文档只写已验证的行为 |

F1–F3 可以在 F0 结果回来前开始，但注册表处理与 `FontNames` 交叉核对的细节等 F0 结果再定。

## 5. 已装公文字体时怎样验证“缺字体”

不卸载本机字体。各项检查分别用下面的方法：

1. **识别已安装字体（不误报）**：直接在两台已装公文字体的机器上运行。这一项本来就需要真实安装的字体，本机环境正好适合。
2. **缺失与回退**：测试专用格式包把正文字体设为一个保证不存在的名字 `P6缺失测试字体`，回退到 `仿宋`。这个字体在 Word 里也真的不存在，所以 Windows 上的精确构建会经历真实的“缺字体”路径：不启用回退时 Word 自行替换、构建照常完成并记录缺失；启用回退时文档写入 `仿宋`。
3. **回退也缺失**：回退列表只写另一个不存在的名字，预期构建失败。
4. **另一个 Windows 用户账户（可选，能覆盖真实缺字体）**：F0 显示本机公文字体是按用户安装的（只在 `HKCU`）。同一台电脑上另建一个本地账户，该账户看不到这些字体，但系统自带的 `仿宋`/`楷体` 与 Word 都在。如果 Word 在该账户中能登录激活，就能在“真正缺少 `仿宋_GB2312`、有 Word”的环境里跑 `--doctor`、精确构建和回退构建，下面“本批不能覆盖的部分”也就能验证。
5. **干净系统上的识别（可选）**：Windows 11 Pro 自带 Windows Sandbox，里面是全新系统、没有公文字体、也没有 Word。可以在其中运行 `--plan`/`--doctor` 与探针，确认 `仿宋_GB2312` 等报告为缺失、系统自带的 `仿宋`/`楷体`/`黑体`/`宋体` 报告为已安装。需要在“启用或关闭 Windows 功能”中开启，属于可选项。

不做第 4 项时本批不能覆盖的部分：真正缺少 `仿宋_GB2312` 时 Word 的替换结果与分页变化。第 2 项中 Word 替换的是一个虚构字体名，没有对应的 `FontSubstitutes` 项，替换规则可能与缺少真实公文字体时不同。这一点在状态记录中如实写明，不写成已验证。

## 6. F0 探针要回答的问题

1. 公文字体在注册表中登记在 `HKLM` 还是 `HKCU`，值名是中文、英文还是两者都有？
2. 每个公文字体文件 `name` 表中的家族名有哪些语言版本，是否包含格式包使用的确切名字（如 `仿宋_GB2312`）？
3. 字体目录中是否有未在注册表登记的文件？
4. `FontSubstitutes` 中是否有与公文字体相关的替换项？
5. Word `FontNames` 中出现的是哪种名字？与 `name` 表结果是否一致？
6. 哨兵名 `P6缺失测试字体` 是否在所有来源中都确实不存在？

## 7. F0 探针用法

脚本：[`docs/acceptance/public-release/font_probe.py`](acceptance/public-release/font_probe.py)。只读：不修改注册表与字体，`--word` 时启动一个隔离的 Word 实例读取字体列表后退出。

```
python docs/acceptance/public-release/font_probe.py --output font-probe-windows.json
python docs/acceptance/public-release/font_probe.py --word --output font-probe-windows.json
```

终端打印每个目标字体在各来源中的结果；完整记录写入 `--output`。输出中的用户目录与 Windows 目录已替换为 `<user-home>`、`<windir>`，但仍含本机字体清单，先放在仓库之外，整理后再决定入库哪些部分。

## 8. F0 结果：Windows 第一次运行（2026-10-18）

环境：Windows 11 Pro（10.0.26300），Python 3.14.4，探针提交 `4984469`。原始 JSON 含本机完整字体清单，未入库。

| 问题（第 6 节） | 结果 |
| --- | --- |
| 1. 注册表位置与值名 | 三个公文字体都按用户安装，只在 `HKCU`（共 5 项），文件在 `<user-home>\AppData\Local\Microsoft\Windows\Fonts`。值名全是英文：`FangSong_GB2312 (TrueType)`、`KaiTi_GB2312 (TrueType)`、`FZXiaoBiaoSong-B05S (TrueType)`。按格式中的中文名查注册表值名，三个都查不到 |
| 2. `name` 表 | 三个文件的 zh-CN 家族名分别是 `仿宋_GB2312`、`楷体_GB2312`、`方正小标宋简体`，与格式中的名字完全一致。系统自带的 `仿宋`/`楷体`/`黑体`/`宋体`/`微软雅黑` 同样以 zh-CN 家族名匹配到 |
| 3. 未登记的文件 | `%WINDIR%\Fonts` 中 3 个（`marlett.ttf`、`mstmc.ttf`、`MTEXTRA.TTF`），都是系统内部字体，与公文无关；用户字体目录没有 |
| 4. `FontSubstitutes` | `HKLM` 中有 `FangSong_GB2312 → FangSong`、`KaiTi_GB2312 → KaiTi`；`HKCU` 没有该键 |
| 5. Word `FontNames` | **未读取**：运行探针的 Python 没有安装 pywin32 |
| 6. 哨兵名 | `P6缺失测试字体` 在 `name` 表和注册表中都不存在；Word 中是否存在待第 5 项确认 |

其他发现：

- 471 个文件中 12 个解析失败：11 个是 `.fon` 位图字体（文件头 `MZ`），1 个是未登记的 `mstmc.ttf`（不是 sfnt 格式）。探针已改为把 `.fon` 记为“不支持”。
- `楷体_GB2312` 的文件名在磁盘上是乱码（`┐¼╠σ_GB2312.ttf`，GBK 字节被当作 CP437 解出），注册表也按这个文件名登记。文件名和注册表值名都不能用来识别字体，进一步说明需要读 `name` 表。
- 结论：第 2 节“统一读取 `name` 表”的设计成立；`HKCU` 与用户字体目录必须纳入，否则会把本机三个公文字体全部误报为缺失。

### 第二次运行（带 `--word`，探针提交 `240b8e7`）

环境同上，Word 16.0（build 16.0.20430）。471 个文件中只剩 `mstmc.ttf` 1 个解析失败，`.fon` 已记为不支持。

| 问题（第 6 节） | 结果 |
| --- | --- |
| 5. Word `FontNames` | 共 1490 个名字，其中 94 个是竖排变体（`@` 开头）。Word 对每个字体只列一个名字：中文字体列 zh-CN 家族名（`仿宋_GB2312`、`楷体_GB2312`、`方正小标宋简体`、`仿宋`、`宋体`、`微软雅黑` 都在），对应的英文名（`FangSong_GB2312`、`SimSun`、`Microsoft YaHei` 等）都不在；西文字体列英文名 |
| 5. 与 `name` 表是否一致 | 本机每个字体文件的子字体，都至少有一个家族名出现在 Word 列表中，唯一例外是可变字体 `Segoe UI Variable`。反过来，Word 列表中有 1134 个名字在本机字体文件里找不到（如 `Aptos`、`Abadi`、`Amasis MT Pro`），其中没有中日韩字体。按名字推断是 Microsoft 365 的云字体（使用时才下载），未单独核实 |
| 6. 哨兵名 | `P6缺失测试字体` 在 `name` 表、注册表和 Word 列表中都不存在 |

结论：

- 对格式包当前用到的字体（公文字体、系统中文字体、`Times New Roman`、`Segoe UI`、`Cambria`），读 `name` 表的结果与 Word 一致，`--plan` 不启动 Word 也能给出正确结论。
- 差异只出现在本机没有文件的字体上：Word 可能通过云字体使用它们。第 3.2 节为此增加 `word_only` 状态。
- 第 5 节第 2、3 项使用的哨兵名已确认在 Word 中也不存在，可以用于缺失与回退测试。
- F0 完成，可以开始 F1。

## 9. F1 结果：`lib/fonts.py`（2026-10-18）

| 接口 | 实现 |
| --- | --- |
| `read_font_file(path, source)` | 由探针的解析器移入。返回 `FontFile`：`status` 为 `ok`/`unsupported`（`.fon`）/`error`，不抛出；TTC 中个别子字体损坏时保留其余子字体并在 `detail` 中说明 |
| `windows_font_files` / `directory_font_files` / `platform_font_files` | Windows：`HKLM` 与 `HKCU` 注册表（相对文件名按 `%WINDIR%\Fonts` 解析，登记了但文件不存在的项跳过），加两个字体目录中未登记的文件；macOS：系统、用户与 Word `DFonts` 目录；Linux：`fc-list`，没有时扫常见目录 |
| `FontIndex.build` / `installed_fonts()` | 按 `family`、`typographic_family`、`full` 三种名字的所有语言建立索引，PostScript 名不参与。一个字体文件都没找到、或列出字体文件时出错，索引记为不完整；个别文件解析失败不影响完整性。`installed_fonts()` 进程内缓存 |
| `required_fonts(config)` | 解析后格式各样式的 `east_asia`/`latin`/`complex_script`，旧配置 `fonts`，以及代码直接写入的 `宋体`（页脚页码，`styles.setup_footer`）。按名字去重，保留全部来源 |
| `check_fonts(required, index, word_font_names=None)` | `installed` / `missing` / `unknown`（索引不完整时查不到）/ `word_only`（只有提供 Word 字体列表时才会出现） |

核查代码直接写入的字体时发现：目录条目字体取自 `fonts` 配置，`styles.build_toc_sdt` 中写死的 `黑体` 当前没有调用方，因此未列入。

`tests/test_office_backends.py` 禁止 `lib/` 中出现 `darwin` 判断（P2 的 Office 隔离约束）。字体位置本来就因平台而异，与 Office 调用无关，因此把 `lib/fonts.py` 加入允许名单，并注明原因。

验证（macOS，Python 3.14.0）：

- `tests/test_fonts.py` 22 项：用合成的最小 TTF/OTF/TTC 字节测解析，用临时目录与注入的注册表项测 Windows 来源（含 F0 中的乱码文件名）。
- 本机：665 个字体文件，全部解析成功，建索引 0.11 秒。`仿宋_GB2312`、`楷体_GB2312`、`方正小标宋简体`、`仿宋`、`楷体`、`黑体`、`宋体`、`微软雅黑`、`Times New Roman`、`Cambria` 都识别为已安装；`Segoe UI` 与哨兵名为缺失（本机确实没有 `Segoe UI`）。默认旧配置需要的 6 个字体都识别为已安装。
- 离线全量 478 项通过，19 项跳过（与 P5 相同的 Windows 专属与需手动开启的真实 Word 用例）。
- 未在 Windows 上运行 `lib/fonts.py`；Windows 来源的逻辑与 F0 探针相同，F5 时实测。

## 10. F2 结果：`--plan` 与 `--doctor` 字体检查（2026-10-18）

| 位置 | 改动 |
| --- | --- |
| `lib/fonts.py` | `font_report(config, index=None, word_font_names=None)` 返回可序列化的检查结果；`report_lines()` 转成终端行：`[已安装]`/`[缺失]`/`[无法判断]`/`[仅 Word 可用]`，来源超过 3 处时缩写为“等 N 处”。有缺失或无法判断的字体时追加一行提示：Word 会替换字体、分页可能不同、精确页码只对当前字体环境有效 |
| `lib/engine.py` | `UnifiedSynthesizer.plan()` 的结果增加 `fonts`；新增可选参数 `word_font_names`，只由 `--doctor` 传入 |
| `synthesize.py` | `--plan` 在“交付文档”之后、“目录预览”之前输出“字体”一节。`--doctor` 在检查项之后输出“字体”一节：指定 `--source`/`--project` 时检查该项目，否则检查默认配置。缺失字体不改变退出码 |
| `lib/office/win_com_worker.py`、`lib/office/win_com.py` | worker 新增单次操作 `font_names`（读取 `Application.FontNames`，经忙碌重试）；`WinComBackend.font_names()` 启动独立 Word 实例读取后退出，失败时返回 `None`。没有加入 `probe`，构建时调用的 `probe` 不受影响 |

`--doctor` 只在后端能精确分页且提供 `font_names()` 时读取 Word 字体列表，目前只有 Windows Word 后端提供；读取失败时输出“未能读取 Word 字体列表，字体只按本机字体文件判断”。macOS Word 未实现读取字体列表，只按本机字体文件判断。

未实现：第 3.1 节中“系统登记了替代字体”（`FontSubstitutes`）的提示，留到 F3 与回退字体一起处理。

验证（macOS，Python 3.14.0，Word 16.113.4）：

- 中文示例 `--plan`：示例配置使用英文字体名（`FangSong_GB2312`、`SimHei`、`KaiTi_GB2312`、`FZXiaoBiaoSong-B05S`），通过 `name` 表中的英文名全部识别为已安装。用覆写把正文字体改为 `P6缺失测试字体` 后，该字体报告为缺失并出现替换提示。
- `--doctor`（无参数，真实 Word）：默认配置的 6 个字体都已安装，退出码 0。`--doctor` 检查 `custom-format-demo`：5 个字体都已安装。
- 新增测试：`tests/test_font_reporting.py` 11 项（用合成字体代替本机字体，结果与运行环境无关）；`tests/test_win_com_backend.py` 新增 3 项（worker `font_names`、后端成功/失败/Word 不可用）。
- 离线全量 492 项通过，19 项跳过；`DOCUMENT_SYNTHESIS_OFFICE_BACKEND=none` 下字体相关 33 项通过。
- 未在 Windows 上运行：`font_names` 的真实 Word 读取与 `--doctor` 输出留到 F5；F0 探针已在同一台 Windows 上用同样的 COM 调用读到 1490 个名字。

## 11. F3 结果：回退字体与记录（2026-10-18）

### 实现

| 位置 | 改动 |
| --- | --- |
| `schemas/format-v1.schema.json`、`lib/format_schema.py`、`lib/format_resolver.py` | 格式包新增可选字段 `font_fallbacks`（`{字体: [候选, ...]}`），随继承链与 `format.overrides` 合并（对象逐键合并，候选数组整体替换），解析为 `ResolvedFormat.font_fallbacks` |
| `lib/fonts.py` | `parse_font_fallbacks` 校验类型、空列表、回退与原字体相同；`resolve_font_fallbacks` 只替换缺失且声明了回退的字体，按顺序取第一个已安装的候选；`substitute_format_fonts` 返回替换后的新 `ResolvedFormat`；`font_build_record` / `build_summary_lines` 生成记录与终端摘要 |
| `lib/config.py` | `ProjectConfig` 构造的最后一步：v3 取格式包的 `font_fallbacks`；v1/v2 接受顶层 `font_fallbacks`，未声明时使用内置公文字体列表（中英文名都有，因为示例配置用英文名）。启用时同时替换格式样式与 `fonts`，结果存入 `font_resolution` |
| `lib/build_plan.py` | 启用回退时把 `font_resolution` 计入配置契约快照；未启用时不加字段，既有配置的契约哈希不变 |
| `lib/engine.py` | 构建开始时打印 `[字体] 回退: A → B` 或 `[字体] 本机缺少且未替换: ...`；v3 构建元数据新增 `fonts` |
| `synthesize.py`、`smoke_test.py` | 新增 `--font-fallbacks`（环境变量 `DOCUMENT_SYNTHESIS_FONT_FALLBACKS=1`）；`--plan` 的字体一节在启用时先列出 `[回退]` 行 |

### 与第 3 节设计的差异

1. **替换发生在配置加载时**，而不是写样式时。这样计划、渲染、交付格式核验和 smoke 用的是同一份替换后的配置，不会把回退字体当作格式偏差；替换结果计入配置契约哈希，所以用 `--font-fallbacks` 构建的 v3 交付物，smoke 不加 `--font-fallbacks` 时会因契约哈希不一致而失败（已实测）。
2. **记录位置**：`build-diagnostics.json` 只在 v3 构建失败时写出，v1/v2 构建不写任何元数据文件。因此字体记录写在 v3 的 `build-metadata.json`（草稿为 `build-metadata_draft.json`）的 `fonts` 中，所有构建都在终端打印摘要；v1/v2 构建只有终端记录。回退字体全部缺失时在加载配置阶段失败，只有终端错误，不生成诊断文件。第 1 节第 4 条据此理解。
3. **内置预设未改动**：没有给 `legacy-official@1.0.0` 加默认列表，避免修改已发布的预设版本；v1/v2 配置使用内置列表，v3 需要在格式包或 `format.overrides` 中声明。
4. **代码直接写入的字体**（页脚页码的 `宋体`）不经过样式，不参与替换，缺失时列入“缺少且未替换”。
5. 本机字体列表不完整时，启用回退直接失败，不猜测哪些字体缺失。
6. 第 3.1 节“系统登记了替代字体”（`FontSubstitutes`）的提示仍未实现；是否需要，等 F5 在 Windows 上看过真实缺字体时 Word 的表现再定。

### 验证（macOS，Python 3.14.0，Word 16.113.4）

| 场景 | 结果 |
| --- | --- |
| 中文示例（v2），覆写正文字体为哨兵名、回退 `["P6另一个缺失字体", "仿宋"]`，`--font-fallbacks` 精确构建 | 终端 `[字体] 回退: P6缺失测试字体 → 仿宋`；构建与精确 smoke（加 `--font-fallbacks`）通过，5 页；交付物 `styles.xml` 正文样式 `w:eastAsia="仿宋"`，全文没有哨兵名 |
| 同上，`--plan` 不加 `--font-fallbacks` | 哨兵名报告为缺失 |
| 回退候选全部缺失 | `--plan` 报错“声明的回退字体也都没有安装”，退出码 1 |
| `custom-format-demo`（v3），`format.overrides` 把正文字体改为哨兵名并回退到 `宋体`，`--font-fallbacks` 精确构建 | 构建与 smoke 通过；`build-metadata.json` 的 `fonts` 记录 6 个字体的检查结果和替换 `P6缺失测试字体 → 宋体`；smoke 不加 `--font-fallbacks` 时报“配置契约 SHA-256 与当前 PreparedBuild 不一致” |
| 同上，不加 `--font-fallbacks` | 构建通过（Word 自行替换字体）；终端提示缺失并建议 `--font-fallbacks`；元数据 `fallbacks_enabled: false`、`missing_unreplaced: ["P6缺失测试字体"]` |

- `tests/test_font_fallbacks.py` 20 项。
- 离线全量 512 项通过，19 项跳过；设置 `DOCUMENT_SYNTHESIS_FONT_FALLBACKS=1` 后全量结果相同（本机字体齐全，不发生替换）。
- 含真实 Word 用例（`DOCUMENT_SYNTHESIS_WORD_TEST=1`）的全量：539 项通过，13 项跳过（Windows 专属用例）。
- v2 的 smoke 不检查正文字体，不加 `--font-fallbacks` 也能通过；这是 v2 交付格式核验原有的范围，本批未扩大。

## 12. F4 结果：macOS 真实 Word 验证与回归（2026-10-18）

环境：macOS 27.0（arm64），Python 3.14.0，Word 16.113.4，代码提交 `d19332e`。本机已安装公文字体，因此“缺字体”一律用哨兵名 `P6缺失测试字体` 构造（第 5 节第 2、3 项）。

| 第 1 节退出条件 | 证据 | 结果 |
| --- | --- | --- |
| 3. 缺字体时 `--doctor` 报告、构建照常并记录 | 中文示例覆写正文字体为哨兵名后运行 `--doctor`（真实 Word 探针） | 退出码 0，“字体”一节列出 `[缺失] P6缺失测试字体` 及来源与替换提示；加 `--font-fallbacks` 时列出 `[回退] P6缺失测试字体 → 仿宋`。不启用回退的构建结果见第 11 节（v3 元数据 `missing_unreplaced`） |
| 4. 启用回退后精确构建与 smoke 通过，文档写入回退字体并记录 | 第 11 节验证表（v2 中文示例、v3 `custom-format-demo`） | 通过；代码自那次验证后未改动 |
| 6. 回归 | 见下 | 无回退 |

回归：

- **WW01–WW12**：`word_matrix.py run --label macos-word --output ww-results-p6.json` 12/12 `verified`；同样的矩阵在 `DOCUMENT_SYNTHESIS_FONT_FALLBACKS=1` 下运行（`ww-results-p6-fallbacks.json`）也是 12/12。两次结果分别用 `word_matrix.compare_results` 与 P4 入库的 `macos-word/ww-results.json`（`8f53eca`）对照：12 个用例全部可比，书签物理页、打印页码、导出页数与页码标签**差异为 0**。两份结果已入库。
- **示例**：`minimal-demo`、中文示例、`custom-format-demo` 各自在不加与加 `--font-fallbacks` 两种方式下精确构建并运行精确 smoke（对应地加同样的参数），6 次全部通过；本机字体齐全，终端没有字体提示。两种方式生成的交付物，除 `docProps/core.xml` 与 rsid 属性外，DOCX 各部件逐字节相同；v3 元数据的 `page_map` 相同，`config_contract_sha256` 不同（启用回退时 `font_resolution` 计入契约，符合第 11 节设计）。
- **测试**：离线全量 512 项通过、19 项跳过；含真实 Word 用例 539 项通过、13 项跳过（第 11 节）。

F4 完成。剩余：F5 在 Windows 上验证第 1 节第 2–6 条（含 `font_names` 读取 Word 字体列表、`--doctor` 交叉核对、WW 矩阵回归），以及可选的“另一个 Windows 账户”真实缺字体验证（第 5 节第 4 项）；F6 文档。

## 13. F5 执行说明：Windows 验证

脚本：[`docs/acceptance/public-release/p6_font_check.py`](acceptance/public-release/p6_font_check.py)。按第 1 节退出条件串行运行生产 CLI；“缺字体”用 [`p6-fonts/`](acceptance/public-release/p6-fonts/) 中的覆写文件构造（哨兵名 `P6缺失测试字体`）。结果写入 `docs/acceptance/public-release/<label>/p6-font-check.json`，路径已脱敏；构建产物放在临时目录，结束后删除。

| 步骤 | 内容 | 对应条件 |
| --- | --- | --- |
| F01 | 离线全量 | 6 |
| F02 | `--doctor`（默认配置）：6 个默认字体都已安装；Windows 上读取到 Word 字体列表 | 2 |
| F03 | `--plan` 中文示例：示例中的英文字体名全部识别为已安装 | 2 |
| F04 | `--doctor` 中文示例 + 哨兵字体：报告缺失，退出码 0 | 3 |
| F05 | 中文示例 + 哨兵字体，不启用回退：精确构建与 smoke 通过，终端记录缺失，正文样式保留请求的字体名 | 3 |
| F06 | 中文示例 + 哨兵字体，`--font-fallbacks`（候选 `P6另一个缺失字体`、`仿宋`）：构建与 smoke 通过，正文样式为 `仿宋`，文档中没有哨兵名 | 4 |
| F07 | 回退候选全部缺失：`--plan --font-fallbacks` 退出码 1 | 5 |
| F08 | `custom-format-demo`（v3）+ 哨兵字体回退到 `宋体`：构建与 smoke 通过，元数据记录替换；smoke 不加 `--font-fallbacks` 时因配置契约不一致失败 | 4 |
| F09 | 三个示例在不加与加 `--font-fallbacks` 时都通过，DOCX 部件相同（不计 `core.xml` 与 rsid） | 6 |
| F10 | WW01–WW12 两次（不加与加 `DOCUMENT_SYNTHESIS_FONT_FALLBACKS=1`），各自与入库的 `<label>/ww-results.json` 页码对照，差异须为 0；结果写入 `ww-results-p6.json`、`ww-results-p6-fallbacks.json` | 6 |
| F11 | 含真实 Word 用例的全量（`DOCUMENT_SYNTHESIS_WORD_TEST=1`） | 6 |

macOS 试运行（`--skip-tests --skip-matrix --skip-word-tests`，`p6-font-check.json` 未入库，F01/F10/F11 已在 F4 单独完成）：F02–F09 全部 `passed`，约 70 秒。

Windows 上的运行方式（项目环境的 Python，需已安装 pywin32）：

```
git fetch origin
git checkout p6-font-portability
git pull
python -m pip install -r requirements.txt
python docs/acceptance/public-release/p6_font_check.py
```

只想先跑快的部分时加 `--skip-matrix --skip-word-tests`；只重跑个别步骤用 `--only F06,F08`。完成后把 `docs/acceptance/public-release/windows-word/` 下新生成的 `p6-font-check.json`、`ww-results-p6.json`、`ww-results-p6-fallbacks.json` 提交到本分支，或把文件发回整理。

可选（第 5 节第 4 项）：在另一个本地账户中重复 F02–F06（公文字体只装在当前账户），观察真正缺少 `仿宋_GB2312` 时 `--doctor` 的结果，以及默认配置 `--font-fallbacks` 构建的回退与页码。这一步需要该账户中 Word 已登录激活，脚本没有自动化。

## 14. F5 结果：Windows 真实 Word 验证与回归（2026-10-18）

环境：Windows 11 Pro（10.0.26300），Python 3.14.4，Microsoft Word 16.0（build 16.0.20430，64 位），后端 `win-word` / `exact`。

### 第一次完整运行（提交 `95b476c`，[`p6-font-check.json`](acceptance/public-release/windows-word/p6-font-check.json)）

F02–F10 全部 `passed`，F01 与 F11 `failed`：

- **F02**：默认配置的 6 个字体都已安装；Word 字体列表读取到 1490 个。
- **F03**：中文示例中的英文字体名（`FangSong_GB2312`、`KaiTi_GB2312`、`FZXiaoBiaoSong-B05S`、`SimHei` 等）都识别为已安装。
- **F04–F08**：哨兵字体报告为缺失、退出码不变；不启用回退时构建与 smoke 通过、正文保留请求的字体名；启用回退时正文写入 `仿宋`、文档中没有哨兵名；回退候选全部缺失时退出码 1；v3 元数据记录 `P6缺失测试字体 → 宋体`，smoke 不加 `--font-fallbacks` 时因配置契约不一致失败。
- **F09**：三个示例在不加与加 `--font-fallbacks` 时都通过，DOCX 部件相同（不计 `core.xml` 与 rsid）。
- **F10**：WW01–WW12 两次都是 12/12 `verified`，与 P4 入库的 `windows-word/ww-results.json` 对照，12 个用例全部可比、差异为 0。结果入库：[`ww-results-p6.json`](acceptance/public-release/windows-word/ww-results-p6.json)、[`ww-results-p6-fallbacks.json`](acceptance/public-release/windows-word/ww-results-p6-fallbacks.json)。
- **F01/F11 失败原因**：两次都是同样的两个测试，与本机环境有关，不是 P6 功能问题。
  1. `test_fonts.WindowsSourcesTest.test_without_user_fonts_dir`：传入 `local_appdata=None` 时函数按设计回落到 `LOCALAPPDATA`，本机用户字体目录中的 5 个公文字体文件混入结果。CI 的 Windows 机器没有用户字体，所以没有暴露。测试改为临时移除 `LOCALAPPDATA`。
  2. `test_docx_inspector` 导入失败：`tests/` 没有 `__init__.py`，是命名空间包；本机 site-packages 中有第三方包（`ultralytics`）安装的常规包 `tests`，导入时优先于命名空间包，`tests.fixtures` 找不到。新增空的 `tests/__init__.py`。

### 修复后重跑 F01/F11（提交 `597b791`，[`p6-font-check-rerun.json`](acceptance/public-release/windows-word/p6-font-check-rerun.json)）

- **F01**：离线全量 502 项通过、9 项跳过（比第一次多 7 项，即此前没能导入的 `test_docx_inspector`）。
- **F11**：含真实 Word 用例 539 项通过、1 项跳过。

第 1 节第 2–6 条在 Windows 上满足。未覆盖：真正缺少 `仿宋_GB2312` 时 Word 的替换结果与分页变化（第 5 节第 4 项，另一个 Windows 账户），本批没有做，不写成已验证。

F5 完成。

## 15. F6 结果：文档（2026-10-18）

- README 新增“字体”一节：仓库不分发字体文件；系统自带与需授权的公文字体分别怎样获取；`--plan`/`--doctor` 的字体报告；`--font-fallbacks` 的用法、内置回退列表、v3 声明位置、候选全缺时失败，以及检查时须加同样参数。
- 只写已验证的行为。真正缺少 `仿宋_GB2312` 时 Word 的替换与分页（第 5 节第 4 项）仍未验证，README 只写“分页可能与预期不同”。

P6 完成。
