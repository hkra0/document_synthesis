# P6 实施计划：字体可移植性

> **文档状态**：计划中；F0（Windows 字体探针）脚本已提供，尚未在 Windows 上运行，其余步骤未实施  
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

### 3.2 `--plan` 与 `--doctor`

- `--plan`：在构建计划后增加“字体”一节，逐个列出状态和来源（格式包样式名 / 渲染器）。只读，不启动 Office。
- `--doctor`：增加同样的检查；缺失字体计为警告，不改变现有退出码含义（0/2/1）。Windows + Word 时额外用 `FontNames` 交叉核对，两者不一致时如实列出，不自行裁决。

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
4. **干净系统上的识别（可选）**：Windows 11 Pro 自带 Windows Sandbox，里面是全新系统、没有公文字体、也没有 Word。可以在其中运行 `--plan`/`--doctor` 与探针，确认 `仿宋_GB2312` 等报告为缺失、系统自带的 `仿宋`/`楷体`/`黑体`/`宋体` 报告为已安装。需要在“启用或关闭 Windows 功能”中开启，属于可选项。

本批不能覆盖的部分：真正缺少 `仿宋_GB2312` 时 Word 的替换结果与分页变化。第 2 项中 Word 替换的是一个虚构字体名，替换规则可能与缺少真实公文字体时不同。这一点在状态记录中如实写明，不写成已验证。

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
