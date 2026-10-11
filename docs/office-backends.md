# Office 后端对照

本文记录 macOS（`mac-word`）与 Windows（`win-word`）两个精确后端在调用方式和导出参数上的差异，依据 [Windows 支持与公开发布执行计划](windows-and-public-release-plan.md) 第 3.3 节与第 8 节。两者都实现 `lib/office/base.py` 中的 `OfficeBackend` 协议，保真度都是 `exact`；`lib/pagination.py` 的 `inspect_document` 会拒绝任何非 `exact` 结果。

## 1. 调用方式

| 项目 | macOS `mac-word` | Windows `win-word` |
| --- | --- | --- |
| 实现文件 | `lib/office/mac_applescript.py` | `lib/office/win_com.py`（父进程）、`lib/office/win_com_worker.py`（子进程） |
| 驱动 | `osascript` 执行 AppleScript | pywin32 COM，`DispatchEx("Word.Application")` |
| Word 实例 | 用户正在运行的 Word（AppleScript 只能连接单实例） | 每次操作启动一个独立、隐藏的 Word 实例，结束时 `Quit` |
| 隔离 | 每次一个 `osascript` 子进程 | 每次一个 worker 子进程，stdin/stdout 各一行 JSON，日志走 stderr |
| 超时 | AppleScript `with timeout of 600 seconds`，`subprocess` 620 秒 | 父进程 600 秒（`DOCUMENT_SYNTHESIS_OFFICE_TIMEOUT` 可调），超时后结束 worker，并按 pid 文件结束本次启动的 `WINWORD.EXE` |
| 静态检查 | `/usr/bin/osascript` 与 `/Applications/Microsoft Word.app` 是否存在 | 注册表 `HKCR\Word.Application\CLSID` 与 pywin32 是否可导入 |
| `--doctor` 探针 | `tell application "Microsoft Word" to get name` | 启动隔离实例读取 `Version`、`Build`、`Path`（推断位数与 Click-to-Run）后退出 |
| 工作目录 | 固定的 `/private/tmp/document-synthesis-word-access`（避免每次弹出文件夹授权） | `%TEMP%\document-synthesis-word-access\job-<随机>`，每次操作结束删除 |
| 宏 | 沿用 Word 设置 | `AutomationSecurity = 3`（强制禁用） |

`DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR` 在两个平台上都可以覆盖工作目录。

## 2. 书签页码

| 项目 | macOS | Windows |
| --- | --- | --- |
| 重新分页 | `repaginate` | `doc.Repaginate()` |
| 定位 | `create range ... start bookmarkStart end bookmarkStart` | `doc.Range(start, start)`，`start = Bookmarks(name).Range.Start` |
| 物理页 | `active end page number` | `Information(3)`（`wdActiveEndPageNumber`） |
| 打印页码 | `active end adjusted page number` | `Information(1)`（`wdActiveEndAdjustedPageNumber`） |
| 隐藏书签 | 按名访问 | 先设 `Bookmarks.ShowHidden = True`，并用 `Bookmarks.Exists` 报告缺失书签 |
| 输出格式 | `名\t物理页\t打印页\n`，由 `_parse_page_map` 严格解析 | 同左 |

## 3. PDF 导出

| 参数 | macOS（`save as ... file format format PDF`） | Windows（`ExportAsFixedFormat`） |
| --- | --- | --- |
| 范围 | 全文 | `Range=0`（`wdExportAllDocument`） |
| 内容 | 文档内容 | `Item=0`（`wdExportDocumentContent`） |
| 优化 | Word 默认 | `OptimizeFor=0`（打印） |
| PDF 书签 | Word for Mac 的默认行为 | `CreateBookmarks=1`（按标题） |
| 结构标签 | Word for Mac 的默认行为 | `DocStructureTags=True` |
| 缺失字体 | Word for Mac 的默认行为 | `BitmapMissingFonts=True` |
| PDF/A | 否 | `UseISO19005_1=False` |

Word for Mac 的 AppleScript `save as` 不提供上述细项，只能取默认值；两平台的 PDF 大纲是否一致，在 P4 跨平台对照时记录。页级 QA 只读取页面文字和页码标签，不依赖 PDF 大纲。

## 4. 转换

| 操作 | macOS | Windows |
| --- | --- | --- |
| `.doc` → `.docx` | `save as ... file format format document default` | `SaveAs2(FileFormat=16)`（`wdFormatDocumentDefault`），只读打开工作目录中的副本 |
| `.pptx` → PDF | PowerPoint AppleScript `save as PDF` | `Presentations.Open(src, ReadOnly, Untitled=False, WithWindow=False)` → `SaveAs(dst, 32)` |
| PowerPoint 缺失 | AppleScript 报错 | 注册表中没有 `PowerPoint.Application` 时，只有 `.pptx` 转换报 `unsupported`，其他功能不受影响 |

PowerPoint 是单实例应用：用户已打开 PowerPoint 时，Windows 后端只关闭本次打开的演示文稿，不调用 `Quit`，也不会在超时时结束 PowerPoint 进程。

## 5. 失败与清理（Windows）

- Word 只读打开工作目录中的私有副本；副本的 `Zone.Identifier`（“来自网络”标记）会被移除，避免以受保护视图打开。
- 产物先写在工作目录，全部检查通过后，复制到目标所在目录的临时文件，再在同一目录内原子替换（工作目录与目标可能不在同一盘符）。目标被占用时按退避重试，最终失败时保留原目标文件。
- worker 启动 Word 前写下同名进程快照，取得实例后改写为该实例的 PID、映像名和创建时间。父进程只结束与记录完全一致的进程；Word 卡在启动阶段、来不及记录 PID 时，只有“快照之后新出现、命令行带 `/Automation -Embedding`（由 COM 启动）、且唯一”的 `WINWORD.EXE` 才会被结束；不以是否有可见窗口判断，因为插件弹窗会让自动化实例也出现可见窗口。worker 按窗口标题取不到 PID 时，也用同一规则确定本实例。任何情况下都不按进程名结束 Word。
- 常见 HRESULT 翻译为中文提示：未安装（`0x800401F3`、`0x80040154`）、无法启动或卡在首次运行/激活（`0x80080005`）、对象被占用（`0x8001010A`、`0x80010001`）、进程意外退出（`0x800706BA`、`0x800706BE`、`0x80010108`）、拒绝访问（`0x80070005`），以及按错误描述识别的受保护视图与未激活。
- worker 注册 COM 消息过滤器，Word 短暂忙碌（`RETRYLATER`）时自动重试，最长约 60 秒。
