# 2026-09-06 真实 Word 证据

本目录保存本轮 post-review 的可复核产物。执行环境为 macOS arm64、Microsoft Word 16.112.3、Python 3.14.0，使用生产 CLI、`DOCUMENT_SYNTHESIS_WORD_TEST=1` 和固定目录 `DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir>`。

## N8

- NW01–NW12 最终矩阵 12 项全部通过（72.668 秒），没有新的授权弹窗。
- 随机临时目录版本曾使 NW02 发生一次 Word→PDF 导出超时；固定目录复跑已消除该重复授权/稳定性问题。
- 机器可读结果见 [`../../n8-status.json`](../../n8-status.json)。
- 20 页 NW09 交付物为 `nw09-notes.docx`、`nw09-qa-main.pdf` 和 `nw09-build-metadata.json`。

## N9

- `external-export-source.docx` 是同源真实 Word 导出探针的 DOCX。
- `external-export-qa-main-rerun.pdf` 是使用同一匿名 DOCX 经生产 CLI 和真实 Word 重导出的独立探针；三种解析器均读取 6 页，页面框一致。
- pypdf 报告的 3 个 Quartz wrong pointing object offset 警告对应 xref 中不存在且无引用的对象 6、8、10；原始 stderr 和 xref 证据均保留，N9 已完成非破坏性诊断分类。
- 机器可读结果见 [`../../n9-status.json`](../../n9-status.json)。

文件 SHA-256 以 [`../../n8-status.json`](../../n8-status.json) 的 `artifact_sha256` 和 [`../../n9-status.json`](../../n9-status.json) 的 `artifact_sha256` 为准。
