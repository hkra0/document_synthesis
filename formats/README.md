# Document Synthesis Formats & Presets

本目录存放通用的版本化排版格式包（Format Packages）与官方内置预设（Presets）。

## 目录组织

- `presets/`: 内置通用标准格式包
  - `legacy-official/1.0.0.json`: 兼容现有公文标准（GB/T 9704 风格版式）
  - `report-basic/1.0.0.json`: 现代通用商务与技术报告版式
  - `academic-basic/1.0.0.json`: 基础学术论文规范版式（宋体/Times New Roman、1.5 倍行距、首行 2 字符缩进）

## 引用机制

项目清单（v3）可通过以下方式引用格式包：
1. **预设引用**：`"ref": "preset:report-basic@1.0.0"`
2. **相对路径**：`"ref": "formats/custom-style.json"`

格式包自身支持通过 `extends` 单继承父格式包，例如：
```json
{
  "format_schema_version": 1,
  "id": "my-academic",
  "version": "1.0.0",
  "extends": "preset:academic-basic@1.0.0",
  "styles": {
    "body": {
      "run": { "size_pt": 12.5 }
    }
  }
}
```

## 发布约定

- 格式包必须带 `format_schema_version`、稳定 `id`、语义化 `version`，并在 `extends` 中固定父包版本；分析报告中的样例正文不进入格式包。
- v3 项目清单必须带 `schema_version: 3`、版本化 `format.ref` 和显式 `formatting.mode`；需要人工确认的 RoleMap 应绑定源文件哈希与 `on_unmapped` 策略。
- v1/v2 清单迁移只写入独立目标文件，不覆盖输入或默默删除旧交付物。可运行的完整示例见 [`examples/custom-format-demo/README.md`](../examples/custom-format-demo/README.md)，验收记录见 [`docs/acceptance/r11-status.json`](../docs/acceptance/r11-status.json)。
