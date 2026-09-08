# 最小匿名演示

该示例仅包含虚构的流程说明和操作清单，不含个人、单位、地点、联系方式或业务数据。它使用 v2 配置，默认声明一份 `demo_complete.docx`，包含封面、目录和正文。

在项目根目录执行只读验证：

```bash
python3 synthesize.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --plan
```

小众的拆分交付只需追加一份覆写文件，不修改源材料或主配置：

```bash
python3 synthesize.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --override examples/minimal-demo/split.override.json --plan
```

此时只声明 `demo_cover_toc.docx` 和 `demo_toc_body.docx`；独立目录引用后者的正文打印页码，不写无效跳转链接。省略覆写即恢复完整文档。不同配置使用同一输出目录时，旧文件不会自动删除；可以分别指定 `--output-dir output/demo-full` 与 `--output-dir output/demo-split`。

如需重新生成两个匿名 DOCX 源材料，使用：

```bash
python3 examples/minimal-demo/scripts/create_demo_materials.py
```

完整构建仍需要本机 Microsoft Word Automation；先检查权限和环境，再移除 `--plan` 构建：

```bash
python3 synthesize.py --doctor --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json
python3 synthesize.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --output-dir output/demo-full
python3 smoke_test.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --output-dir output/demo-full
```

无 Office 时可以运行只读计划和单元测试。已有输出可用 `smoke_test.py --structure-only` 做有限的结构与文字核验，但这不等同于真实 Word 页码与版式验收。构建默认不保留中间正文骨架；诊断时可加 `--keep-work`。
