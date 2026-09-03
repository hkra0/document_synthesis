# 最小匿名演示

该示例仅包含虚构的流程说明和操作清单，不含个人、单位、地点、联系方式或业务数据。它用于验证统一入口能加载显式配置并生成可审阅的大纲计划。

在项目根目录执行只读验证：

```bash
python3 synthesize.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --plan
```

如需重新生成两个匿名 DOCX 源材料，使用：

```bash
python3 examples/minimal-demo/scripts/create_demo_materials.py
```

完整构建仍需要本机 Microsoft Word Automation；先运行 `--doctor`，确认权限和环境均通过后再构建。
