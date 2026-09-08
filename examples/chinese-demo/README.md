# 中文示例

两份中文源材料合成一份完整 Word 文档，包含封面、可跳转目录和连续编页的正文。

封面为 21 × 28 cm 的 3∶4 竖版，采用红色小标宋主标题与黑色功能副标题，便于导出图片用于社交媒体。目录和正文保持 A4 版式。此示例借用公文的字体和层次，不代表符合某项公文格式标准。

字体使用方正小标宋简体、仿宋 GB2312、楷体 GB2312 和黑体。配置中使用字体的英文族名，便于不同软件识别。请先确认这些字体已在系统中启用，仅把字体文件放入文件夹可能仍会触发字体替换。仓库不分发字体文件。

## 预览与生成

在仓库根目录执行。完整构建需要本机 Microsoft Word 及项目依赖。

```bash
python3 synthesize.py --source examples/chinese-demo/source --manifest examples/chinese-demo/manifest.json --plan
python3 synthesize.py --source examples/chinese-demo/source --manifest examples/chinese-demo/manifest.json --doctor
python3 synthesize.py --source examples/chinese-demo/source --manifest examples/chinese-demo/manifest.json --output-dir output/中文示例
```

## 修改文案

修改 `copy.json`，再生成源材料和封面模板。若更改章标题，同步调整清单中的 `title`。

```bash
python3 examples/chinese-demo/scripts/create_demo_materials.py
```

生成脚本只准备示例输入，最终文档始终由统一入口 `synthesize.py` 构建。页面尺寸由封面模板和现有配置表达，无需修改引擎。

如需拆成交付两份文档，可在构建命令末尾追加示例的配置覆写。

```bash
--override examples/chinese-demo/split.override.json
```

默认完整文档与可选拆分交付共用同一套中文源材料。
