# Formatting Test Fixtures

本目录包含自定义排版、样式继承、OPC 导入与复杂对象保留的基准测试样本和规范说明。

## 设计与使用原则

1. **零业务材料**：所有样本必须为虚构内容、匿名样本或测试生成器创建的微型 DOCX。严禁将用户隐私或业务材料放入本目录。
2. **只读与隔离**：测试执行期间严禁修改本目录中的持久化文件。动态测试样本只能生成到 `TemporaryDirectory` 中。
3. **样本类型覆盖**：
   - `standard_styles.docx`: 使用标准内置样式（Title, Heading 1-3, Normal）的规范文档。
   - `manual_formatting.docx`: 全文使用 Normal 样式但通过直接格式（Direct Formatting: 字体、字号、加粗、行距）实现层次的文档。
   - `mixed_fonts.docx`: 包含中西文混合字体（w:rFonts eastAsia, ascii, hAnsi, cs）、字符样式及直接格式覆盖的文档。
   - `complex_fields.docx`: 包含跨段复杂字段（w:fldChar begin/separate/end，如 PAGE/NUMPAGES, TOC, STYLEREF）与简单字段（w:fldSimple）的文档。
   - `footnotes_sample.docx`: 包含脚注（footnotes.xml）及正文引用的文档。
   - `multilevel_list.docx`: 包含原生多级编号列表（w:num, w:abstractNumId）的文档。
   - `tables_and_images.docx`: 包含表格条件样式、嵌入图片与横纵混排分节的文档。
