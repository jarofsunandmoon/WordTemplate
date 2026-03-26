# WordTemplate

基于现有 `docx` 模板生成 Word 报告。

当前实现采用“固定模板 + 结构化内容替换”的方式：

- `templates/base.docx`：原始样式、目录、页眉页脚、页码、关系文件
- `templates/document.xml`：从 `docx` 中提取的正文 XML 模板
- `templates/header1.xml` / `templates/header2.xml` / `templates/footer1.xml`：页眉页脚 XML 模板
- `templates/document_structure.xml`：固定章节顺序配置
- `src/data_loader.py`：输出结构化数据
- `src/generator.py`：按章节替换正文段落与表格，并重新打包为 `docx`

说明：

- 目录中的标题、子标题保持模板固定
- 动态替换封面大标题、正文内容和指定表格数据
- 如果源文件还是 `.doc`，请先在 Word/WPS 中另存为 `.docx`
