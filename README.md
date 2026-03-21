# cugbWord
用于格式化中地大（京）的毕业论文格式

## 样式配置

`win32com_demo.py` 现在会从 `style_config.json` 读取样式配置，再将配置应用到 Word 文档的内置样式上。

当前配置文件包含：
- `document_path`：要处理的 `.docx` 文件路径
- `styles`：样式列表

每个样式项需要包含：
- `style_id`：内部标识
- `builtin_names.english` / `builtin_names.chinese`：Word 内置样式的中英文名称
- `font`：字体相关配置
- `paragraph`：段落相关配置

## 扩展方式

如果要新增或调整样式，优先修改 `style_config.json`，不需要改 Python 主流程。

为了兼容不同语言版本的 Word，请为每个内置样式同时提供英文名和中文名。

当前模块分工：
- `win32com_demo.py`：主流程入口
- `style_operations.py`：样式定义更新、直接段落格式覆盖
- `paragraph_utils.py`：段落文本清洗与前后段落访问
- `paragraph_rules.py`：标题、图注、表注、图片段落、结构标题识别
- `paragraph_processing.py`：遍历段落、应用识别结果并执行摘要/关键词内容校验

当前已支持的论文结构识别：
- 中文摘要
- 关键词
- 参考文献
- 目录
- 致谢
- 附录

说明：
- 当前对 `目录` 只做保守识别和样式入口，不包含自动生成目录、自动更新目录、目录点线和页码对齐优化。
