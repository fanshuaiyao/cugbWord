# cugbWord - 毕业论文 Word 排版自动化

用于自动化处理毕业论文的 Word 排版，让你从繁琐的格式调整中解放出来。

## 项目现阶段能力

本项目目前分为两个阶段：

### ✅ 阶段一：自动创建预设样式（已完成）

程序会自动在 Word 文档中创建/更新以下样式定义：

| 样式用途 | 样式ID | 说明 |
|---------|--------|------|
| 一级标题 | `heading_1` | 支持 `1.`、`一、` 等格式 |
| 二级标题 | `heading_2` | 支持 `1.1` 格式 |
| 三级标题 | `heading_3` | 支持 `1.1.1` 格式 |
| 四级标题 | `heading_4` | 支持 `1.1.1.1` 格式 |
| 正文 | `normal` | 默认正文样式 |
| 图注/表注 | `caption` | 图表标题说明 |
| 图片段落 | `figure_block` | 图片所在段落 |
| 表格正文 | `table_body` | 表格内文字 |
| 中文摘要标题 | `abstract_title` | 如"摘 要" |
| 中文摘要正文 | `abstract_body` | 摘要内容 |
| 中文关键词 | `keywords_line` | 关键词行 |
| 英文摘要标题 | `english_abstract_title` | 如"Abstract" |
| 英文摘要正文 | `english_abstract_body` | 英文摘要内容 |
| 英文关键词 | `english_keywords_line` | Keywords行 |
| 参考文献标题 | `references_title` | 如"参考文献" |
| 参考文献条目 | `reference_entry` | 参考文献内容 |
| 目录标题 | `contents_title` | 如"目录" |
| 目录条目 | `contents_entry` | 目录内容 |
| 致谢标题/正文 | `acknowledgements_title/body` | 致谢部分 |
| 附录标题/正文 | `appendix_title/body` | 附录部分 |

同时还会自动设置：
- **页面设置**：页边距、页眉页脚距离
- **页眉页脚**：首页不同选项

### 🚧 阶段二：智能应用样式（进行中）

程序会遍历文档段落，根据内容特征自动识别并套用对应样式：

- 识别标题层级（根据编号格式 `1.`、`1.1` 等）
- 识别摘要、关键词、参考文献等论文结构
- 识别图表标题和图片段落
- 对摘要/关键词进行内容校验（字数、标点等）

**如何开启/关闭阶段二**：

在 `runtime_config.json` 中设置：
```json
{
  "processing": {
    "apply_paragraph_styles": true   // true = 执行智能应用，false = 仅创建样式
  }
}
```

---

## 如果你是地大（北京）的学生

地大北京是本项目默认支持的学校，已经配置好了符合学校要求的样式参数。

### 使用方法

1. **准备环境**
   - Windows 系统
   - 已安装 Microsoft Word
   - 安装 Python 依赖：`pip install pywin32`

2. **配置要处理的文档**

   编辑 `runtime_config.json`：
   ```json
   {
     "document_path": "C:\\Users\\你的用户名\\Desktop\\你的论文.docx",
     "style_template": "cugb",
     "processing": {
       "apply_paragraph_styles": true
     }
   }
   ```

3. **运行程序**

   ```bash
   python win32com_demo.py
   ```

4. **查看结果**

   程序会在原文件同目录生成 `_处理后.docx` 文件，打开检查格式是否正确。

### 地大样式包含的特殊设置

- 标题 1：黑体 15pt，居中，加粗
- 标题 2~4：黑体 14pt/12pt，左对齐/两端对齐
- 正文：宋体 12pt，首行缩进 2 字符，1.5 倍行距
- 摘要/参考文献等结构：已预设符合地大格式的样式

---

## 如果你不是地大的学生

本项目支持通过配置文件适配其他学校的格式要求。

### 快速开始（推荐）

**找一个最接近的现有模板进行修改**：

1. **复制模板文件**

   将 `style/cugb.json` 复制一份，重命名为你们学校的代号，例如：
   ```
   style/njfu.json    // 南京林业大学
   style/swufe.json   // 西南财经大学
   ```

2. **修改样式参数**

   打开复制的 JSON 文件，修改以下关键参数为你学校的要求：

   ```json
   {
     "page_setup": {
       "margins_cm": {
         "top": 2.5,      // 上边距（厘米）
         "bottom": 2.0,   // 下边距
         "left": 2.5,     // 左边距
         "right": 2.0     // 右边距
       }
     },
     "styles": [
       {
         "style_id": "heading_1",
         "font": {
           "name_far_east": "黑体",    // 中文字体
           "size": 15,                   // 字号（pt）
           "bold": true                  // 是否加粗
         },
         "paragraph": {
           "alignment": "center"         // 对齐方式：left/center/justify
         }
       },
       {
         "style_id": "normal",
         "font": {
           "name_far_east": "宋体",
           "size": 12
         },
         "paragraph": {
           "first_line_indent_chars": 2   // 首行缩进字符数
         }
       }
       // ... 其他样式
     ]
   }
   ```

3. **切换到你们的模板**

   编辑 `runtime_config.json`：
   ```json
   {
     "document_path": "你的论文.docx",
     "style_template": "njfu",    // 改成你的学校代号
     "processing": {
       "apply_paragraph_styles": true
     }
   }
   ```

4. **运行测试**

   ```bash
   python win32com_demo.py
   ```

### 常见问题

**Q: 我们学校要求特殊的段落结构（如独创性声明、版权使用授权书），怎么办？**

A: 在模板 JSON 的 `styles` 数组中添加自定义样式：
```json
{
  "style_id": "originality_statement",
  "builtin_names": {
    "english": "Originality Statement",
    "chinese": "独创性声明"
  },
  "font": {
    "name_ascii": "Times New Roman",
    "name_far_east": "宋体",
    "size": 12,
    "bold": false
  },
  "paragraph": {
    "alignment": "justify",
    "first_line_indent_chars": 2
  }
}
```

然后在 `paragraph_rules.py` 中添加对应的识别规则（需要懂一点 Python）。

**Q: 只想用阶段一（创建样式），暂时不想用阶段二的智能识别？**

A: 在 `runtime_config.json` 中设置：
```json
{
  "processing": {
    "apply_paragraph_styles": false
  }
}
```

这样程序只会把样式定义写入文档，不会改变任何段落内容。然后你可以**手动在 Word 里选择段落并应用样式**。

**Q: 如何验证配置是否正确？**

A: 运行程序后，打开生成的 `_处理后.docx`，在 Word 中：
1. 点击"开始"→"样式"右下角的小箭头
2. 检查各样式的字体、段落设置是否符合学校要求
3. 直接选中段落应用对应样式查看效果

---

## 配置结构说明

### 1. 运行配置 `runtime_config.json`

控制"这次处理哪个文档、用什么模板、执行哪些流程"。

```json
{
  "document_path": "test/你的论文.docx",    // 要处理的文档路径
  "style_template": "cugb",                  // 使用的学校模板
  "processing": {
    "apply_paragraph_styles": true           // 是否执行智能段落匹配
  }
}
```

### 2. 模板配置 `style/*.json`

控制"某个学校的具体格式要求"。

当前内置模板：
- `style/cugb.json` - 中国地质大学（北京）
- `style/gznfxy.json` - 广州南方学院

---

## 注意事项

1. **备份原文档**：程序会生成新文件（`_处理后.docx`），不会直接修改原文档
2. **需要 Word 环境**：必须在安装了 Microsoft Word 的 Windows 系统上运行
3. **样式识别不是 100% 准确**：复杂排版建议先用阶段一创建样式，然后手动微调
4. **字体依赖系统**：确保系统安装了所需字体（如宋体、黑体、Times New Roman）

---

## 技术栈

- Python 3.x
- pywin32（COM 接口调用 Word）
- JSON 配置驱动

## 项目状态

- ✅ 样式定义与页面设置自动化
- ✅ 论文结构识别（摘要、参考文献等）
- ✅ 标题层级自动识别
- ✅ 多学校模板支持
- 🚧 段落智能匹配持续优化中
- ⏳ 目录自动生成（规划中）
- ⏳ 页码自动处理（规划中）
