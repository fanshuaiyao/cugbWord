# style_config 说明版

## 文件作用

`style_config.json` 是当前项目的样式配置文件。

程序会按下面的流程使用它：

1. 读取 `document_path`
2. 读取 `processing`
3. 读取 `page_setup` 和 `header_footer`
4. 打开对应的 Word 文档
5. 读取 `styles` 里的样式定义
6. 先更新 Word 内置样式和自定义样式本身
7. 再应用页面设置与页眉页脚
8. 最后根据开关决定是否继续做段落匹配

注意：
- `style_config.json` 必须是**严格 JSON**
- 不能直接写 `//` 或 `#` 注释
- 所以说明信息单独放在本文件里

---

## 顶层结构

```json
{
  "document_path": "test/123.docx",
  "processing": {
    "apply_paragraph_styles": true
  },
  "page_setup": {
    "enabled": true,
    "margins_cm": {
      "top": 2.5,
      "bottom": 2.0,
      "left": 2.5,
      "right": 2.0
    },
    "header_distance_cm": 1.5,
    "footer_distance_cm": 1.5
  },
  "header_footer": {
    "enabled": true,
    "different_first_page": true,
    "header": {
      "enabled": false,
      "text": "",
      "style_ref": "thesis_header"
    },
    "footer": {
      "enabled": false,
      "text": "",
      "style_ref": "thesis_footer"
    },
    "first_page": {
      "header": {
        "enabled": false,
        "text": "",
        "style_ref": "thesis_header"
      },
      "footer": {
        "enabled": false,
        "text": "",
        "style_ref": "thesis_footer"
      }
    }
  },
  "styles": [
    { ... }
  ]
}
```

### 顶层字段说明

| 字段 | 含义 | 当前示例 |
|---|---|---|
| `document_path` | 目标 Word 文档路径。可以是相对路径，也可以是绝对路径。相对路径会基于项目目录解析。 | `test/123.docx` |
| `processing` | 主流程控制配置。当前用于决定样式定义更新后是否继续执行段落匹配。 | `{ "apply_paragraph_styles": true }` |
| `page_setup` | 页面设置配置。当前用于统一页边距和页眉页脚距离。 | `{ "enabled": true, ... }` |
| `header_footer` | 页眉页脚配置。当前用于首页不同与页眉页脚样式入口。 | `{ "enabled": true, ... }` |
| `styles` | 样式配置列表。每一项定义一个 Word 内置样式或自定义段落样式。 | 23 个样式项 |

### processing 字段说明

| 字段 | 含义 | 默认值 |
|---|---|---|
| `processing.apply_paragraph_styles` | 是否在样式定义更新后继续执行段落匹配与内容校验。`false` 时只更新样式定义并保存处理后副本。 | `true` |

### page_setup 字段说明

| 字段 | 含义 | 默认值 |
|---|---|---|
| `page_setup.enabled` | 是否启用页面设置。 | `false` |
| `page_setup.margins_cm.top` | 上边距，单位厘米。 | `2.5` |
| `page_setup.margins_cm.bottom` | 下边距，单位厘米。 | `2.0` |
| `page_setup.margins_cm.left` | 左边距，单位厘米。 | `2.5` |
| `page_setup.margins_cm.right` | 右边距，单位厘米。 | `2.0` |
| `page_setup.header_distance_cm` | 页眉距离页边界，单位厘米。 | `1.5` |
| `page_setup.footer_distance_cm` | 页脚距离页边界，单位厘米。 | `1.5` |

### header_footer 字段说明

| 字段 | 含义 | 默认值 |
|---|---|---|
| `header_footer.enabled` | 是否启用页眉页脚处理。 | `false` |
| `header_footer.different_first_page` | 是否启用首页不同。 | `true` |
| `header_footer.header.enabled` | 是否写入非首页页眉内容。 | `false` |
| `header_footer.header.text` | 非首页页眉文本。当前留空时只保留样式入口，不写内容。 | `""` |
| `header_footer.header.style_ref` | 非首页页眉引用的样式 `style_id`。 | `thesis_header` |
| `header_footer.footer.enabled` | 是否写入非首页页脚内容。 | `false` |
| `header_footer.footer.text` | 非首页页脚文本。当前留空时只保留样式入口，不写内容。 | `""` |
| `header_footer.footer.style_ref` | 非首页页脚引用的样式 `style_id`。 | `thesis_footer` |
| `header_footer.first_page.header.enabled` | 是否写入首页页眉内容。 | `false` |
| `header_footer.first_page.header.text` | 首页页眉文本。 | `""` |
| `header_footer.first_page.header.style_ref` | 首页页眉引用的样式 `style_id`。 | `thesis_header` |
| `header_footer.first_page.footer.enabled` | 是否写入首页页脚内容。 | `false` |
| `header_footer.first_page.footer.text` | 首页页脚文本。 | `""` |
| `header_footer.first_page.footer.style_ref` | 首页页脚引用的样式 `style_id`。 | `thesis_footer` |

---

## 单个样式对象结构

```json
{
  "style_id": "heading_1",
  "builtin_names": {
    "english": "Heading 1",
    "chinese": "标题 1"
  },
  "font": {
    "name_ascii": "Times New Roman",
    "name_far_east": "黑体",
    "size": 15,
    "bold": false,
    "color": "black",
    "color_index": "black"
  },
  "paragraph": {
    "alignment": "center",
    "outline_level": 1,
    "first_line_indent_chars": 0,
    "line_spacing_rule": "1.5_lines",
    "space_before": 0,
    "space_after": 0
  }
}
```

---

## style_id

| 字段 | 含义 |
|---|---|
| `style_id` | 样式的内部标识。Python 代码会根据这个值决定把某个段落设置成哪个样式。 |

### 当前已使用的 `style_id`

| style_id | 含义 |
|---|---|
| `heading_1` | 一级标题 |
| `heading_2` | 二级标题 |
| `heading_3` | 三级标题 |
| `heading_4` | 四级标题 |
| `normal` | 正文 |
| `caption` | 图注 / 表注共用题注 |
| `figure_block` | 图片所在段落 |
| `abstract_title` | 中文摘要标题 |
| `abstract_body` | 中文摘要正文 |
| `keywords_line` | 中文关键词整行 |
| `english_abstract_title` | 英文摘要标题 |
| `english_abstract_body` | 英文摘要正文 |
| `english_keywords_line` | 英文关键词整行 |
| `references_title` | 参考文献标题 |
| `reference_entry` | 参考文献条目 |
| `contents_title` | 目录标题 |
| `contents_entry` | 目录条目 |
| `acknowledgements_title` | 致谢标题 |
| `acknowledgements_body` | 致谢正文 |
| `appendix_title` | 附录标题 |
| `appendix_body` | 附录正文 |
| `thesis_header` | 论文页眉 |
| `thesis_footer` | 论文页脚 |

---

## builtin_names

`builtin_names` 用来告诉程序：这个配置对应的是 Word 里的哪个**内置样式**。

| 字段 | 含义 |
|---|---|
| `builtin_names.english` | Word 英文界面的内置样式名 |
| `builtin_names.chinese` | Word 中文界面的内置样式名 |

### 示例

| 配置 | 含义 |
|---|---|
| `Heading 1` / `标题 1` | Word 的一级标题内置样式 |
| `Heading 2` / `标题 2` | Word 的二级标题内置样式 |
| `Normal` / `正文` | Word 的正文内置样式 |
| `Caption` / `题注` | Word 的题注内置样式 |

程序会优先尝试英文名，找不到时再尝试中文名。

说明：
- `figure_block` 当前使用独立的自定义段落样式 `正文图片`，由代码在 Word 中获取或创建。
- `abstract_title`、`abstract_body`、`keywords_line`、`english_abstract_title`、`english_abstract_body`、`english_keywords_line`、`references_title`、`reference_entry`、`contents_title`、`contents_entry`、`acknowledgements_title`、`acknowledgements_body`、`appendix_title`、`appendix_body`、`thesis_header`、`thesis_footer` 也属于自定义段落样式，由代码在 Word 中获取或创建。
- 这意味着这些结构样式和页面样式都是独立入口，不再与 `Normal / 正文` 或内置标题样式共用样式对象。

---

## font 字段说明

`font` 控制字体格式。

| 字段 | 含义 | 备注 |
|---|---|---|
| `name_ascii` | 西文、数字、拉丁字符的字体 | 例如 `Times New Roman` |
| `name_far_east` | 中文字符的字体 | 例如 `黑体`、`宋体` |
| `size` | 字号，单位是 pt | 例如 `10.5`、`12`、`14`、`15` |
| `bold` | 是否加粗 | `true` 表示加粗，`false` 表示不加粗 |
| `color` | 字体颜色 | 当前代码只支持 `black` |
| `color_index` | Word 颜色索引 | 当前代码只支持 `black` |

### 当前字号对照

| pt 值 | 常见中文字号 |
|---|---|
| `15` | 小三 |
| `14` | 四号 |
| `12` | 小四 |
| `10.5` | 五号 |

---

## paragraph 字段说明

`paragraph` 控制段落格式。

| 字段 | 含义 | 备注 |
|---|---|---|
| `alignment` | 对齐方式 | 当前支持 `left`、`center`、`justify` |
| `outline_level` | Word 大纲级别 | 1~4 对应标题级别，10 通常用于正文、题注或图片段落 |
| `first_line_indent_chars` | 首行缩进字符数 | 标题、题注和图片段落一般为 0，正文常见为 2 |
| `line_spacing_rule` | 行距规则 | 当前只支持 `1.5_lines` |
| `space_before` | 段前距离 | 当前数值直接写入 Word |
| `space_after` | 段后距离 | 当前数值直接写入 Word |

### alignment 说明

| 值 | 含义 |
|---|---|
| `left` | 左对齐 |
| `center` | 居中 |
| `justify` | 两端对齐 |

---

## 当前样式分别表示什么

### 1. `heading_1`
- 对应 Word 内置样式：`Heading 1` / `标题 1`
- 当前基于样式：`Normal` / `正文`
- 当前字体：`Times New Roman + 黑体`
- 当前字号：`15`
- 当前对齐：`center`
- 当前大纲级别：`1`

当前代码会把这些段落识别成它：
- `1. 绪论`
- `一、绪论`

### 2. `heading_2`
- 对应 Word 内置样式：`Heading 2` / `标题 2`
- 当前基于样式：`Normal` / `正文`
- 当前字体：`Times New Roman + 黑体`
- 当前字号：`14`
- 当前对齐：`left`
- 当前大纲级别：`2`

当前代码会把这些段落识别成它：
- `1.1 研究背景`

### 3. `heading_3`
- 对应 Word 内置样式：`Heading 3` / `标题 3`
- 当前基于样式：`Normal` / `正文`
- 当前字体：`Times New Roman + 黑体`
- 当前字号：`14`
- 当前对齐：`justify`
- 当前大纲级别：`3`

当前代码会把这些段落识别成它：
- `1.1.1 国内研究现状`

### 4. `heading_4`
- 对应 Word 内置样式：`Heading 4` / `标题 4`
- 当前基于样式：`Normal` / `正文`
- 当前字体：`Times New Roman + 黑体`
- 当前字号：`12`
- 当前对齐：`justify`
- 当前大纲级别：`4`

当前代码会把这些段落识别成它：
- `1.1.1.1 数据来源`

### 5. `normal`
- 对应 Word 内置样式：`Normal` / `正文`
- 当前字体：`Times New Roman + 宋体`
- 当前字号：`12`
- 当前对齐：`justify`
- 当前大纲级别：`10`
- 当前首行缩进：`2`

当前代码会把**未匹配到标题、题注或图片段落规则的非空段落**识别成它。

### 6. `caption`
- 对应 Word 内置样式：`Caption` / `题注`
- 当前字体：`Times New Roman + 宋体`
- 当前字号：`10.5`
- 当前对齐：`center`
- 当前大纲级别：`10`
- 当前首行缩进：`0`
- 当前行距：`1.5_lines`

当前代码会把图注和表注都识别成它。

### 7. `figure_block`
- 当前用于图片所在段落
- 当前效果：居中、无缩进、段前一行、段后 0
- 段前一行的高度按正文字号动态设置

当前代码会把**包含行内图片的段落**识别成它。

### 8. `abstract_title`
- 当前用于中文摘要标题段
- 当前效果：黑体、小三、居中、无缩进、大纲级别 1

当前代码会把文本精确等于 `摘要` 的段落识别成它。

### 9. `abstract_body`
- 当前用于中文摘要正文段落
- 当前效果：宋体小四、两端对齐、首行缩进 2 字符、大纲级别正文文本
- 当前版本会额外校验摘要字数是否在 800~1000 字之间，以及摘要区块是否超过 1 页

当前代码会把 `摘要` 标题之后、`关键词` 行之前的正文段落识别成它。

### 10. `keywords_line`
- 当前用于关键词整行
- 当前效果：宋体小四、两端对齐、无首行缩进、大纲级别正文文本
- 当前版本会额外把行首 `关键词：` 直接加粗，其后关键词保持常规字重
- 当前版本会额外校验关键词数量是否为 3~5 个、是否使用全角逗号 `，` 分隔，以及末尾是否没有标点

当前代码会把以 `关键词` / `关键词：` / `关键词:` 开头的整行识别成它。

### 11. `english_abstract_title`
- 当前用于英文摘要标题段
- 当前效果先按中文摘要标题类比：标题居中、无缩进、大纲级别 1
- 当前版本先保守识别文本精确等于 `Abstract` / `ABSTRACT` / 其他大小写变体的段落

当前代码会把文本大小写忽略后等于 `abstract` 的段落识别成它。

### 12. `english_abstract_body`
- 当前用于英文摘要正文段落
- 当前效果先按中文摘要正文类比：小四、两端对齐、首行缩进 2 字符、大纲级别正文文本
- 当前版本只提供识别与样式入口，不做英文摘要字数、页数校验

当前代码会把英文摘要标题之后、英文关键词行之前的正文段落识别成它。

### 13. `english_keywords_line`
- 当前用于英文关键词整行
- 当前效果先按中文关键词行类比：小四、两端对齐、无首行缩进、大纲级别正文文本
- 当前版本会把行首 `Keywords:` / `Key words:` 直接局部加粗，其后英文关键词保持常规字重
- 当前版本只提供识别与样式入口，不做英文关键词数量、分隔符、末尾标点校验

当前代码会把以 `Keywords` / `Keywords:` / `Key words` / `Key words:` 开头的整行识别成它。

### 14. `references_title`
- 当前用于参考文献标题段
- 当前效果：黑体、小三、居中、无缩进、大纲级别 1

当前代码会把文本精确等于 `参考文献` 的段落识别成它。

### 15. `reference_entry`
- 当前用于参考文献条目段
- 当前效果：宋体五号、两端对齐、无首行缩进、大纲级别正文文本

当前代码会把 `参考文献` 标题之后的条目段落识别成它。

### 16. `contents_title`
- 当前用于目录标题段
- 当前效果：黑体、小三、居中、无缩进、大纲级别 1

当前代码会把文本精确等于 `目录` 的段落识别成它。

### 17. `contents_entry`
- 当前用于目录条目段
- 当前效果：宋体小四、左对齐、无首行缩进、大纲级别正文文本
- 当前版本只提供保守的目录条目样式入口，不承诺自动目录域、点线、页码对齐和层级缩进优化

当前代码会把目录标题之后、且文本形态明显像目录项的段落识别成它。

### 18. `acknowledgements_title`
- 当前用于致谢标题段
- 当前效果：黑体、小三、居中、无缩进、大纲级别 1

当前代码会把文本精确等于 `致谢` 的段落识别成它。

### 19. `acknowledgements_body`
- 当前用于致谢正文段落
- 当前效果：宋体小四、两端对齐、首行缩进 2 字符、大纲级别正文文本

当前代码会把 `致谢` 标题之后、直到下一个结构标题或普通标题之前的正文段落识别成它。

### 20. `appendix_title`
- 当前用于附录标题段
- 当前效果：黑体、小三、居中、无缩进、大纲级别 1

当前代码会把 `附录`、`附录A`、`附录 A`、`附录一` 这类标题识别成它。

### 21. `appendix_body`
- 当前用于附录正文段落
- 当前效果：宋体小四、两端对齐、首行缩进 2 字符、大纲级别正文文本
- 第一版采取保守策略：遇到新的结构标题或普通标题就结束附录区块，不在本轮处理附录内部专用标题体系

当前代码会把附录标题之后、直到下一个结构标题或普通标题之前的正文段落识别成它。

### 22. `thesis_header`
- 当前用于非首页和首页页眉的样式入口
- 当前效果：宋体五号、居中、无首行缩进
- 当前版本只负责页眉文本所在段落的样式，不负责章节联动和页码

### 23. `thesis_footer`
- 当前用于非首页和首页页脚的样式入口
- 当前效果：宋体五号、居中、无首行缩进
- 当前版本只负责页脚文本所在段落的样式，不负责页码

---

## 当前项目里，这个配置文件实际控制什么

当前 `style_config.json` 主要控制七类东西：

### 1. 控制 Word 样式库本身
比如：
- 标题 1 用什么字体
- 正文字号是多少
- 题注是否居中

### 2. 控制图表题注样式
比如：
- 图注 / 表注用什么字号
- 是否居中
- 是否缩进

### 3. 控制图片段落样式规则
比如：
- 图片段落是否居中
- 图片段落是否缩进
- 图片段落段前是否空一行

### 4. 控制论文结构区块样式
比如：
- 中文摘要标题和正文长什么样
- 关键词整行用什么格式
- 参考文献标题和条目用什么格式

### 5. 控制主流程是否继续做段落匹配
比如：
- 只更新 Word 样式定义，不继续按规则给段落套样式
- 在样式定义更新后继续执行摘要、关键词、标题、目录等结构识别

### 6. 控制页面设置
比如：
- 上下左右页边距是多少
- 页眉距离和页脚距离是多少

### 7. 控制页眉页脚入口
比如：
- 是否启用首页不同
- 非首页和首页分别使用哪个样式
- 某个页眉或页脚当前是否写入固定文本

---

## 当前代码已支持的识别规则

| 段落特征 | 会被识别成 |
|---|---|
| 段落中包含行内图片 | `figure_block` |
| `表3-3 标题` | `caption` |
| `图2-1 标题` | `caption` |
| 文本精确等于 `摘要` | `abstract_title` |
| 大小写忽略后精确等于 `abstract` | `english_abstract_title` |
| 以 `关键词` / `关键词：` / `关键词:` 开头 | `keywords_line` |
| 以 `Keywords` / `Keywords:` / `Key words` / `Key words:` 开头 | `english_keywords_line` |
| 文本精确等于 `参考文献` | `references_title` |
| 文本精确等于 `目录` | `contents_title` |
| 文本精确等于 `致谢` | `acknowledgements_title` |
| `附录` / `附录A` / `附录 A` / `附录一` | `appendix_title` |
| `摘要` 标题之后且位于 `关键词` 前的正文段落 | `abstract_body` |
| `Abstract` 标题之后且位于英文关键词前的正文段落 | `english_abstract_body` |
| `参考文献` 标题之后的条目段落 | `reference_entry` |
| 目录标题之后且明显像目录项的段落 | `contents_entry` |
| `致谢` 标题之后的正文段落 | `acknowledgements_body` |
| `附录` 标题之后的正文段落 | `appendix_body` |
| `1.` | `heading_1` |
| `一、` | `heading_1` |
| `1.1` | `heading_2` |
| `1.1.1` | `heading_3` |
| `1.1.1.1` | `heading_4` |
| 其他非空段落 | `normal` |

补充说明：
- 当识别到 `keywords_line` 时，程序会同时检查中文关键词数量、分隔符和末尾标点。
- 当中文摘要区块结束时，程序会同时检查中文摘要字数和页数。
- 英文摘要 / 英文关键词当前版本只做识别和样式入口，不做英文内容校验。
- 当前版本对 `目录` 只做保守识别和样式入口，不承诺自动目录域、点线、页码对齐和层级缩进优化。

### 第一版图表题注识别规则

#### 图片段落
- 如果段落中包含行内图片，则识别为 `figure_block`
- 当前版本主要适配“行内图片”场景

#### 表注
- 文本格式需匹配：`表x-x 标题`
- 例如：
  - `表3-3 COT-CMP在ADMET基准组上的对比实验结果`
  - `表3-4 CoT-CMP在ADMET基准组上部分分类任务消融实验结果`
- 并且该段落的**紧随其后的下一个段落位于表格中**

#### 图注
- 文本格式需匹配：`图x-x 标题`
- 例如：
  - `图2-1 阿司匹林的多模态化学表征`
  - `图2-2 消息传递机制示意图`
- 第一版优先检查该段落的**上一个非空段落是否包含行内图片**
- 如果图片检测失败，但文本高度符合图注格式，当前版本仍会按图注处理

注意：
- 图片段落识别优先于图注 / 表注 / 标题识别
- 图注 / 表注识别优先于标题识别
- 标题识别仍然按 **四级 → 三级 → 二级 → 一级** 的顺序匹配

---

## 修改配置时的注意事项

1. `style_config.json` 必须保持合法 JSON
2. 不能在里面直接写注释
3. `processing.apply_paragraph_styles: false` 时，程序只会更新样式定义和页面设置，不会继续执行段落匹配和摘要/关键词内容校验
4. `page_setup.margins_cm` 中未显式写出的边距字段会回退到默认值
5. `header_footer.header`、`header_footer.footer`、`header_footer.first_page.header`、`header_footer.first_page.footer` 都必须引用已在 `styles` 中定义的 `style_id`
6. 当前页眉页脚留空时，程序会清空对应区域内容，不会顺带生成页码
7. `size` 现在必须写数字 pt，不能直接写 `小四`
8. `alignment` 目前只支持：
   - `left`
   - `center`
   - `justify`
9. `line_spacing_rule` 目前只支持：
   - `1.5_lines`
10. `color` 和 `color_index` 目前只支持：
   - `black`

---

## 一个最常见的修改例子

### 如果你想把题注设置成宋体五号居中
当前 `caption` 已经是：

```json
{
  "style_id": "caption",
  "builtin_names": {
    "english": "Caption",
    "chinese": "题注"
  },
  "font": {
    "name_ascii": "Times New Roman",
    "name_far_east": "宋体",
    "size": 10.5,
    "bold": false,
    "color": "black",
    "color_index": "black"
  },
  "paragraph": {
    "alignment": "center",
    "outline_level": 10,
    "first_line_indent_chars": 0,
    "line_spacing_rule": "1.5_lines",
    "space_before": 0,
    "space_after": 0
  }
}
```

### 如果你想让图片所在段落居中并段前一行
当前 `figure_block` 已经是：

```json
{
  "style_id": "figure_block",
  "builtin_names": {
    "english": "Body Figure",
    "chinese": "正文图片"
  },
  "font": {
    "name_ascii": "Times New Roman",
    "name_far_east": "宋体",
    "size": 12,
    "bold": false,
    "color": "black",
    "color_index": "black"
  },
  "paragraph": {
    "alignment": "center",
    "outline_level": 10,
    "first_line_indent_chars": 0,
    "line_spacing_rule": "1.5_lines",
    "space_before": 12,
    "space_after": 0
  }
}
```

其中：
- `alignment: center` 表示图片段落居中
- `first_line_indent_chars: 0` 表示无缩进
- `space_before` 最终会按 `normal.font.size` 动态覆盖

---

## 你最常改的几个地方

如果以后要调论文格式，通常最常改的是：

- `font.name_far_east`
- `font.size`
- `font.bold`
- `paragraph.alignment`
- `paragraph.first_line_indent_chars`
- `paragraph.space_before`
- `paragraph.space_after`

---

## 总结

你可以把 `style_config.json` 理解成：

- 它定义了“标题 1 / 标题 2 / 标题 3 / 标题 4 / 正文 / 图表题注 / 图片段落 / 摘要 / 关键词 / 参考文献 / 目录 / 致谢 / 附录 / 页眉 / 页脚”分别长什么样
- 它还能定义页面级参数，比如页边距、页眉页脚距离和首页不同
- 代码负责识别段落属于哪一类，并把这份配置里定义好的样式和页面设置应用到 Word 文档中
