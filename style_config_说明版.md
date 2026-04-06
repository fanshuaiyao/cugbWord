# style 模板配置说明版

## 文件作用

当前项目已经拆成两层配置：

1. `runtime_config.json`
   - 控制本次处理哪个文档
   - 控制是否继续执行段落匹配
   - 控制当前选择哪个样式模板

2. `style/*.json`
   - 控制某个学校或模板本身的论文格式
   - 当前默认模板是 `style/cugb.json`

本说明文档描述的是 **样式模板文件** 的结构，也就是 `style/*.json` 的写法。

程序会按下面的流程使用它：

1. 读取 `runtime_config.json`
2. 读取 `style_template`
3. 定位到对应的 `style/*.json`
4. 读取 `page_setup`、`header_footer` 和 `page_numbering`
5. 打开对应的 Word 文档
6. 读取 `styles` 里的样式定义
7. 先更新 Word 内置样式和自定义样式本身
8. 再应用页面设置、页眉页脚和页码
9. 最后根据运行配置中的开关决定是否继续做段落匹配

注意：
- 模板文件必须是**严格 JSON**
- 不能直接写 `//` 或 `#` 注释
- 所以说明信息单独放在本文件里

---

## 顶层结构

```json
{
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
  "page_numbering": {
    "enabled": true,
    "sections": [
      {
        "section_index": 1,
        "enabled": true,
        "show_in_footer": true,
        "restart_at": null,
        "number_style": "upper_roman",
        "different_first_page": true
      },
      {
        "section_index": 2,
        "enabled": true,
        "show_in_footer": true,
        "restart_at": 1,
        "number_style": "arabic",
        "different_first_page": false
      }
    ]
  },
  "styles": [
    { ... }
  ]
}
```

### 顶层字段说明

| 字段 | 含义 | 当前示例 |
|---|---|---|
| `page_setup` | 页面设置配置。当前用于统一页边距和页眉页脚距离。 | `{ "enabled": true, ... }` |
| `header_footer` | 页眉页脚配置。当前用于首页不同与页眉页脚样式入口。 | `{ "enabled": true, ... }` |
| `page_numbering` | 页码配置。当前用于按节设置罗马数字 / 阿拉伯数字、是否重新编号、某节是否首页不同。 | `{ "enabled": true, ... }` |
| `styles` | 样式配置列表。每一项定义一个 Word 内置样式或自定义段落样式。 | 24 个样式项 |

### 运行配置补充说明

运行配置文件 `runtime_config.json` 当前负责：

| 字段 | 含义 | 默认值 |
|---|---|---|
| `document_path` | 目标 Word 文档路径。可以是相对路径，也可以是绝对路径。相对路径会基于项目目录解析。 | 无默认值，必须填写 |
| `style_template` | 要使用的模板标识或模板路径。写 `cugb` 时会自动解析到 `style/cugb.json`。 | `cugb` |
| `processing.apply_paragraph_styles` | 是否在样式定义更新后继续执行段落匹配与内容校验。`false` 时只更新样式定义与页面设置并保存处理后副本。 | `true` |

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

### page_numbering 字段说明

| 字段 | 含义 | 默认值 |
|---|---|---|
| `page_numbering.enabled` | 是否启用页码处理。 | `false` |
| `page_numbering.sections` | 按节配置页码规则的列表。 | `[]` |
| `page_numbering.sections[].section_index` | 目标节序号，从 1 开始。 | 必填 |
| `page_numbering.sections[].enabled` | 是否启用该节页码配置。 | `true` |
| `page_numbering.sections[].show_in_footer` | 是否在该节普通页脚插入页码。 | `true` |
| `page_numbering.sections[].restart_at` | 该节是否重新编号；`null` 表示延续，数字表示从该值开始。 | `null` |
| `page_numbering.sections[].number_style` | 页码样式。当前支持 `arabic`、`lower_roman`、`upper_roman`。 | `arabic` |
| `page_numbering.sections[].different_first_page` | 是否为该节单独设置首页不同；适合封面所在节。 | `null` |

说明：
- `section_index: 1` 通常对应封面、摘要、目录等前置部分。
- `section_index: 2` 通常对应目录后的正文节。
- 如果第一节第一页是封面，建议把第一节配置为 `different_first_page: true`，这样封面不显示页码，后续前置页再按罗马数字编号。

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
| `table_body` | 表格正文 |
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
- `table_body` 当前使用独立的自定义段落样式 `表格正文`，由代码在 Word 中获取或创建。
- `abstract_title`、`abstract_body`、`keywords_line`、`english_abstract_title`、`english_abstract_body`、`english_keywords_line`、`references_title`、`reference_entry`、`contents_title`、`contents_entry`、`acknowledgements_title`、`acknowledgements_body`、`appendix_title`、`appendix_body` 也属于自定义段落样式，由代码在 Word 中获取或创建。
- `thesis_header` 当前复用 Word 内置样式 `页眉`，`thesis_footer` 当前复用 Word 内置样式 `页脚`。
- 这意味着大多数结构样式仍然是独立入口，而页眉页脚参数则通过模板配置直接覆盖 Word 自带页眉 / 页脚样式。

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

## 当前项目里，这个模板文件实际控制什么

当前 `style/cugb.json` 主要控制五类东西：

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

### 5. 控制页面设置与页眉页脚入口
比如：
- 上下左右页边距是多少
- 页眉距离和页脚距离是多少
- 是否启用首页不同
- 非首页和首页分别使用哪个样式
- 某个页眉或页脚当前是否写入固定文本
- 某一节页码是否使用罗马数字或阿拉伯数字
- 某一节是否从 1 重新编号

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

## 修改模板时的注意事项

1. `style/*.json` 必须保持合法 JSON
2. 不能在里面直接写注释
3. `runtime_config.json` 中 `processing.apply_paragraph_styles: false` 时，程序只会更新样式定义和页面设置，不会继续执行段落匹配和摘要/关键词内容校验
4. `page_setup.margins_cm` 中未显式写出的边距字段会回退到默认值
5. `header_footer.header`、`header_footer.footer`、`header_footer.first_page.header`、`header_footer.first_page.footer` 都必须引用已在 `styles` 中定义的 `style_id`
6. 当前页眉页脚留空时，程序会清空对应区域内容，不会顺带生成固定文本页脚
7. 页码是否生成由 `page_numbering` 独立控制，不再依赖 `header_footer.footer.text`
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
- `page_setup.margins_cm`
- `header_footer`

---

## 总结

你可以把 `style/cugb.json` 理解成：

- 它定义了“标题 1 / 标题 2 / 标题 3 / 标题 4 / 正文 / 图表题注 / 图片段落 / 摘要 / 关键词 / 参考文献 / 目录 / 致谢 / 附录 / 页眉 / 页脚”分别长什么样
- 它还能定义页面级参数，比如页边距、页眉页脚距离和首页不同
- `runtime_config.json` 负责指定处理哪个文档、当前选择哪个模板、以及是否继续做段落匹配
- 代码负责识别段落属于哪一类，并把模板里定义好的样式和页面设置应用到 Word 文档中
