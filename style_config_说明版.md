# style_config 说明版

## 文件作用

`style_config.json` 是当前项目的样式配置文件。

程序会按下面的流程使用它：

1. 读取 `document_path`
2. 打开对应的 Word 文档
3. 读取 `styles` 里的样式定义
4. 先更新 Word 内置样式本身
5. 再根据段落开头规则，把段落设置成对应样式

注意：
- `style_config.json` 必须是**严格 JSON**
- 不能直接写 `//` 或 `#` 注释
- 所以说明信息单独放在本文件里

---

## 顶层结构

```json
{
  "document_path": "test/123.docx",
  "styles": [
    { ... }
  ]
}
```

### 顶层字段说明

| 字段 | 含义 | 当前示例 |
|---|---|---|
| `document_path` | 目标 Word 文档路径。可以是相对路径，也可以是绝对路径。相对路径会基于项目目录解析。 | `test/123.docx` |
| `styles` | 样式配置列表。每一项定义一个 Word 内置样式。 | 7 个样式项 |

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
- `figure_block` 当前为了复用现有机制，暂时映射到 `Normal / 正文` 这个内置样式对象，再由代码额外覆盖段落对齐、缩进和段前距等效果。
- 这意味着它是“图片段落格式规则”的一个配置入口，而不是一个独立的 Word 内置样式名。

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

## 当前 7 个样式分别表示什么

### 1. `heading_1`
- 对应 Word 内置样式：`Heading 1` / `标题 1`
- 当前字体：`Times New Roman + 黑体`
- 当前字号：`15`
- 当前对齐：`center`
- 当前大纲级别：`1`

当前代码会把这些段落识别成它：
- `1. 绪论`
- `一、绪论`

### 2. `heading_2`
- 对应 Word 内置样式：`Heading 2` / `标题 2`
- 当前字体：`Times New Roman + 黑体`
- 当前字号：`14`
- 当前对齐：`left`
- 当前大纲级别：`2`

当前代码会把这些段落识别成它：
- `1.1 研究背景`

### 3. `heading_3`
- 对应 Word 内置样式：`Heading 3` / `标题 3`
- 当前字体：`Times New Roman + 黑体`
- 当前字号：`14`
- 当前对齐：`justify`
- 当前大纲级别：`3`

当前代码会把这些段落识别成它：
- `1.1.1 国内研究现状`

### 4. `heading_4`
- 对应 Word 内置样式：`Heading 4` / `标题 4`
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

---

## 当前项目里，这个配置文件实际控制什么

当前 `style_config.json` 主要控制三类东西：

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

---

## 当前代码已支持的识别规则

| 段落特征 | 会被识别成 |
|---|---|
| 段落中包含行内图片 | `figure_block` |
| `表3-3 标题` | `caption` |
| `图2-1 标题` | `caption` |
| `1.` | `heading_1` |
| `一、` | `heading_1` |
| `1.1` | `heading_2` |
| `1.1.1` | `heading_3` |
| `1.1.1.1` | `heading_4` |
| 其他非空段落 | `normal` |

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
3. `size` 现在必须写数字 pt，不能直接写 `小四`
4. `alignment` 目前只支持：
   - `left`
   - `center`
   - `justify`
5. `line_spacing_rule` 目前只支持：
   - `1.5_lines`
6. `color` 和 `color_index` 目前只支持：
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
    "english": "Normal",
    "chinese": "正文"
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

- 它定义了“标题 1 / 标题 2 / 标题 3 / 标题 4 / 正文 / 图表题注 / 图片段落”分别长什么样
- 代码负责识别段落属于哪一类
- 然后把这份配置里定义好的样式应用到 Word 文档中
