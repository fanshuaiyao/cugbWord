"""集中定义 Word 自动化中使用的常量和枚举映射。"""

WD_ALIGN_PARAGRAPH_LEFT = 0
"""Word 段落左对齐常量值。"""

WD_ALIGN_PARAGRAPH_CENTER = 1
"""Word 段落居中对齐常量值。"""

WD_ALIGN_PARAGRAPH_JUSTIFY = 3
"""Word 段落两端对齐常量值。"""

WD_LINE_SPACE_1PT5 = 1
"""Word 1.5 倍行距常量值。"""

WD_COLOR_BLACK = 0
"""Word 黑色字体颜色常量值。"""

WD_COLOR_INDEX_BLACK = 1
"""Word 黑色字体颜色索引常量值。"""

WD_STYLE_TYPE_PARAGRAPH = 1
"""Word 段落样式类型常量值。"""

WD_STATISTIC_PAGES = 2
"""Word 统计页数常量值。"""

WD_HEADER_FOOTER_PRIMARY = 1
"""Word 普通页页眉页脚常量值。"""

WD_HEADER_FOOTER_FIRST_PAGE = 2
"""Word 首页页眉页脚常量值。"""

WD_SECTION_BREAK_NEXT_PAGE = 2
"""Word 下一页分节符常量值。"""

WD_SECTION_BREAK_CONTINUOUS = 3
"""Word 连续分节符常量值。"""

WD_BREAK_PAGE = 7
"""Word 分页符常量值。"""

ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH_LEFT,
    "center": WD_ALIGN_PARAGRAPH_CENTER,
    "justify": WD_ALIGN_PARAGRAPH_JUSTIFY,
}
"""段落对齐方式字符串到 Word 常量值的映射。"""

LINE_SPACING_RULE_MAP = {
    "1.5_lines": WD_LINE_SPACE_1PT5,
}
"""行距规则字符串到 Word 常量值的映射。"""

COLOR_MAP = {
    "black": WD_COLOR_BLACK,
}
"""字体颜色字符串到 Word 常量值的映射。"""

COLOR_INDEX_MAP = {
    "black": WD_COLOR_INDEX_BLACK,
}
"""字体颜色索引字符串到 Word 常量值的映射。"""
