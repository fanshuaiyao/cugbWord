from config_loader import resolve_enum_value
from word_constants import (
    ALIGNMENT_MAP,
    COLOR_INDEX_MAP,
    COLOR_MAP,
    LINE_SPACING_RULE_MAP,
    WD_STYLE_TYPE_PARAGRAPH,
)


HEADING_STYLE_IDS = {"heading_1", "heading_2", "heading_3", "heading_4"}


def apply_style_config(style, style_config):
    """将单个样式配置应用到 Word 样式对象。"""
    style_id = style_config["style_id"]
    font_config = style_config["font"]
    paragraph_config = style_config["paragraph"]
    paragraph_format = style.ParagraphFormat

    if style_id in HEADING_STYLE_IDS:
        style.BaseStyle = ""

    style.Font.Name = font_config["name_ascii"]
    style.Font.NameAscii = font_config["name_ascii"]
    style.Font.NameFarEast = font_config["name_far_east"]
    style.Font.Size = font_config["size"]
    style.Font.Bold = font_config["bold"]
    style.Font.Color = resolve_enum_value(
        "font.color", font_config.get("color", "black"), COLOR_MAP, style_id
    )
    style.Font.ColorIndex = resolve_enum_value(
        "font.color_index",
        font_config.get("color_index", "black"),
        COLOR_INDEX_MAP,
        style_id,
    )

    paragraph_format.Alignment = resolve_enum_value(
        "paragraph.alignment",
        paragraph_config["alignment"],
        ALIGNMENT_MAP,
        style_id,
    )
    paragraph_format.OutlineLevel = paragraph_config["outline_level"]
    paragraph_format.LeftIndent = 0
    paragraph_format.RightIndent = 0
    paragraph_format.FirstLineIndent = 0
    paragraph_format.CharacterUnitLeftIndent = 0
    paragraph_format.CharacterUnitRightIndent = 0
    paragraph_format.CharacterUnitFirstLineIndent = paragraph_config[
        "first_line_indent_chars"
    ]
    paragraph_format.LineSpacingRule = resolve_enum_value(
        "paragraph.line_spacing_rule",
        paragraph_config.get("line_spacing_rule", "1.5_lines"),
        LINE_SPACING_RULE_MAP,
        style_id,
    )
    paragraph_format.SpaceBeforeAuto = False
    paragraph_format.SpaceAfterAuto = False
    paragraph_format.LineUnitBefore = 0
    paragraph_format.LineUnitAfter = 0
    paragraph_format.SpaceBefore = paragraph_config.get("space_before", 0)
    paragraph_format.SpaceAfter = paragraph_config.get("space_after", 0)



def get_builtin_style(doc, english_name, chinese_name):
    """获取 Word 内置样式，兼容中英文界面名称。"""
    try:
        return doc.Styles(english_name)
    except Exception:
        try:
            return doc.Styles(chinese_name)
        except Exception as exc:
            raise ValueError(
                f"无法找到 Word 内置样式: {english_name} / {chinese_name}"
            ) from exc



def get_or_create_custom_style(doc, style_name):
    """获取自定义段落样式，若不存在则创建。"""
    try:
        return doc.Styles(style_name)
    except Exception:
        return doc.Styles.Add(Name=style_name, Type=WD_STYLE_TYPE_PARAGRAPH)



def apply_styles(doc, style_configs):
    """遍历配置并获取多个 Word 样式对象。"""
    style_lookup = {}

    for style_config in style_configs:
        style_id = style_config["style_id"]
        builtin_names = style_config["builtin_names"]

        if style_config.get("custom", False):
            style = get_or_create_custom_style(doc, builtin_names["chinese"])
            apply_style_config(style, style_config)
            style_lookup[style_id] = style
            continue

        style = get_builtin_style(
            doc,
            builtin_names["english"],
            builtin_names["chinese"],
        )
        apply_style_config(style, style_config)
        style_lookup[style_id] = style

    return style_lookup



def build_style_config_lookup(style_configs):
    """构建以 style_id 为键的样式配置映射。"""
    return {style_config["style_id"]: style_config for style_config in style_configs}



def apply_direct_font_format(paragraph, style_config):
    """按配置直接覆盖单个段落的字体格式，而不修改 Word 样式定义。"""
    style_id = style_config["style_id"]
    font_config = style_config["font"]

    paragraph.Range.Font.Name = font_config["name_ascii"]
    paragraph.Range.Font.NameAscii = font_config["name_ascii"]
    paragraph.Range.Font.NameFarEast = font_config["name_far_east"]
    paragraph.Range.Font.Size = font_config["size"]
    paragraph.Range.Font.Bold = font_config["bold"]
    paragraph.Range.Font.Color = resolve_enum_value(
        "font.color", font_config.get("color", "black"), COLOR_MAP, style_id
    )
    paragraph.Range.Font.ColorIndex = resolve_enum_value(
        "font.color_index",
        font_config.get("color_index", "black"),
        COLOR_INDEX_MAP,
        style_id,
    )



def apply_direct_paragraph_format(paragraph, style_config, space_before_override=None):
    """按配置直接覆盖单个段落的段落格式，而不修改 Word 样式定义。"""
    style_id = style_config["style_id"]
    paragraph_config = style_config["paragraph"]
    paragraph_format = paragraph.Range.ParagraphFormat

    paragraph_format.LeftIndent = 0
    paragraph_format.RightIndent = 0
    paragraph_format.FirstLineIndent = 0
    paragraph_format.CharacterUnitLeftIndent = 0
    paragraph_format.CharacterUnitRightIndent = 0
    paragraph_format.CharacterUnitFirstLineIndent = 0
    paragraph_format.Alignment = resolve_enum_value(
        "paragraph.alignment",
        paragraph_config["alignment"],
        ALIGNMENT_MAP,
        style_id,
    )
    paragraph_format.OutlineLevel = paragraph_config["outline_level"]
    paragraph_format.CharacterUnitFirstLineIndent = paragraph_config[
        "first_line_indent_chars"
    ]
    paragraph_format.LineSpacingRule = resolve_enum_value(
        "paragraph.line_spacing_rule",
        paragraph_config.get("line_spacing_rule", "1.5_lines"),
        LINE_SPACING_RULE_MAP,
        style_id,
    )
    paragraph_format.SpaceBeforeAuto = False
    paragraph_format.SpaceAfterAuto = False
    paragraph_format.LineUnitBefore = 0
    paragraph_format.LineUnitAfter = 0
    paragraph_format.SpaceBefore = (
        space_before_override
        if space_before_override is not None
        else paragraph_config.get("space_before", 0)
    )
    paragraph_format.SpaceAfter = paragraph_config.get("space_after", 0)
