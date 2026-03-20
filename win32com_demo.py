import os
import re

import pythoncom
from win32com.client import DispatchEx

from config_loader import load_style_config, resolve_enum_value, resolve_path
from word_constants import (
    ALIGNMENT_MAP,
    COLOR_INDEX_MAP,
    COLOR_MAP,
    LINE_SPACING_RULE_MAP,
    WD_STYLE_TYPE_PARAGRAPH,
)


DEFAULT_CONFIG_FILE = "style_config.json"
TABLE_CAPTION_PATTERN = re.compile(r"^表\s*\d+(?:[-\.]\d+)+\s+\S.+")
FIGURE_CAPTION_PATTERN = re.compile(r"^图\s*\d+(?:[-\.]\d+)+\s+\S.+")
HEADING_STYLE_PATTERNS = (
    ("heading_4", re.compile(r"^\d+\.\d+\.\d+\.\d+(?!\.\d)")),
    ("heading_3", re.compile(r"^\d+\.\d+\.\d+(?!\.\d)")),
    ("heading_2", re.compile(r"^\d+\.\d+(?!\.\d)")),
    ("heading_1", re.compile(r"^(?:\d+\.(?!\d)|[一二三四五六七八九十百千万]+、)")),
)


def apply_style_config(style, style_config):
    """将单个样式配置应用到 Word 样式对象。

    Args:
        style: Word 样式对象。
        style_config: 单个样式的配置对象。
    """
    style_id = style_config["style_id"]
    font_config = style_config["font"]
    paragraph_config = style_config["paragraph"]

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

    style.ParagraphFormat.Alignment = resolve_enum_value(
        "paragraph.alignment",
        paragraph_config["alignment"],
        ALIGNMENT_MAP,
        style_id,
    )
    style.ParagraphFormat.OutlineLevel = paragraph_config["outline_level"]
    style.ParagraphFormat.LeftIndent = 0
    style.ParagraphFormat.RightIndent = 0
    style.ParagraphFormat.FirstLineIndent = 0
    style.ParagraphFormat.CharacterUnitFirstLineIndent = paragraph_config[
        "first_line_indent_chars"
    ]
    style.ParagraphFormat.LineSpacingRule = resolve_enum_value(
        "paragraph.line_spacing_rule",
        paragraph_config.get("line_spacing_rule", "1.5_lines"),
        LINE_SPACING_RULE_MAP,
        style_id,
    )
    style.ParagraphFormat.SpaceBefore = paragraph_config.get("space_before", 0)
    style.ParagraphFormat.SpaceAfter = paragraph_config.get("space_after", 0)


def get_builtin_style(doc, english_name, chinese_name):
    """获取 Word 内置样式，兼容中英文界面名称。

    Args:
        doc: Word 文档对象。
        english_name: 内置样式英文名。
        chinese_name: 内置样式中文名。

    Returns:
        匹配到的 Word 样式对象。

    Raises:
        ValueError: 当中英文名称都无法匹配到内置样式时抛出。
    """
    try:
        # 获取到的样式对象必须是内置样式，否则后续设置属性会失败
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

        if style_id == "figure_block":
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
    paragraph_format.SpaceBefore = (
        space_before_override
        if space_before_override is not None
        else paragraph_config.get("space_before", 0)
    )
    paragraph_format.SpaceAfter = paragraph_config.get("space_after", 0)


def normalize_paragraph_text(text):
    """清理 Word 段落文本中的段落结束符和首尾空白。

    Args:
        text: 原始段落文本。

    Returns:
        清理后的段落文本。
    """
    return text.replace("\r", "").replace("\x07", "").strip()


def get_next_paragraph(doc, start_index):
    """获取指定位置之后的下一个段落。"""
    if start_index >= doc.Paragraphs.Count:
        return None
    return doc.Paragraphs(start_index + 1)


def get_next_non_empty_paragraph(doc, start_index):
    """获取指定位置之后的下一个非空段落。"""
    for index in range(start_index + 1, doc.Paragraphs.Count + 1):
        paragraph = doc.Paragraphs(index)
        text = normalize_paragraph_text(paragraph.Range.Text)
        if text:
            return paragraph
    return None


def get_previous_non_empty_paragraph(doc, start_index):
    """获取指定位置之前的上一个非空段落。"""
    for index in range(start_index - 1, 0, -1):
        paragraph = doc.Paragraphs(index)
        text = normalize_paragraph_text(paragraph.Range.Text)
        if text:
            return paragraph
    return None


def is_table_caption_text(text):
    """判断文本是否符合表注格式。"""
    return bool(TABLE_CAPTION_PATTERN.match(text))


def is_figure_caption_text(text):
    """判断文本是否符合图注格式。"""
    return bool(FIGURE_CAPTION_PATTERN.match(text))


def has_table_in_paragraph(paragraph):
    """判断段落范围内是否包含表格。"""
    return paragraph is not None and paragraph.Range.Tables.Count > 0


def has_inline_shape_in_paragraph(paragraph):
    """判断段落范围内是否包含行内图片。"""
    return paragraph is not None and paragraph.Range.InlineShapes.Count > 0


def is_figure_block_paragraph(paragraph):
    """判断指定段落是否为图片所在段落。"""
    return has_inline_shape_in_paragraph(paragraph)


def is_table_caption_paragraph(doc, index, text):
    """判断指定段落是否为表注。

    第一版规则：文本匹配“表x-x 标题”，且紧随其后的下一个段落位于表格中。
    这里不能跳过空段落，因为表格首个单元格为空时，其段落文本仍可能为空。
    """
    if not is_table_caption_text(text):
        return False
    next_paragraph = get_next_paragraph(doc, index)
    return has_table_in_paragraph(next_paragraph)


def is_figure_caption_paragraph(doc, index, text):
    """判断指定段落是否为图注。

    第一版规则：文本匹配“图x-x 标题”，且上一个非空段落中包含行内图片。
    若图片检测失败，但文本非常像图注，则仍按图注处理。
    """
    if not is_figure_caption_text(text):
        return False
    previous_paragraph = get_previous_non_empty_paragraph(doc, index)
    if has_inline_shape_in_paragraph(previous_paragraph):
        return True
    return True


def match_heading_style_id(text):
    """根据段落开头内容判断标题样式标识。"""
    for style_id, pattern in HEADING_STYLE_PATTERNS:
        if pattern.match(text):
            return style_id
    return None


def match_paragraph_style_id(doc, paragraph, index, text):
    """根据段落文本和上下文判断应该应用的样式标识。"""
    if is_figure_block_paragraph(paragraph):
        return "figure_block"
    if is_table_caption_paragraph(doc, index, text):
        return "caption"
    if is_figure_caption_paragraph(doc, index, text):
        return "caption"

    heading_style_id = match_heading_style_id(text)
    if heading_style_id is not None:
        return heading_style_id
    return "normal"


def apply_paragraph_style(paragraph, style_id, style_lookup, style_config_lookup):
    """将指定样式应用到段落。"""
    style = style_lookup.get(style_id)
    style_config = style_config_lookup.get(style_id)
    if style is None or style_config is None:
        raise ValueError(f"段落匹配到未配置的样式: {style_id}")

    paragraph.Range.Style = style.NameLocal

    if style_id == "figure_block":
        normal_config = style_config_lookup.get("normal")
        if normal_config is None:
            raise ValueError("应用图片段落样式时缺少 normal 配置")

        apply_direct_font_format(paragraph, normal_config)
        apply_direct_paragraph_format(
            paragraph,
            style_config,
            space_before_override=normal_config["font"]["size"],
        )
        return

    apply_direct_paragraph_format(paragraph, style_config)


def apply_paragraph_styles(doc, style_lookup, style_config_lookup):
    """遍历文档段落并按识别结果应用对应样式。"""
    processed_count = 0
    for index in range(1, doc.Paragraphs.Count + 1):
        paragraph = doc.Paragraphs(index)
        text = normalize_paragraph_text(paragraph.Range.Text)
        if not text:
            continue

        style_id = match_paragraph_style_id(doc, paragraph, index, text)
        apply_paragraph_style(paragraph, style_id, style_lookup, style_config_lookup)
        processed_count += 1
    return processed_count


def main():
    """读取配置文件并执行 Word 样式更新流程。

    该方法会加载 JSON 配置、打开目标 Word 文档、更新样式定义、
    按段落内容识别标题层级并套用对应样式，最后保存结果。
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, DEFAULT_CONFIG_FILE)
    config = load_style_config(config_path)

    output_path = resolve_path(script_dir, config["document_path"])
    if not os.path.exists(output_path):
        raise FileNotFoundError(f"目标文档不存在: {output_path}")

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(output_path)
        style_config_lookup = build_style_config_lookup(config["styles"])
        style_lookup = apply_styles(doc, config["styles"])
        processed_count = apply_paragraph_styles(doc, style_lookup, style_config_lookup)

        doc.Save()
        print(
            f"已更新 {len(config['styles'])} 个样式定义，并处理 {processed_count} 个非空段落: {output_path}"
        )
    finally:
        if doc is not None:
            doc.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
