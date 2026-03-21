import re

from paragraph_utils import (
    get_next_paragraph,
    get_previous_non_empty_paragraph,
)


TABLE_CAPTION_PATTERN = re.compile(r"^表\s*\d+(?:[-\.]\d+)+\s+\S.+")
FIGURE_CAPTION_PATTERN = re.compile(r"^图\s*\d+(?:[-\.]\d+)+\s+\S.+")
HEADING_STYLE_PATTERNS = (
    ("heading_4", re.compile(r"^\d+\.\d+\.\d+\.\d+(?!\.\d)")),
    ("heading_3", re.compile(r"^\d+\.\d+\.\d+(?!\.\d)")),
    ("heading_2", re.compile(r"^\d+\.\d+(?!\.\d)")),
    ("heading_1", re.compile(r"^(?:\d+\.(?!\d)|[一二三四五六七八九十百千万]+、)")),
)



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
    """判断指定段落是否为表注。"""
    if not is_table_caption_text(text):
        return False
    next_paragraph = get_next_paragraph(doc, index)
    return has_table_in_paragraph(next_paragraph)



def is_figure_caption_paragraph(doc, index, text):
    """判断指定段落是否为图注。"""
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
