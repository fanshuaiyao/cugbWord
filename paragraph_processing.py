from paragraph_rules import match_paragraph_style_id
from paragraph_utils import normalize_paragraph_text
from style_operations import apply_direct_font_format, apply_direct_paragraph_format



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
