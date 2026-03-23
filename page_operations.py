"""处理 Word 文档的页面设置与页眉页脚。"""

from style_operations import apply_direct_font_format, apply_direct_paragraph_format
from word_constants import WD_HEADER_FOOTER_FIRST_PAGE, WD_HEADER_FOOTER_PRIMARY



def centimeters_to_points(value_cm):
    """将厘米值转换为 Word 所需的磅值。"""
    return value_cm * 72 / 2.54



def apply_page_setup(doc, page_setup_config):
    """按配置对全文各节应用统一页面设置。"""
    if not page_setup_config["enabled"]:
        return

    margins_cm = page_setup_config["margins_cm"]
    top_margin = centimeters_to_points(margins_cm["top"])
    bottom_margin = centimeters_to_points(margins_cm["bottom"])
    left_margin = centimeters_to_points(margins_cm["left"])
    right_margin = centimeters_to_points(margins_cm["right"])
    header_distance = centimeters_to_points(page_setup_config["header_distance_cm"])
    footer_distance = centimeters_to_points(page_setup_config["footer_distance_cm"])

    for index in range(1, doc.Sections.Count + 1):
        section = doc.Sections(index)
        page_setup = section.PageSetup
        page_setup.TopMargin = top_margin
        page_setup.BottomMargin = bottom_margin
        page_setup.LeftMargin = left_margin
        page_setup.RightMargin = right_margin
        page_setup.HeaderDistance = header_distance
        page_setup.FooterDistance = footer_distance



def clear_header_footer(header_footer):
    """清空单个页眉或页脚内容。"""
    header_footer.LinkToPrevious = False
    header_footer.Range.Text = ""



def apply_header_footer_block(header_footer, block_config, style_lookup, style_config_lookup):
    """应用单个页眉或页脚配置块。"""
    if not block_config["enabled"] or not block_config["text"]:
        clear_header_footer(header_footer)
        return

    header_footer.LinkToPrevious = False
    header_footer.Range.Text = block_config["text"]
    style = style_lookup[block_config["style_ref"]]
    style_config = style_config_lookup[block_config["style_ref"]]

    for index in range(1, header_footer.Range.Paragraphs.Count + 1):
        paragraph = header_footer.Range.Paragraphs(index)
        paragraph.Range.Style = style
        apply_direct_font_format(paragraph, style_config)
        apply_direct_paragraph_format(paragraph, style_config)



def apply_header_footer(doc, header_footer_config, style_lookup, style_config_lookup):
    """按配置对全文各节应用页眉页脚。"""
    if not header_footer_config["enabled"]:
        return

    for index in range(1, doc.Sections.Count + 1):
        section = doc.Sections(index)
        section.PageSetup.DifferentFirstPageHeaderFooter = header_footer_config[
            "different_first_page"
        ]

        apply_header_footer_block(
            section.Headers(WD_HEADER_FOOTER_PRIMARY),
            header_footer_config["header"],
            style_lookup,
            style_config_lookup,
        )
        apply_header_footer_block(
            section.Footers(WD_HEADER_FOOTER_PRIMARY),
            header_footer_config["footer"],
            style_lookup,
            style_config_lookup,
        )

        if header_footer_config["different_first_page"]:
            first_page = header_footer_config["first_page"]
            apply_header_footer_block(
                section.Headers(WD_HEADER_FOOTER_FIRST_PAGE),
                first_page["header"],
                style_lookup,
                style_config_lookup,
            )
            apply_header_footer_block(
                section.Footers(WD_HEADER_FOOTER_FIRST_PAGE),
                first_page["footer"],
                style_lookup,
                style_config_lookup,
            )
