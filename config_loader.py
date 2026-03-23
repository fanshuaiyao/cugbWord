"""提供样式配置文件的读取、校验和路径解析能力。"""

import json
import os

from word_constants import ALIGNMENT_MAP, COLOR_INDEX_MAP, COLOR_MAP, LINE_SPACING_RULE_MAP


DEFAULT_PROCESSING_CONFIG = {"apply_paragraph_styles": True}
DEFAULT_PAGE_SETUP_CONFIG = {
    "enabled": False,
    "margins_cm": {"top": 2.5, "bottom": 2.0, "left": 2.5, "right": 2.0},
    "header_distance_cm": 1.5,
    "footer_distance_cm": 1.5,
}
DEFAULT_HEADER_FOOTER_CONFIG = {
    "enabled": False,
    "different_first_page": True,
    "header": {
        "enabled": False,
        "text": "",
        "style_ref": "thesis_header",
    },
    "footer": {
        "enabled": False,
        "text": "",
        "style_ref": "thesis_footer",
    },
    "first_page": {
        "header": {
            "enabled": False,
            "text": "",
            "style_ref": "thesis_header",
        },
        "footer": {
            "enabled": False,
            "text": "",
            "style_ref": "thesis_footer",
        },
    },
}


def require_non_empty_string(value, field_name):
    """校验配置字段是否为非空字符串。

    Args:
        value: 待校验的字段值。
        field_name: 字段名称，用于生成错误提示。

    Raises:
        ValueError: 当字段值不是非空字符串时抛出。
    """
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field_name} 必须是非空字符串")



def require_number(value, field_name):
    """校验配置字段是否为数字。

    Args:
        value: 待校验的字段值。
        field_name: 字段名称，用于生成错误提示。

    Raises:
        ValueError: 当字段值不是数字时抛出。
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{field_name} 必须是数字")



def resolve_enum_value(field_name, value, value_map, style_id):
    """将配置中的枚举字符串解析为 Word 所需的常量值。

    Args:
        field_name: 当前解析的字段名称。
        value: 配置中的枚举字符串。
        value_map: 枚举字符串到常量值的映射表。
        style_id: 当前样式标识，用于生成错误提示。

    Returns:
        对应的 Word 常量值。

    Raises:
        ValueError: 当配置值不在映射表中时抛出。
    """
    if value not in value_map:
        allowed_values = ", ".join(value_map.keys())
        raise ValueError(
            f"样式 {style_id} 的 {field_name} 配置无效: {value}，可选值为: {allowed_values}"
        )
    return value_map[value]



def validate_header_footer_block(config_block, block_path, style_ids):
    """校验单个页眉或页脚配置块。"""
    if not isinstance(config_block, dict):
        raise ValueError(f"header_footer.{block_path} 必须是对象")
    enabled = config_block.get("enabled")
    if enabled is not None and not isinstance(enabled, bool):
        raise ValueError(f"header_footer.{block_path}.enabled 必须是布尔值")
    require_non_empty_string(
        config_block.get("style_ref"), f"header_footer.{block_path}.style_ref"
    )
    if config_block["style_ref"] not in style_ids:
        raise ValueError(
            f"header_footer.{block_path}.style_ref 未在 styles 中定义: {config_block['style_ref']}"
        )
    text = config_block.get("text", "")
    if not isinstance(text, str):
        raise ValueError(f"header_footer.{block_path}.text 必须是字符串")



def validate_style_config(config):
    """校验样式配置结构和关键字段是否合法。

    Args:
        config: 从 JSON 文件读取出的配置对象。

    Raises:
        ValueError: 当配置结构、字段类型或枚举值不合法时抛出。
    """
    if not isinstance(config, dict):
        raise ValueError("样式配置必须是 JSON 对象")

    document_path = config.get("document_path")
    require_non_empty_string(document_path, "document_path")

    styles = config.get("styles")
    if not isinstance(styles, list) or not styles:
        raise ValueError("styles 必须是非空列表")

    processing = config.get("processing")
    if processing is not None:
        if not isinstance(processing, dict):
            raise ValueError("processing 必须是对象")
        apply_paragraph_styles = processing.get("apply_paragraph_styles")
        if apply_paragraph_styles is not None and not isinstance(
            apply_paragraph_styles, bool
        ):
            raise ValueError("processing.apply_paragraph_styles 必须是布尔值")

    style_ids = set()
    for index, style_config in enumerate(styles):
        if not isinstance(style_config, dict):
            raise ValueError(f"styles[{index}] 必须是对象")

        style_id = style_config.get("style_id")
        require_non_empty_string(style_id, f"styles[{index}].style_id")
        if style_id in style_ids:
            raise ValueError(f"style_id 重复: {style_id}")
        style_ids.add(style_id)

        builtin_names = style_config.get("builtin_names")
        if not isinstance(builtin_names, dict):
            raise ValueError(f"样式 {style_id} 的 builtin_names 必须是对象")
        require_non_empty_string(
            builtin_names.get("english"), f"样式 {style_id} 的 builtin_names.english"
        )
        require_non_empty_string(
            builtin_names.get("chinese"), f"样式 {style_id} 的 builtin_names.chinese"
        )

        font = style_config.get("font")
        if not isinstance(font, dict):
            raise ValueError(f"样式 {style_id} 的 font 必须是对象")
        require_non_empty_string(font.get("name_ascii"), f"样式 {style_id} 的 font.name_ascii")
        require_non_empty_string(
            font.get("name_far_east"), f"样式 {style_id} 的 font.name_far_east"
        )
        require_number(font.get("size"), f"样式 {style_id} 的 font.size")
        if font["size"] <= 0:
            raise ValueError(f"样式 {style_id} 的 font.size 必须大于 0")
        if not isinstance(font.get("bold"), bool):
            raise ValueError(f"样式 {style_id} 的 font.bold 必须是布尔值")

        color = font.get("color", "black")
        color_index = font.get("color_index", "black")
        resolve_enum_value("font.color", color, COLOR_MAP, style_id)
        resolve_enum_value("font.color_index", color_index, COLOR_INDEX_MAP, style_id)

        paragraph = style_config.get("paragraph")
        if not isinstance(paragraph, dict):
            raise ValueError(f"样式 {style_id} 的 paragraph 必须是对象")
        alignment = paragraph.get("alignment")
        require_non_empty_string(alignment, f"样式 {style_id} 的 paragraph.alignment")
        resolve_enum_value("paragraph.alignment", alignment, ALIGNMENT_MAP, style_id)

        outline_level = paragraph.get("outline_level")
        if isinstance(outline_level, bool) or not isinstance(outline_level, int):
            raise ValueError(f"样式 {style_id} 的 paragraph.outline_level 必须是整数")

        first_line_indent_chars = paragraph.get("first_line_indent_chars")
        require_number(
            first_line_indent_chars,
            f"样式 {style_id} 的 paragraph.first_line_indent_chars",
        )

        line_spacing_rule = paragraph.get("line_spacing_rule", "1.5_lines")
        require_non_empty_string(
            line_spacing_rule, f"样式 {style_id} 的 paragraph.line_spacing_rule"
        )
        resolve_enum_value(
            "paragraph.line_spacing_rule",
            line_spacing_rule,
            LINE_SPACING_RULE_MAP,
            style_id,
        )

        for field_name in ("space_before", "space_after"):
            value = paragraph.get(field_name, 0)
            require_number(value, f"样式 {style_id} 的 paragraph.{field_name}")

    page_setup = config.get("page_setup")
    if page_setup is not None:
        if not isinstance(page_setup, dict):
            raise ValueError("page_setup 必须是对象")
        enabled = page_setup.get("enabled")
        if enabled is not None and not isinstance(enabled, bool):
            raise ValueError("page_setup.enabled 必须是布尔值")
        margins_cm = page_setup.get("margins_cm")
        if margins_cm is not None:
            if not isinstance(margins_cm, dict):
                raise ValueError("page_setup.margins_cm 必须是对象")
            merged_margins = {**DEFAULT_PAGE_SETUP_CONFIG["margins_cm"], **margins_cm}
            for field_name in ("top", "bottom", "left", "right"):
                value = merged_margins.get(field_name)
                require_number(value, f"page_setup.margins_cm.{field_name}")
                if value <= 0:
                    raise ValueError(
                        f"page_setup.margins_cm.{field_name} 必须大于 0"
                    )
        for field_name in ("header_distance_cm", "footer_distance_cm"):
            value = page_setup.get(field_name)
            if value is not None:
                require_number(value, f"page_setup.{field_name}")
                if value <= 0:
                    raise ValueError(f"page_setup.{field_name} 必须大于 0")

    header_footer = config.get("header_footer")
    if header_footer is not None:
        if not isinstance(header_footer, dict):
            raise ValueError("header_footer 必须是对象")
        enabled = header_footer.get("enabled")
        if enabled is not None and not isinstance(enabled, bool):
            raise ValueError("header_footer.enabled 必须是布尔值")
        different_first_page = header_footer.get("different_first_page")
        if different_first_page is not None and not isinstance(
            different_first_page, bool
        ):
            raise ValueError("header_footer.different_first_page 必须是布尔值")
        validate_header_footer_block(
            {**DEFAULT_HEADER_FOOTER_CONFIG["header"], **(header_footer.get("header") or {})},
            "header",
            style_ids,
        )
        validate_header_footer_block(
            {**DEFAULT_HEADER_FOOTER_CONFIG["footer"], **(header_footer.get("footer") or {})},
            "footer",
            style_ids,
        )
        first_page = header_footer.get("first_page")
        if first_page is not None:
            if not isinstance(first_page, dict):
                raise ValueError("header_footer.first_page 必须是对象")
            validate_header_footer_block(
                {
                    **DEFAULT_HEADER_FOOTER_CONFIG["first_page"]["header"],
                    **(first_page.get("header") or {}),
                },
                "first_page.header",
                style_ids,
            )
            validate_header_footer_block(
                {
                    **DEFAULT_HEADER_FOOTER_CONFIG["first_page"]["footer"],
                    **(first_page.get("footer") or {}),
                },
                "first_page.footer",
                style_ids,
            )



def normalize_processing_config(config):
    """为流程配置补齐默认值。"""
    processing = config.get("processing") or {}
    normalized_processing = {**DEFAULT_PROCESSING_CONFIG, **processing}
    config["processing"] = normalized_processing
    return config



def normalize_page_setup_config(config):
    """为页面设置配置补齐默认值。"""
    page_setup = config.get("page_setup") or {}
    normalized_margins = {
        **DEFAULT_PAGE_SETUP_CONFIG["margins_cm"],
        **(page_setup.get("margins_cm") or {}),
    }
    normalized_page_setup = {
        **DEFAULT_PAGE_SETUP_CONFIG,
        **page_setup,
        "margins_cm": normalized_margins,
    }
    config["page_setup"] = normalized_page_setup
    return config



def normalize_header_footer_config(config):
    """为页眉页脚配置补齐默认值。"""
    header_footer = config.get("header_footer") or {}
    first_page = header_footer.get("first_page") or {}
    normalized_header = {
        **DEFAULT_HEADER_FOOTER_CONFIG["header"],
        **(header_footer.get("header") or {}),
    }
    normalized_footer = {
        **DEFAULT_HEADER_FOOTER_CONFIG["footer"],
        **(header_footer.get("footer") or {}),
    }
    normalized_first_page_header = {
        **DEFAULT_HEADER_FOOTER_CONFIG["first_page"]["header"],
        **(first_page.get("header") or {}),
    }
    normalized_first_page_footer = {
        **DEFAULT_HEADER_FOOTER_CONFIG["first_page"]["footer"],
        **(first_page.get("footer") or {}),
    }
    normalized_header_footer = {
        **DEFAULT_HEADER_FOOTER_CONFIG,
        **header_footer,
        "header": normalized_header,
        "footer": normalized_footer,
        "first_page": {
            **DEFAULT_HEADER_FOOTER_CONFIG["first_page"],
            **first_page,
            "header": normalized_first_page_header,
            "footer": normalized_first_page_footer,
        },
    }
    config["header_footer"] = normalized_header_footer
    return config



def load_style_config(config_path):
    """读取并校验样式配置文件。

    Args:
        config_path: 样式配置 JSON 文件路径。

    Returns:
        校验通过后的配置对象。

    Raises:
        FileNotFoundError: 当配置文件不存在时抛出。
        ValueError: 当配置文件不是合法 JSON 或配置内容不合法时抛出。
    """
    try:
        with open(config_path, "r", encoding="utf-8") as config_file:
            config = json.load(config_file)
    except FileNotFoundError as exc:
        raise FileNotFoundError(f"样式配置文件不存在: {config_path}") from exc
    except json.JSONDecodeError as exc:
        raise ValueError(f"样式配置文件不是合法的 JSON: {config_path}") from exc

    validate_style_config(config)
    config = normalize_processing_config(config)
    config = normalize_page_setup_config(config)
    return normalize_header_footer_config(config)



def resolve_path(base_dir, target_path):
    """将相对路径解析为绝对路径。

    Args:
        base_dir: 作为基准的目录路径。
        target_path: 目标路径，可以是相对路径或绝对路径。

    Returns:
        解析后的绝对路径。
    """
    if os.path.isabs(target_path):
        return target_path
    return os.path.abspath(os.path.join(base_dir, target_path))
