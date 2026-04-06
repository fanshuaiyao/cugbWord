"""提供运行配置、样式模板的读取、校验和路径解析能力。"""

import json
import os

from word_constants import (
    ALIGNMENT_MAP,
    COLOR_INDEX_MAP,
    COLOR_MAP,
    LINE_SPACING_RULE_MAP,
    PAGE_NUMBER_STYLE_MAP,
)


DEFAULT_PROCESSING_CONFIG = {"apply_paragraph_styles": True, "toc": {"enabled": True, "update_mode": "full"}}
DEFAULT_RUNTIME_CONFIG = {"style_template": "cugb"}
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
DEFAULT_PAGE_NUMBERING_CONFIG = {
    "enabled": False,
    "sections": [],
}


def require_non_empty_string(value, field_name):
    """校验配置字段是否为非空字符串。"""
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{field_name} 必须是非空字符串")



def require_number(value, field_name):
    """校验配置字段是否为数字。"""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{field_name} 必须是数字")



def resolve_enum_value(field_name, value, value_map, style_id):
    """将配置中的枚举字符串解析为 Word 所需的常量值。"""
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



def validate_page_numbering_config(page_numbering):
    """校验页码配置块。"""
    if not isinstance(page_numbering, dict):
        raise ValueError("page_numbering 必须是对象")
    enabled = page_numbering.get("enabled")
    if enabled is not None and not isinstance(enabled, bool):
        raise ValueError("page_numbering.enabled 必须是布尔值")
    sections = page_numbering.get("sections")
    if sections is not None:
        if not isinstance(sections, list):
            raise ValueError("page_numbering.sections 必须是数组")
        for index, section_config in enumerate(sections):
            if not isinstance(section_config, dict):
                raise ValueError(f"page_numbering.sections[{index}] 必须是对象")
            section_index = section_config.get("section_index")
            if isinstance(section_index, bool) or not isinstance(section_index, int):
                raise ValueError(
                    f"page_numbering.sections[{index}].section_index 必须是整数"
                )
            if section_index <= 0:
                raise ValueError(
                    f"page_numbering.sections[{index}].section_index 必须大于 0"
                )
            item_enabled = section_config.get("enabled")
            if item_enabled is not None and not isinstance(item_enabled, bool):
                raise ValueError(
                    f"page_numbering.sections[{index}].enabled 必须是布尔值"
                )
            show_in_footer = section_config.get("show_in_footer")
            if show_in_footer is not None and not isinstance(show_in_footer, bool):
                raise ValueError(
                    f"page_numbering.sections[{index}].show_in_footer 必须是布尔值"
                )
            show_on_first_page = section_config.get("show_on_first_page")
            if show_on_first_page is not None and not isinstance(show_on_first_page, bool):
                raise ValueError(
                    f"page_numbering.sections[{index}].show_on_first_page 必须是布尔值"
                )
            number_style = section_config.get("number_style", "arabic")
            if not isinstance(number_style, str) or number_style not in PAGE_NUMBER_STYLE_MAP:
                allowed_values = ", ".join(PAGE_NUMBER_STYLE_MAP.keys())
                raise ValueError(
                    f"page_numbering.sections[{index}].number_style 配置无效: {number_style}，可选值为: {allowed_values}"
                )
            restart_at = section_config.get("restart_at")
            if restart_at is not None:
                if isinstance(restart_at, bool) or not isinstance(restart_at, int):
                    raise ValueError(
                        f"page_numbering.sections[{index}].restart_at 必须是整数或 null"
                    )
                if restart_at < 0:
                    raise ValueError(
                        f"page_numbering.sections[{index}].restart_at 必须大于等于 0"
                    )
            different_first_page = section_config.get("different_first_page")
            if different_first_page is not None and not isinstance(different_first_page, bool):
                raise ValueError(
                    f"page_numbering.sections[{index}].different_first_page 必须是布尔值"
                )



def validate_style_template(config):
    """校验样式模板结构和关键字段是否合法。"""
    if not isinstance(config, dict):
        raise ValueError("样式模板必须是 JSON 对象")

    styles = config.get("styles")
    if not isinstance(styles, list) or not styles:
        raise ValueError("styles 必须是非空列表")

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

    page_numbering = config.get("page_numbering")
    if page_numbering is not None:
        validate_page_numbering_config(page_numbering)



def validate_processing_config(processing):
    """校验流程配置是否合法。"""
    if processing is None:
        return
    if not isinstance(processing, dict):
        raise ValueError("processing 必须是对象")
    apply_paragraph_styles = processing.get("apply_paragraph_styles")
    if apply_paragraph_styles is not None and not isinstance(
        apply_paragraph_styles, bool
    ):
        raise ValueError("processing.apply_paragraph_styles 必须是布尔值")

    # 校验 toc 配置
    toc = processing.get("toc")
    if toc is not None:
        if not isinstance(toc, dict):
            raise ValueError("processing.toc 必须是对象")
        enabled = toc.get("enabled")
        if enabled is not None and not isinstance(enabled, bool):
            raise ValueError("processing.toc.enabled 必须是布尔值")
        update_mode = toc.get("update_mode")
        if update_mode is not None and update_mode not in ("full", "page_numbers_only"):
            raise ValueError("processing.toc.update_mode 必须是 'full' 或 'page_numbers_only'")



def validate_runtime_config(config):
    """校验运行配置结构和关键字段是否合法。"""
    if not isinstance(config, dict):
        raise ValueError("运行配置必须是 JSON 对象")

    document_path = config.get("document_path")
    require_non_empty_string(document_path, "document_path")

    style_template = config.get("style_template")
    if style_template is not None:
        require_non_empty_string(style_template, "style_template")

    validate_processing_config(config.get("processing"))



def normalize_processing_config(config):
    """为流程配置补齐默认值。"""
    processing = config.get("processing") or {}
    normalized_processing = {**DEFAULT_PROCESSING_CONFIG, **processing}
    config["processing"] = normalized_processing
    return config



def normalize_runtime_config(config):
    """为运行配置补齐默认值。"""
    config = normalize_processing_config(config)
    config["style_template"] = config.get("style_template") or DEFAULT_RUNTIME_CONFIG[
        "style_template"
    ]
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



def normalize_page_numbering_config(config):
    """为页码配置补齐默认值。"""
    page_numbering = config.get("page_numbering") or {}
    normalized_sections = []
    for section_config in page_numbering.get("sections") or []:
        normalized_sections.append(
            {
                "enabled": True,
                "show_in_footer": True,
                "show_on_first_page": False,
                "number_style": "arabic",
                "restart_at": None,
                "different_first_page": None,
                **section_config,
            }
        )
    config["page_numbering"] = {
        **DEFAULT_PAGE_NUMBERING_CONFIG,
        **page_numbering,
        "sections": normalized_sections,
    }
    return config



def load_json_config(config_path, config_label):
    """读取 JSON 配置文件。"""
    try:
        with open(config_path, "r", encoding="utf-8") as config_file:
            return json.load(config_file)
    except FileNotFoundError as exc:
        raise FileNotFoundError(f"{config_label}不存在: {config_path}") from exc
    except json.JSONDecodeError as exc:
        raise ValueError(f"{config_label}不是合法的 JSON: {config_path}") from exc



def load_runtime_config(config_path):
    """读取并校验运行配置文件。"""
    config = load_json_config(config_path, "运行配置文件")
    validate_runtime_config(config)
    return normalize_runtime_config(config)



def load_style_template(config_path):
    """读取并校验样式模板文件。"""
    config = load_json_config(config_path, "样式模板文件")
    validate_style_template(config)
    config = normalize_page_setup_config(config)
    config = normalize_header_footer_config(config)
    return normalize_page_numbering_config(config)



def resolve_style_template_path(base_dir, style_template):
    """将模板标识或路径解析为样式模板文件绝对路径。"""
    if style_template.lower().endswith(".json") or any(
        separator in style_template for separator in ("/", "\\")
    ):
        return resolve_path(base_dir, style_template)
    return resolve_path(base_dir, os.path.join("style", f"{style_template}.json"))



def merge_execution_config(runtime_config, template_config):
    """将运行配置与模板配置合并为主流程使用的统一配置。"""
    merged_config = dict(template_config)
    merged_config["document_path"] = runtime_config["document_path"]
    merged_config["processing"] = runtime_config["processing"]
    return merged_config



def load_execution_config(base_dir, runtime_config_path):
    """读取运行配置与样式模板，并合并为主流程使用的统一配置。"""
    runtime_config = load_runtime_config(runtime_config_path)
    template_path = resolve_style_template_path(base_dir, runtime_config["style_template"])
    template_config = load_style_template(template_path)
    return merge_execution_config(runtime_config, template_config)



def resolve_path(base_dir, target_path):
    """将相对路径解析为绝对路径。"""
    if os.path.isabs(target_path):
        return target_path
    return os.path.abspath(os.path.join(base_dir, target_path))
