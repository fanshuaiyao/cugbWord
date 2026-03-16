import os

import pythoncom
from win32com.client import DispatchEx

from config_loader import load_style_config, resolve_enum_value, resolve_path
from word_constants import ALIGNMENT_MAP, COLOR_INDEX_MAP, COLOR_MAP, LINE_SPACING_RULE_MAP


DEFAULT_CONFIG_FILE = "style_config.json"


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
        return doc.Styles(english_name)
    except Exception:
        try:
            return doc.Styles(chinese_name)
        except Exception as exc:
            raise ValueError(
                f"无法找到 Word 内置样式: {english_name} / {chinese_name}"
            ) from exc


def apply_styles(doc, style_configs):
    """遍历配置并批量应用多个 Word 内置样式。

    Args:
        doc: Word 文档对象。
        style_configs: 样式配置列表。
    """
    for style_config in style_configs:
        builtin_names = style_config["builtin_names"]
        style = get_builtin_style(
            doc,
            builtin_names["english"],
            builtin_names["chinese"],
        )
        apply_style_config(style, style_config)


def main():
    """读取配置文件并执行 Word 样式更新流程。

    该方法会加载 JSON 配置、打开目标 Word 文档、应用样式并保存结果。
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
        apply_styles(doc, config["styles"])

        doc.Save()
        print(f"已应用 {len(config['styles'])} 个样式，文档已更新: {output_path}")
    finally:
        if doc is not None:
            doc.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
