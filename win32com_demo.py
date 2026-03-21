import os

import pythoncom
from win32com.client import DispatchEx

from config_loader import load_style_config, resolve_path
from paragraph_processing import apply_paragraph_styles
from style_operations import apply_styles, build_style_config_lookup


DEFAULT_CONFIG_FILE = "style_config.json"


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
