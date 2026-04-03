import os
import sys

import pythoncom
from win32com.client import DispatchEx

from config_loader import load_execution_config, resolve_path
from page_operations import apply_header_footer, apply_page_setup
from paragraph_processing import apply_paragraph_styles
from style_operations import apply_styles, build_style_config_lookup
from toc_operations import process_toc


DEFAULT_CONFIG_FILE = "runtime_config.json"
OUTPUT_SUFFIX = "_处理后"


def build_output_path(input_path):
    """根据输入文档路径生成固定的处理后副本路径。"""
    base_path, extension = os.path.splitext(input_path)
    return f"{base_path}{OUTPUT_SUFFIX}{extension}"


def show_progress(current, total):
    """在控制台输出简单进度条。"""
    bar_width = 30
    progress_ratio = current / total if total else 1
    filled_width = int(bar_width * progress_ratio)
    bar = "#" * filled_width + "-" * (bar_width - filled_width)
    sys.stdout.write(f"\r处理段落进度: [{bar}] {current}/{total}")
    sys.stdout.flush()
    if current >= total:
        sys.stdout.write("\n")



def main():
    """读取配置文件并执行 Word 样式与页面设置流程。

    该方法会加载 JSON 配置、打开目标 Word 文档、更新样式定义、
    应用页面设置和页眉页脚，再按段落内容识别标题层级并套用对应样式，最后保存结果。
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, DEFAULT_CONFIG_FILE)
    print("[1/7] 正在读取运行配置与样式模板...", flush=True)
    config = load_execution_config(script_dir, config_path)

    input_path = resolve_path(script_dir, config["document_path"])
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"目标文档不存在: {input_path}")

    output_path = build_output_path(input_path)

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        print("[2/7] 正在启动 Word...", flush=True)
        word = DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        print(f"[3/7] 正在打开文档: {input_path}", flush=True)
        doc = word.Documents.Open(input_path)
        print(f"[4/7] 正在更新 {len(config['styles'])} 个样式定义...", flush=True)
        style_config_lookup = build_style_config_lookup(config["styles"])
        style_lookup = apply_styles(doc, config["styles"])

        print("[5/7] 正在应用页面设置...", flush=True)
        apply_page_setup(doc, config["page_setup"])
        apply_header_footer(doc, config["header_footer"], style_lookup, style_config_lookup)

        processed_count = 0
        validation_issues = []
        validation_counts = {"abstract_count": 0, "keywords_count": 0}
        if config["processing"]["apply_paragraph_styles"]:
            print(f"[6/7] 正在处理 {doc.Paragraphs.Count} 个段落...", flush=True)
            processed_count, validation_issues, validation_counts = apply_paragraph_styles(
                doc, style_lookup, style_config_lookup, progress_callback=show_progress
            )
        else:
            print("[6/7] 已按配置跳过段落匹配与内容校验...", flush=True)

        # 目录处理：必须在段落样式应用完成后执行
        print("[6.5/7] 正在处理目录...", flush=True)
        toc_config = config.get("processing", {}).get("toc", {})
        if toc_config.get("enabled", True):
            success, action, message = process_toc(doc, toc_config)
            if success:
                print(f"  ✓ {message}")
            else:
                print(f"  ⚠ {message}")
        else:
            print("  - 已按配置跳过目录处理")

        print(f"[7/7] 正在保存处理后文档: {output_path}", flush=True)
        doc.SaveAs2(output_path)
        if config["processing"]["apply_paragraph_styles"]:
            print(
                f"已更新 {len(config['styles'])} 个样式定义，并处理 {processed_count} 个非空段落: {output_path}"
            )
            if validation_counts["abstract_count"] == 0 and validation_counts["keywords_count"] == 0:
                print("未识别到摘要或关键词区块，未执行内容校验。")
            elif validation_issues:
                print("检测到以下内容校验问题：")
                for issue in validation_issues:
                    print(f"- {issue}")
            else:
                print("摘要与关键词内容校验通过。")
        else:
            print(f"已更新 {len(config['styles'])} 个样式定义，并跳过段落匹配: {output_path}")
    finally:
        if doc is not None:
            doc.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
