"""
清除样式测试 Demo
用于测试 ParagraphFormat.Reset() 和 Font.Reset() 对段落和字符级别直接格式的影响。
Word 保持可见，方便持续调试观察。
"""

import os
import pythoncom
from win32com.client import DispatchEx


def clear_paragraph_direct_formatting(paragraph):
    """清除段落的直接格式（段落级 + 字符级）。"""
    paragraph.Range.ParagraphFormat.Reset()
    paragraph.Range.Font.Reset()


def clear_document_direct_formatting(doc):
    """清除文档中所有段落的直接格式。"""
    print("正在清除所有段落的直接格式...")
    count = 0
    for para in doc.Paragraphs:
        clear_paragraph_direct_formatting(para)
        count += 1
    print(f"已清除 {count} 个段落的直接格式")


def apply_style_to_paragraph(paragraph, style_name, clear_first=True):
    """给段落应用样式，可选择先清除直接格式。"""
    if clear_first:
        clear_paragraph_direct_formatting(paragraph)
    paragraph.Range.Style = style_name


def main():
    """打开文档，清除直接格式并应用样式。Word 保持可见方便调试。"""
    # 配置：修改为你的测试文档路径
    test_doc_path = os.path.abspath("test/test_clear_format.docx")

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        print("启动 Word（可见模式）...")
        word = DispatchEx("Word.Application")
        word.Visible = True  # 保持可见，方便调试
        word.DisplayAlerts = 0

        # 如果没有测试文档，先创建一个带混乱格式的测试文档
        if not os.path.exists(test_doc_path):
            print(f"测试文档不存在，创建测试文档: {test_doc_path}")
            os.makedirs(os.path.dirname(test_doc_path), exist_ok=True)
            doc = word.Documents.Add()

            # 插入一段带混乱格式的文本
            para1 = doc.Paragraphs.Add()
            para1.Range.Text = "1.1 这是标题（带手动格式）\n"
            para1.Range.Font.Name = "微软雅黑"  # 手动设置字体
            para1.Range.Font.Size = 18  # 手动设置字号
            para1.Range.Font.Bold = True  # 手动加粗
            para1.Range.ParagraphFormat.Alignment = 0  # 左对齐

            para2 = doc.Paragraphs.Add()
            para2.Range.Text = "这是正文段落（部分文字有手动格式）。\n"
            # 让部分文字有不同格式
            para2.Range.Characters(1).Font.Name = "Arial"
            para2.Range.Characters(1).Font.Size = 20
            para2.Range.Characters(1).Font.Color = 255  # 红色

            doc.SaveAs2(test_doc_path)
            print(f"测试文档已创建: {test_doc_path}")
            print("请观察文档，手动格式（字体、颜色、字号）已经应用")
            print("按 Enter 键继续清除格式...")
            input()
        else:
            print(f"打开现有测试文档: {test_doc_path}")
            doc = word.Documents.Open(test_doc_path)

        # 测试1：清除所有段落格式
        print("\n=== 测试：清除所有段落和字符的直接格式 ===")
        clear_document_direct_formatting(doc)
        print("已清除直接格式，请观察 Word 文档变化")
        print("按 Enter 键继续应用样式...")
        input()

        # 测试2：应用样式
        print("\n=== 测试：应用 Heading 1 样式到第一段 ===")
        first_para = doc.Paragraphs(1)
        apply_style_to_paragraph(first_para, "Heading 1", clear_first=True)
        print("已应用 Heading 1 样式")

        print("\n=== 测试：应用 Normal 样式到第二段 ===")
        if doc.Paragraphs.Count >= 2:
            second_para = doc.Paragraphs(2)
            apply_style_to_paragraph(second_para, "Normal", clear_first=True)
            print("已应用 Normal 样式")

        doc.Save()
        print(f"\n文档已保存: {test_doc_path}")
        print("Word 保持打开，你可以继续观察和调试")
        print("调试完成后手动关闭 Word，或按 Ctrl+C 结束脚本")

        # 保持脚本运行，Word 可见
        print("\n脚本等待中... 按 Enter 键结束并关闭 Word")
        input()

    except KeyboardInterrupt:
        print("\n收到中断信号")
    finally:
        if doc is not None:
            doc.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
        print("Word 已关闭")


if __name__ == "__main__":
    main()
