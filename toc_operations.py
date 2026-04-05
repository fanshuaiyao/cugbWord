"""提供目录检测、插入和更新功能。

目录插入策略：始终在英文摘要（English Abstract）区块之后的下一页插入。
流程：分页 → 插入"目录"标题并套样式 → 在标题之后插入自动目录域。
"""

import re

from paragraph_utils import normalize_paragraph_text
from style_operations import apply_direct_font_format, apply_direct_paragraph_format


def detect_toc(doc):
    """检测文档中是否存在目录。

    Args:
        doc: Word 文档对象。

    Returns:
        bool: 是否存在目录。
    """
    return doc.TablesOfContents.Count > 0


def find_english_keywords_paragraph(doc):
    """找到英文关键词行所在的段落。

    从文档开头向下遍历，找到 "Key Words" / "Keywords" 行（英文摘要区块的最后一段）。

    Args:
        doc: Word 文档对象。

    Returns:
        Paragraph 对象或 None。
    """
    pattern = re.compile(r"^key\s*words(?:\s*[：:].*)?$", re.IGNORECASE)
    for para in doc.Paragraphs:
        text = normalize_paragraph_text(para.Range.Text)
        if text and pattern.match(text):
            return para
    return None


def insert_toc_title(doc, range_obj, style_lookup, style_config_lookup):
    """在指定位置插入"目录"标题文字并应用 contents_title 样式。

    Args:
        doc: Word 文档对象。
        range_obj: 插入位置的 Range 对象。
        style_lookup: 样式名到 Word 样式对象的映射。
        style_config_lookup: 样式名到配置字典的映射。

    Returns:
        Paragraph 对象: 新插入的"目录"标题段落。
    """
    range_obj.InsertAfter("目录\r")
    title_para = doc.Range(range_obj.Start, range_obj.End).Paragraphs(1)

    style = style_lookup.get("contents_title")
    style_config = style_config_lookup.get("contents_title")
    if style is not None and style_config is not None:
        title_para.Range.Style = style.NameLocal
        apply_direct_font_format(title_para, style_config)
        apply_direct_paragraph_format(title_para, style_config)

    return title_para


def find_toc_insertion_point(doc):
    """找到目录应该插入的位置：英文关键词行之后。

    Args:
        doc: Word 文档对象。

    Returns:
        Range 对象或 None: 用于插入目录标题的 Range。
    """
    keywords_para = find_english_keywords_paragraph(doc)
    if keywords_para is None:
        return None

    # 因为目录标题已经自带“段前分页”属性，我们直接在英文关键词之后插入即可，不需要再硬塞一个分页符
    end_range = keywords_para.Range.Duplicate
    end_range.Collapse(0)  # wdCollapseEnd
    return end_range


def insert_toc(doc, range_obj, upper_level=1, lower_level=4):
    """在指定位置插入自动目录域。

    Args:
        doc: Word 文档对象。
        range_obj: 插入位置的 Range 对象。
        upper_level: 最高标题级别，默认 1。
        lower_level: 最低标题级别，默认 4。

    Returns:
        TableOfContents 对象: 新插入的目录。
    """
    toc = doc.TablesOfContents.Add(
        Range=range_obj,
        UseHeadingStyles=True,
        UpperHeadingLevel=upper_level,
        LowerHeadingLevel=lower_level,
        IncludePageNumbers=True,
        RightAlignPageNumbers=True,
        UseHyperlinks=True
    )

    # 强制在底层域代码中追加 \h 开关以启用超链接
    doc.ActiveWindow.View.ShowFieldCodes = False # 确保不是处于域代码显示状态
    if toc.Range.Fields.Count > 0:
        field = toc.Range.Fields(1)
        code_text = field.Code.Text
        if r"\h" not in code_text:
            field.Code.Text = code_text + r" \h \z"
            toc.Update()

    return toc


def update_toc(doc, update_mode="full"):
    """更新已有目录。

    Args:
        doc: Word 文档对象。
        update_mode: 更新模式，"full" 为完整更新，"page_numbers_only" 只更新页码。

    Returns:
        bool: 是否成功更新。
    """
    if doc.TablesOfContents.Count == 0:
        return False

    toc = doc.TablesOfContents(1)
    if update_mode == "page_numbers_only":
        toc.UpdatePageNumbers()
    else:
        toc.Update()
    return True


def process_toc(doc, config=None, style_lookup=None, style_config_lookup=None):
    """主入口：检测并处理目录。

    流程：
    1. 检测文档中是否已有目录
    2. 有则更新，无则在英文摘要后的下一页插入"目录"标题 + 自动目录
    3. 插入分节符（下一页）隔离目录与正文

    Args:
        doc: Word 文档对象。
        config: 目录相关配置字典，可选。
        style_lookup: 样式名到 Word 样式对象的映射，插入新目录时必需。
        style_config_lookup: 样式名到配置字典的映射，插入新目录时必需。

    Returns:
        tuple: (success: bool, action: str, message: str)
            - success: 是否成功处理
            - action: 执行的操作 ("updated", "inserted", "skipped")
            - message: 描述信息
    """
    config = config or {}
    update_mode = config.get("update_mode", "full")

    # 声明 toc 引用以便后续打分节符
    toc = None

    if detect_toc(doc):
        success = update_toc(doc, update_mode)
        if success:
            toc = doc.TablesOfContents(1)
            msg = f"已更新目录（模式: {update_mode}）"
            action = "updated"
        else:
            return False, "skipped", "目录更新失败"
    else:
        insertion_point = find_toc_insertion_point(doc)
        if insertion_point:
            try:
                title_para = insert_toc_title(
                    doc, insertion_point, style_lookup or {}, style_config_lookup or {}
                )
                toc_range = title_para.Range.Duplicate
                toc_range.Collapse(0)  # wdCollapseEnd
                toc = insert_toc(doc, toc_range)
                msg = "已在英文摘要后的下一页插入目录标题与自动目录"
                action = "inserted"
            except Exception as e:
                return False, "skipped", f"插入目录失败: {e}"
        else:
            return False, "skipped", "未找到英文关键词行，无法确定目录插入位置"

    return True, action, msg
