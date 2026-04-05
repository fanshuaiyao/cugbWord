"""提供目录检测、插入和更新功能。

目录插入策略：始终在英文摘要（English Abstract）区块之后的下一页插入。
流程：分页 → 插入"目录"标题并套样式 → 在标题之后插入自动目录域 → 可选添加分节符隔离。
"""

import re

from paragraph_utils import normalize_paragraph_text
from style_operations import apply_direct_font_format, apply_direct_paragraph_format
from word_constants import WD_BREAK_PAGE, WD_SECTION_BREAK_NEXT_PAGE


def detect_toc(doc):
    """检测文档中是否存在目录。

    Args:
        doc: Word 文档对象。

    Returns:
        bool: 是否存在目录。
    """
    return doc.TablesOfContents.Count > 0


def remove_existing_tocs(doc):
    """删除文档中所有现有目录及其所在段落。

    Args:
        doc: Word 文档对象。

    Returns:
        bool: 是否删除了目录。
    """
    removed = False
    while doc.TablesOfContents.Count > 0:
        toc = doc.TablesOfContents(1)
        # 获取目录的范围
        toc_range = toc.Range
        start_pos = toc_range.Start
        end_pos = toc_range.End

        # 删除目录对象本身
        toc.Delete()

        # 删除目录所占用的文本范围
        try:
            deletion_range = doc.Range(start_pos, end_pos)
            deletion_range.Delete()
        except Exception:
            pass  # 范围可能已无效，忽略
        removed = True
    return removed


def insert_section_break_after_range(doc, range_obj, break_type=WD_SECTION_BREAK_NEXT_PAGE):
    """在指定范围之后插入分节符。

    Args:
        doc: Word 文档对象。
        range_obj: Range 对象，分节符将插入在此范围之后。
        break_type: 分节符类型，默认下一页分节符。

    Returns:
        Section: 新创建的分节对象。
    """
    # 确保范围折叠到末尾
    range_obj.Collapse(0)  # wdCollapseEnd = 0

    # 在当前位置插入分节符
    range_obj.InsertBreak(break_type)

    # 返回新创建的分节
    return doc.Sections(doc.Sections.Count)


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
    """在指定位置插入分页符和"目录"标题文字并应用 contents_title 样式。

    插入逻辑：
    1. 先在当前位置插入分页符（让目录独占一页）
    2. 插入目录标题段落
    3. 应用样式，但禁用段前分页（避免重复分页）

    Args:
        doc: Word 文档对象。
        range_obj: 插入位置的 Range 对象。
        style_lookup: 样式名到 Word 样式对象的映射。
        style_config_lookup: 样式名到配置字典的映射。

    Returns:
        Paragraph 对象: 新插入的"目录"标题段落。
    """
    # 步骤1：在当前位置插入分页符，让目录标题独占一页
    range_obj.Collapse(0)  # wdCollapseEnd = 0
    range_obj.InsertBreak(WD_BREAK_PAGE)

    # 步骤2：在分页符之后插入目录标题
    range_obj.InsertAfter("目录\r")
    title_para = doc.Range(range_obj.Start, range_obj.End).Paragraphs(1)

    # 步骤3：应用样式，但禁用段前分页（样式中可能有 page_break_before）
    style = style_lookup.get("contents_title")
    style_config = style_config_lookup.get("contents_title")
    if style is not None and style_config is not None:
        title_para.Range.Style = style.NameLocal
        apply_direct_font_format(title_para, style_config)
        apply_direct_paragraph_format(title_para, style_config)
        # 关键：手动禁用段前分页，避免样式的 page_break_before 导致重复分页
        title_para.Range.ParagraphFormat.PageBreakBefore = False

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
    """主入口：检测并处理目录（删除旧目录、插入新目录、添加分节符）。

    流程：
    1. 如配置允许，删除文档中所有现有目录
    2. 在英文摘要后的下一页插入"目录"标题 + 自动目录
    3. 如配置允许，在目录后插入空段落 + 下一页分节符

    Args:
        doc: Word 文档对象。
        config: 目录相关配置字典，可选。
        style_lookup: 样式名到 Word 样式对象的映射，插入新目录时必需。
        style_config_lookup: 样式名到配置字典的映射，插入新目录时必需。

    Returns:
        tuple: (success: bool, action: str, message: str)
            - success: 是否成功处理
            - action: 执行的操作 ("inserted", "failed", "skipped")
            - message: 描述信息
    """
    config = config or {}
    add_section_break = config.get("add_section_break_after", True)
    force_replace = config.get("force_replace_existing", True)

    # 步骤1：如有现有目录且配置要求替换，先删除
    had_existing_toc = detect_toc(doc)
    if had_existing_toc and force_replace:
        try:
            remove_existing_tocs(doc)
            action_note = "已删除现有目录，"
        except Exception as e:
            return False, "failed", f"删除现有目录失败: {e}"
    elif had_existing_toc and not force_replace:
        # 不强制替换时，尝试更新现有目录
        try:
            update_mode = config.get("update_mode", "full")
            update_toc(doc, update_mode)
            return True, "updated", f"已更新现有目录（模式: {update_mode}）"
        except Exception as e:
            return False, "failed", f"更新目录失败: {e}"
    else:
        action_note = ""

    # 步骤2：插入新目录（标题 + 自动目录域）
    insertion_point = find_toc_insertion_point(doc)
    if insertion_point is None:
        return False, "skipped", "未找到英文关键词行，无法确定目录插入位置"

    try:
        # 插入目录标题
        title_para = insert_toc_title(
            doc, insertion_point, style_lookup or {}, style_config_lookup or {}
        )

        # 在标题后插入自动目录域
        toc_range = title_para.Range.Duplicate
        toc_range.Collapse(0)  # wdCollapseEnd
        toc = insert_toc(doc, toc_range)

        # 步骤3：如配置允许，在目录后插入三个空段落（正文样式）和分节符
        if add_section_break:
            normal_style = style_lookup.get("normal")

            # 获取目录范围，用于计算插入位置
            toc_end_pos = toc.Range.End

            # 插入三个空段落，并分别应用正文样式
            for i in range(3):
                # 在文档末尾（当前最后一个段落后）插入新段落
                insert_range = doc.Range(toc_end_pos, toc_end_pos)
                insert_range.Collapse(0)
                insert_range.InsertAfter("\r")

                # 新段落的起始位置就是原来的 toc_end_pos
                # 新段落的结束位置需要重新获取
                new_para_end = insert_range.End
                new_para_range = doc.Range(toc_end_pos, new_para_end)
                new_para = new_para_range.Paragraphs(1)

                # 应用正文样式到新段落
                if normal_style is not None:
                    new_para.Range.Style = normal_style.NameLocal

                # 更新位置，为下一个段落做准备
                toc_end_pos = new_para_end

            # 在第三个空段落末尾插入分节符
            final_range = doc.Range(toc_end_pos, toc_end_pos)
            final_range.Collapse(0)
            final_range.InsertBreak(WD_SECTION_BREAK_NEXT_PAGE)

            section_msg = "，已添加三个空段落（正文样式）和下一页分节符"
        else:
            section_msg = ""

        msg = f"{action_note}已在英文摘要后插入目录标题与自动目录{section_msg}"
        action = "inserted"

    except Exception as e:
        return False, "failed", f"插入目录失败: {e}"

    return True, action, msg
