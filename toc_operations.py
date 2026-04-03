"""提供目录检测、插入和更新功能。"""


def detect_toc(doc):
    """检测文档中是否存在目录。

    Args:
        doc: Word 文档对象。

    Returns:
        bool: 是否存在目录。
    """
    return doc.TablesOfContents.Count > 0


def find_toc_insertion_point(doc):
    """找到目录应该插入的位置（'目录'标题之后）。

    Args:
        doc: Word 文档对象。

    Returns:
        Range 对象或 None: 目录标题段落的 Range，在其后插入目录。
    """
    for para in doc.Paragraphs:
        text = para.Range.Text.strip()
        if text == "目录":
            return para.Range
    return None


def insert_toc(doc, range_obj, upper_level=1, lower_level=4):
    """在指定位置插入新目录。

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
        UseHyperlinks=False
    )
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


def process_toc(doc, config=None):
    """主入口：检测并处理目录。

    流程：
    1. 检测文档中是否已有目录
    2. 有则更新，无则尝试在"目录"标题后插入

    Args:
        doc: Word 文档对象。
        config: 目录相关配置字典，可选。

    Returns:
        tuple: (success: bool, action: str, message: str)
            - success: 是否成功处理
            - action: 执行的操作 ("updated", "inserted", "skipped")
            - message: 描述信息
    """
    config = config or {}
    update_mode = config.get("update_mode", "full")

    if detect_toc(doc):
        # 已有目录，更新它
        success = update_toc(doc, update_mode)
        if success:
            return True, "updated", f"已更新目录（模式: {update_mode}）"
        return False, "skipped", "目录更新失败"
    else:
        # 没有目录，尝试插入
        insertion_point = find_toc_insertion_point(doc)
        if insertion_point:
            try:
                insert_toc(doc, insertion_point)
                return True, "inserted", "已在'目录'标题后插入新目录"
            except Exception as e:
                return False, "skipped", f"插入目录失败: {e}"
        return False, "skipped", "未找到'目录'标题，无法插入目录"
