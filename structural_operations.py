import re
from paragraph_utils import normalize_paragraph_text

def remove_empty_paragraphs(doc):
    """
    安全地删除文档中的无用空段落。
    采用“倒序遍历”策略以防止索引偏移。
    条件：纯文本为空、无图、无表、无分节符/分页符（\x0c）。

    Args:
        doc: Word 文档对象。
    Returns:
        int: 成功删除的空段落数量。
    """
    removed_count = 0
    # 倒序遍历：这样删除当前元素不会影响还没遍历到的索引
    for i in range(doc.Paragraphs.Count, 0, -1):
        para = doc.Paragraphs(i)
        rng = para.Range

        # 判断是否包含图片（包含 InlineShapes 或浮动 Shapes）
        if rng.InlineShapes.Count > 0 or rng.ShapeRange.Count > 0:
            continue

        # 判断是否在表格内或者包含表格
        try:
            if getattr(rng.Information(12), 'numerator', 0) or rng.Tables.Count > 0:  # 12 = wdWithInTable
                continue
        except Exception:
            pass # COM对象可能抛错，安全略过

        text = rng.Text
        # 判断是否包含控制字符（12 = 0x0c 代表分页符或分节符）
        if '\x0c' in text:
            continue

        # 标准化后若为空（即只剩换行、空白符等），安全删除
        norm_text = normalize_paragraph_text(text)
        if not norm_text:
            rng.Delete()
            removed_count += 1

    return removed_count
