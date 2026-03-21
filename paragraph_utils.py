def normalize_paragraph_text(text):
    """清理 Word 段落文本中的段落结束符和首尾空白。"""
    return text.replace("\r", "").replace("\x07", "").strip()



def get_next_paragraph(doc, start_index):
    """获取指定位置之后的下一个段落。"""
    if start_index >= doc.Paragraphs.Count:
        return None
    return doc.Paragraphs(start_index + 1)



def get_next_non_empty_paragraph(doc, start_index):
    """获取指定位置之后的下一个非空段落。"""
    for index in range(start_index + 1, doc.Paragraphs.Count + 1):
        paragraph = doc.Paragraphs(index)
        text = normalize_paragraph_text(paragraph.Range.Text)
        if text:
            return paragraph
    return None



def get_previous_non_empty_paragraph(doc, start_index):
    """获取指定位置之前的上一个非空段落。"""
    for index in range(start_index - 1, 0, -1):
        paragraph = doc.Paragraphs(index)
        text = normalize_paragraph_text(paragraph.Range.Text)
        if text:
            return paragraph
    return None
