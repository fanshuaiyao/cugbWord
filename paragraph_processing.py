import re

from paragraph_rules import (
    is_abstract_title_text,
    is_acknowledgements_title_text,
    is_appendix_title_text,
    is_contents_entry_text,
    is_contents_title_text,
    is_keywords_line_text,
    is_references_title_text,
    match_heading_style_id,
    match_paragraph_style_id,
)
from paragraph_utils import normalize_paragraph_text
from style_operations import apply_direct_font_format, apply_direct_paragraph_format
from word_constants import WD_STATISTIC_PAGES


ABSTRACT_MIN_CHAR_COUNT = 800
ABSTRACT_MAX_CHAR_COUNT = 1000
KEYWORDS_MIN_COUNT = 3
KEYWORDS_MAX_COUNT = 5
KEYWORDS_LABEL_PATTERN = re.compile(r"^\s*(关键词\s*[：:])")
KEYWORDS_NON_FULLWIDTH_SEPARATOR_PATTERN = re.compile(r"[,、；;]")
KEYWORDS_SPLIT_PATTERN = re.compile(r"[，,、；;]")
KEYWORDS_TRAILING_PUNCTUATION_PATTERN = re.compile(r"[，,、；;。.!！？?：:]$")



def count_non_whitespace_characters(text):
    """统计去除空白后的字符数。"""
    return len(re.sub(r"\s+", "", text))



def apply_keywords_label_format(paragraph):
    """将关键词行中的“关键词：”局部设置为加粗。"""
    raw_text = paragraph.Range.Text.replace("\r", "").replace("\x07", "")
    match = KEYWORDS_LABEL_PATTERN.match(raw_text)
    if match is None:
        return

    label_range = paragraph.Range.Duplicate
    label_range.Start = label_range.Start + match.start(1)
    label_range.End = label_range.Start + len(match.group(1))
    label_range.Font.Bold = True



def apply_paragraph_style(paragraph, style_id, style_lookup, style_config_lookup):
    """将指定样式应用到段落。"""
    style = style_lookup.get(style_id)
    style_config = style_config_lookup.get(style_id)
    if style is None or style_config is None:
        raise ValueError(f"段落匹配到未配置的样式: {style_id}")

    paragraph.Range.Style = style.NameLocal

    if style_id == "figure_block":
        normal_config = style_config_lookup.get("normal")
        if normal_config is None:
            raise ValueError("应用图片段落样式时缺少 normal 配置")

        apply_direct_font_format(paragraph, normal_config)
        apply_direct_paragraph_format(
            paragraph,
            style_config,
            space_before_override=normal_config["font"]["size"],
        )
        return

    apply_direct_paragraph_format(paragraph, style_config)

    if style_id == "keywords_line":
        apply_keywords_label_format(paragraph)



def append_keywords_validation_issues(validation_issues, paragraph_index, text):
    """校验关键词数量、分隔符与末尾标点。"""
    match = KEYWORDS_LABEL_PATTERN.match(text)
    if match is None:
        validation_issues.append(f"第 {paragraph_index} 段关键词行格式无法识别")
        return

    keywords_text = text[match.end() :].strip()
    if not keywords_text:
        validation_issues.append(f"第 {paragraph_index} 段关键词内容不能为空")
        return

    if KEYWORDS_NON_FULLWIDTH_SEPARATOR_PATTERN.search(keywords_text):
        validation_issues.append(f"第 {paragraph_index} 段关键词应使用全角逗号“，”分隔")

    if KEYWORDS_TRAILING_PUNCTUATION_PATTERN.search(keywords_text):
        validation_issues.append(f"第 {paragraph_index} 段关键词末尾不应有标点符号")

    keyword_items = [
        keyword.strip()
        for keyword in KEYWORDS_SPLIT_PATTERN.split(keywords_text)
        if keyword.strip()
    ]
    keyword_count = len(keyword_items)
    if keyword_count < KEYWORDS_MIN_COUNT or keyword_count > KEYWORDS_MAX_COUNT:
        validation_issues.append(
            f"第 {paragraph_index} 段关键词数量应为 {KEYWORDS_MIN_COUNT}~{KEYWORDS_MAX_COUNT} 个，当前为 {keyword_count} 个"
        )



def append_abstract_validation_issues(
    doc,
    validation_issues,
    title_index,
    body_range_start,
    body_range_end,
    body_texts,
):
    """校验摘要字数与页数。"""
    if title_index is None:
        return

    abstract_text = "".join(body_texts)
    abstract_char_count = count_non_whitespace_characters(abstract_text)
    if (
        abstract_char_count < ABSTRACT_MIN_CHAR_COUNT
        or abstract_char_count > ABSTRACT_MAX_CHAR_COUNT
    ):
        validation_issues.append(
            f"第 {title_index} 段摘要字数应为 {ABSTRACT_MIN_CHAR_COUNT}~{ABSTRACT_MAX_CHAR_COUNT} 字，当前为 {abstract_char_count} 字"
        )

    if body_range_start is None or body_range_end is None:
        return

    abstract_range = doc.Range(body_range_start, body_range_end)
    page_count = abstract_range.ComputeStatistics(WD_STATISTIC_PAGES)
    if page_count > 1:
        validation_issues.append(
            f"第 {title_index} 段摘要内容应限制在 1 页内，当前为 {page_count} 页"
        )



def finalize_current_block(validation_issues, current_block, abstract_state, doc):
    """统一收尾当前区块。"""
    if current_block == "abstract":
        append_abstract_validation_issues(
            doc,
            validation_issues,
            abstract_state["title_index"],
            abstract_state["range_start"],
            abstract_state["range_end"],
            abstract_state["body_texts"],
        )

    abstract_state["title_index"] = None
    abstract_state["range_start"] = None
    abstract_state["range_end"] = None
    abstract_state["body_texts"] = []



def apply_paragraph_styles(doc, style_lookup, style_config_lookup):
    """遍历文档段落并按识别结果应用对应样式，同时返回内容校验结果。"""
    processed_count = 0
    validation_issues = []
    validation_counts = {"abstract_count": 0, "keywords_count": 0}
    current_block = None
    abstract_state = {
        "title_index": None,
        "range_start": None,
        "range_end": None,
        "body_texts": [],
    }

    for index in range(1, doc.Paragraphs.Count + 1):
        paragraph = doc.Paragraphs(index)
        text = normalize_paragraph_text(paragraph.Range.Text)
        if not text:
            continue

        style_id = match_paragraph_style_id(doc, paragraph, index, text)

        if style_id in {"figure_block", "caption"}:
            apply_paragraph_style(paragraph, style_id, style_lookup, style_config_lookup)
            processed_count += 1
            continue

        if is_abstract_title_text(text):
            finalize_current_block(validation_issues, current_block, abstract_state, doc)
            current_block = "abstract"
            validation_counts["abstract_count"] += 1
            abstract_state["title_index"] = index
            apply_paragraph_style(paragraph, "abstract_title", style_lookup, style_config_lookup)
            processed_count += 1
            continue

        if is_references_title_text(text):
            finalize_current_block(validation_issues, current_block, abstract_state, doc)
            current_block = "references"
            apply_paragraph_style(paragraph, "references_title", style_lookup, style_config_lookup)
            processed_count += 1
            continue

        if is_contents_title_text(text):
            finalize_current_block(validation_issues, current_block, abstract_state, doc)
            current_block = "contents"
            apply_paragraph_style(paragraph, "contents_title", style_lookup, style_config_lookup)
            processed_count += 1
            continue

        if is_acknowledgements_title_text(text):
            finalize_current_block(validation_issues, current_block, abstract_state, doc)
            current_block = "acknowledgements"
            apply_paragraph_style(paragraph, "acknowledgements_title", style_lookup, style_config_lookup)
            processed_count += 1
            continue

        if is_appendix_title_text(text):
            finalize_current_block(validation_issues, current_block, abstract_state, doc)
            current_block = "appendix"
            apply_paragraph_style(paragraph, "appendix_title", style_lookup, style_config_lookup)
            processed_count += 1
            continue

        if is_keywords_line_text(text):
            apply_paragraph_style(paragraph, "keywords_line", style_lookup, style_config_lookup)
            validation_counts["keywords_count"] += 1
            append_keywords_validation_issues(validation_issues, index, text)
            finalize_current_block(validation_issues, current_block, abstract_state, doc)
            current_block = None
            processed_count += 1
            continue

        if current_block == "contents" and is_contents_entry_text(text):
            apply_paragraph_style(paragraph, "contents_entry", style_lookup, style_config_lookup)
            processed_count += 1
            continue

        heading_style_id = match_heading_style_id(text)
        if heading_style_id is not None:
            finalize_current_block(validation_issues, current_block, abstract_state, doc)
            current_block = None
            apply_paragraph_style(paragraph, heading_style_id, style_lookup, style_config_lookup)
            processed_count += 1
            continue

        if current_block == "contents":
            finalize_current_block(validation_issues, current_block, abstract_state, doc)
            current_block = None

        if current_block == "abstract":
            if abstract_state["range_start"] is None:
                abstract_state["range_start"] = paragraph.Range.Start
            abstract_state["range_end"] = paragraph.Range.End
            abstract_state["body_texts"].append(text)
            apply_paragraph_style(paragraph, "abstract_body", style_lookup, style_config_lookup)
            processed_count += 1
            continue

        if current_block == "references":
            apply_paragraph_style(paragraph, "reference_entry", style_lookup, style_config_lookup)
            processed_count += 1
            continue

        if current_block == "acknowledgements":
            apply_paragraph_style(paragraph, "acknowledgements_body", style_lookup, style_config_lookup)
            processed_count += 1
            continue

        if current_block == "appendix":
            apply_paragraph_style(paragraph, "appendix_body", style_lookup, style_config_lookup)
            processed_count += 1
            continue

        apply_paragraph_style(paragraph, "normal", style_lookup, style_config_lookup)
        processed_count += 1

    finalize_current_block(validation_issues, current_block, abstract_state, doc)
    return processed_count, validation_issues, validation_counts
