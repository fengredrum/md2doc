"""Semantic analysis engine for Markdown documents.

Performs a single-pass parse of a Markdown document and runs six check
categories against it, generating a structured modification suggestions
report ({stem}-修改建议.md) in the same directory as the input.

Check categories:
  1. Heading hierarchy — skipped levels, missing/multiple H1, long titles
  2. Required 公文 sections — missing elements common in Chinese gov docs
  3. Format issues — broken images, table problems
  4. Content issues — empty sections, placeholders, duplicates, length
  5. Numbered list continuity — gaps, restarts
  6. Chinese punctuation — half-width marks, missing sentence endings
"""

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


# ── Data structures ───────────────────────────────────────────────────


@dataclass
class Heading:
    """A parsed Markdown heading."""
    level: int
    text: str
    line_num: int  # 1-based


@dataclass
class ImageRef:
    """A parsed Markdown image reference."""
    alt_text: str
    path: str
    line_num: int


@dataclass
class OrderedList:
    """A contiguous ordered list in the document."""
    items: list[tuple[int, int]]  # [(number, line_num), ...]
    start_line: int
    end_line: int


@dataclass
class Issue:
    """A single issue found during analysis."""
    severity: str   # "error", "warning", "info"
    category: str   # e.g. "heading_hierarchy", "content"
    line_num: int | None
    message: str
    snippet: str = ""


@dataclass
class DocumentAnalysis:
    """Complete analysis result for a document."""
    path: Path
    lines: list[str] = field(default_factory=list)
    headings: list[Heading] = field(default_factory=list)
    images: list[ImageRef] = field(default_factory=list)
    ordered_lists: list[OrderedList] = field(default_factory=list)
    html_table_count: int = 0
    html_table_open_count: int = 0
    html_table_close_count: int = 0
    issues: list[Issue] = field(default_factory=list)


# ── Category labels (Chinese) ─────────────────────────────────────────

CATEGORY_LABELS = {
    "heading_hierarchy": "标题层级问题",
    "required_sections": "必要章节检查",
    "format_issues": "格式问题",
    "content": "内容问题",
    "list_continuity": "编号连续性",
    "punctuation": "中文标点规范",
}

SEVERITY_LABELS = {
    "error": "错误",
    "warning": "警告",
    "info": "提示",
}

SEVERITY_SORT = {"error": 0, "warning": 1, "info": 2}


# ── Patterns ──────────────────────────────────────────────────────────

HEADING_RE = re.compile(r"^(#{1,6})\s+(.+)$")
IMAGE_RE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
ORDERED_ITEM_RE = re.compile(r"^(\s*)(\d+)[.)]\s+")
TABLE_OPEN_RE = re.compile(r"<\s*table[\s>]", re.IGNORECASE)
TABLE_CLOSE_RE = re.compile(r"<\s*/\s*table\s*>", re.IGNORECASE)

# Half-width punctuation that should be full-width in Chinese prose
HALF_TO_FULL = {
    ",": "，", ".": "。", ":": "：", ";": "；",
    "?": "？", "!": "！", "(": "（", ")": "）",
}

# Text that looks like a placeholder
PLACEHOLDER_RE = re.compile(
    r"(TODO|TBD|FIXME|XXX|待填写|待补充|此处填写|（.*填写.*）|"
    r"\{\{.*?\}\}|<.*填写.*>)",
    re.IGNORECASE,
)

# Chinese gov document expected sections (keywords to look for in headings)
REQUIRED_PATTERNS = [
    ("公文标题", re.compile(r"关于.*(通知|报告|请示|批复|函|意见|决定|通报|公告|通告)")),
    ("发文机关/单位名称", re.compile(r"(学校|大学|学院|公司|企业|部门|单位|机构|委员会|办公室|局|厅|部|处|组)")),
    ("日期信息", re.compile(r"(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|二[〇○零一二三四五六七八九]+年)")),
    ("附件/附录", re.compile(r"附件|附录|附表|附图")),
]


# ── Parsing ───────────────────────────────────────────────────────────


def _parse_document(input_path: Path) -> DocumentAnalysis:
    """Single-pass line-by-line parser.

    Extracts headings, images, ordered lists, and HTML table markers.
    Does not build a full AST — stores only metadata needed for checks.
    """
    lines = input_path.read_text(encoding="utf-8").splitlines(keepends=False)
    doc = DocumentAnalysis(path=input_path, lines=lines)

    in_html_table = False
    in_code_block = False
    current_list_items: list[tuple[int, int]] = []
    current_list_indent: int | None = None

    for i, line in enumerate(lines):
        line_num = i + 1

        # Code block fence detection (simple — ``` or ~~~)
        stripped = line.strip()
        if stripped.startswith("```") or stripped.startswith("~~~"):
            in_code_block = not in_code_block
            # Finalize any open list
            if current_list_items:
                doc.ordered_lists.append(OrderedList(
                    items=current_list_items,
                    start_line=current_list_items[0][1],
                    end_line=current_list_items[-1][1],
                ))
                current_list_items = []
                current_list_indent = None
            continue

        if in_code_block:
            continue

        # ── Headings ──
        heading_m = HEADING_RE.match(line)
        if heading_m:
            # Finalize any open list
            if current_list_items:
                doc.ordered_lists.append(OrderedList(
                    items=current_list_items,
                    start_line=current_list_items[0][1],
                    end_line=current_list_items[-1][1],
                ))
                current_list_items = []
                current_list_indent = None

            level = len(heading_m.group(1))
            text = heading_m.group(2).strip()
            doc.headings.append(Heading(level=level, text=text, line_num=line_num))
            continue

        # ── Images ──
        for img_m in IMAGE_RE.finditer(line):
            alt_text = img_m.group(1)
            img_path = img_m.group(2)
            doc.images.append(ImageRef(
                alt_text=alt_text,
                path=img_path,
                line_num=line_num,
            ))

        # ── HTML tables ──
        if TABLE_OPEN_RE.search(line):
            if not in_html_table:
                doc.html_table_open_count += 1
                in_html_table = True
        if TABLE_CLOSE_RE.search(line):
            if in_html_table:
                doc.html_table_close_count += 1
                in_html_table = False
        if in_html_table:
            doc.html_table_count = max(doc.html_table_count, 1)
        # Count complete tables (open + close on different matches)
        # We use the counts at the end

        # ── Ordered lists ──
        list_m = ORDERED_ITEM_RE.match(line)
        if list_m:
            indent = len(list_m.group(1))
            number = int(list_m.group(2))

            if current_list_indent is None or indent == current_list_indent:
                # Same list or new list at same indent
                current_list_items.append((number, line_num))
                current_list_indent = indent
            else:
                # Different indent — new list
                if current_list_items:
                    doc.ordered_lists.append(OrderedList(
                        items=current_list_items,
                        start_line=current_list_items[0][1],
                        end_line=current_list_items[-1][1],
                    ))
                current_list_items = [(number, line_num)]
                current_list_indent = indent
        else:
            # Non-list line ends the current list
            if current_list_items and stripped:
                doc.ordered_lists.append(OrderedList(
                    items=current_list_items,
                    start_line=current_list_items[0][1],
                    end_line=current_list_items[-1][1],
                ))
                current_list_items = []
                current_list_indent = None

    # Finalize any remaining list
    if current_list_items:
        doc.ordered_lists.append(OrderedList(
            items=current_list_items,
            start_line=current_list_items[0][1],
            end_line=current_list_items[-1][1],
        ))

    return doc


# ── Checkers ──────────────────────────────────────────────────────────


def _check_heading_hierarchy(doc: DocumentAnalysis) -> list[Issue]:
    """Check for heading hierarchy issues."""
    issues: list[Issue] = []

    if not doc.headings:
        issues.append(Issue(
            severity="error",
            category="heading_hierarchy",
            line_num=None,
            message="文档中未找到任何标题（H1-H6），建议添加标题以建立文档结构",
        ))
        return issues

    # Count H1s
    h1_count = sum(1 for h in doc.headings if h.level == 1)
    if h1_count == 0:
        issues.append(Issue(
            severity="error",
            category="heading_hierarchy",
            line_num=None,
            message="缺少一级标题（H1），公文应有文档主标题",
        ))
    elif h1_count > 1:
        issues.append(Issue(
            severity="warning",
            category="heading_hierarchy",
            line_num=doc.headings[0].line_num,
            message=f"存在 {h1_count} 个一级标题（H1），通常公文只有一个主标题",
        ))

    # Check for skipped levels
    for i in range(len(doc.headings) - 1):
        curr = doc.headings[i]
        nxt = doc.headings[i + 1]
        if nxt.level > curr.level + 1:
            issues.append(Issue(
                severity="warning",
                category="heading_hierarchy",
                line_num=nxt.line_num,
                message=(
                    f"标题层级跳跃: H{curr.level} → H{nxt.level}，缺少 H{curr.level + 1}"
                ),
                snippet=nxt.text[:60],
            ))

    # Check for excessively long heading text
    for h in doc.headings:
        if len(h.text) > 50:
            issues.append(Issue(
                severity="info",
                category="heading_hierarchy",
                line_num=h.line_num,
                message=f"标题文字过长（{len(h.text)}字符），建议精简",
                snippet=h.text[:60] + ("…" if len(h.text) > 60 else ""),
            ))

    return issues


def _check_required_sections(doc: DocumentAnalysis) -> list[Issue]:
    """Check for expected 公文 sections."""
    issues: list[Issue] = []

    # Combine all heading text + first 200 lines for pattern search
    heading_texts = " ".join(h.text for h in doc.headings)
    body_prefix = " ".join(doc.lines[:200]) if len(doc.lines) > 200 else " ".join(doc.lines)

    search_text = heading_texts + "\n" + body_prefix

    for pattern_name, pattern in REQUIRED_PATTERNS:
        if not pattern.search(search_text):
            severity = "warning" if pattern_name in ("公文标题", "日期信息") else "info"
            issues.append(Issue(
                severity=severity,
                category="required_sections",
                line_num=None,
                message=f"未检测到「{pattern_name}」，建议检查是否遗漏",
            ))

    return issues


def _check_format_issues(doc: DocumentAnalysis) -> list[Issue]:
    """Check for format problems: broken image refs, table issues."""
    issues: list[Issue] = []

    # Check local image references
    doc_dir = doc.path.parent
    for img in doc.images:
        # Skip URLs
        if img.path.startswith(("http://", "https://", "data:")):
            continue
        # Resolve relative path
        img_full = doc_dir / img.path
        if not img_full.exists():
            issues.append(Issue(
                severity="error",
                category="format_issues",
                line_num=img.line_num,
                message=f"图片文件不存在: {img.path}",
                snippet=img.path,
            ))

    # Check HTML table tag balance
    if doc.html_table_open_count != doc.html_table_close_count:
        issues.append(Issue(
            severity="warning",
            category="format_issues",
            line_num=None,
            message=(
                f"HTML 表格标签不匹配: "
                f"<table> 出现 {doc.html_table_open_count} 次, "
                f"</table> 出现 {doc.html_table_close_count} 次"
            ),
        ))

    return issues


def _check_content(doc: DocumentAnalysis) -> list[Issue]:
    """Check for content issues: empty sections, placeholders, duplicates, length."""
    issues: list[Issue] = []

    # ── Empty sections ──
    for i, h in enumerate(doc.headings):
        start_line = h.line_num
        # Find the next heading at the SAME or HIGHER level, not a sub-heading.
        # This ensures a parent section's content includes all its sub-sections.
        end_line = len(doc.lines) + 1  # default: end of document
        for j in range(i + 1, len(doc.headings)):
            next_h = doc.headings[j]
            if next_h.level <= h.level:  # same or higher level → section boundary
                end_line = next_h.line_num
                break
        # Extract content between this heading and the section boundary
        content_lines = doc.lines[start_line:end_line - 1]  # exclude boundary heading
        # Filter out blank lines and sub-headings (sub-sections are valid content)
        non_empty = [l for l in content_lines if l.strip() and not HEADING_RE.match(l)]
        if len(non_empty) == 0:  # truly empty section
            issues.append(Issue(
                severity="warning",
                category="content",
                line_num=h.line_num,
                message=f"「{h.text[:40]}」下方内容为空或仅有空行",
                snippet=h.text[:60],
            ))

    # ── Placeholder text ──
    for i, line in enumerate(doc.lines):
        if PLACEHOLDER_RE.search(line):
            issues.append(Issue(
                severity="warning",
                category="content",
                line_num=i + 1,
                message="检测到占位符文本，可能尚未完成填写",
                snippet=line.strip()[:80],
            ))

    # Duplicate paragraphs check — disabled per user request
    # # ── Duplicate paragraphs ──
    # seen: dict[str, int] = {}
    # ... (omitted for brevity)

    # Overly long paragraphs check — disabled per user request
    # # ── Overly long paragraphs ──
    # ... (omitted for brevity)

    return issues


def _check_list_continuity(doc: DocumentAnalysis) -> list[Issue]:
    """Check ordered lists for gaps and restarts."""
    issues: list[Issue] = []

    for ol in doc.ordered_lists:
        if len(ol.items) < 2:
            continue

        numbers = [item[0] for item in ol.items]
        lines = [item[1] for item in ol.items]

        # Restart detection: if first number != 1, it's suspicious
        if numbers[0] != 1:
            issues.append(Issue(
                severity="info",
                category="list_continuity",
                line_num=lines[0],
                message=f"编号列表未从 1 开始（起始编号为 {numbers[0]}）",
            ))

        # Gap detection
        for j in range(len(numbers) - 1):
            expected = numbers[j] + 1
            actual = numbers[j + 1]
            if actual != expected:
                issues.append(Issue(
                    severity="warning",
                    category="list_continuity",
                    line_num=lines[j + 1],
                    message=(
                        f"编号不连续: {numbers[j]} 之后应为 {expected}，"
                        f"实际为 {actual}"
                    ),
                ))

    return issues


def _check_punctuation(doc: DocumentAnalysis) -> list[Issue]:
    """Check for Chinese punctuation issues in prose paragraphs."""
    issues: list[Issue] = []

    heading_lines = {h.line_num for h in doc.headings}

    # Patterns for excluding non-prose contexts from half-width checks
    # Markdown image syntax: ![alt](path) — the ! ( ) are markup, not prose
    MD_IMAGE_RE = re.compile(r"!\[[^\]]*\]\([^)]*\)")
    # Markdown link syntax: [text](url) — the ( ) are markup
    MD_LINK_RE = re.compile(r"\[[^\]]*\]\([^)]*\)")
    # HTML entities: &quot; &amp; &#160; etc. — the ; is markup
    HTML_ENTITY_RE = re.compile(r"&[a-zA-Z#][a-zA-Z0-9]*;")
    # Numeric ratios: 1:17.00, 17.93:1, ≥1:20 — the : is math notation
    RATIO_RE = re.compile(r"\d[\d.]*\s*:\s*\d")

    in_code_block = False

    for i, line in enumerate(doc.lines):
        line_num = i + 1
        stripped = line.strip()

        # Skip headings
        if line_num in heading_lines:
            continue

        # Code fence toggle
        if stripped.startswith("```") or stripped.startswith("~~~"):
            in_code_block = not in_code_block
            continue
        if in_code_block:
            continue

        # Table row detection
        if stripped.startswith("|") and stripped.endswith("|"):
            continue
        if re.match(r"^\|?[\s\-:]+\|", stripped):  # separator row
            continue

        # Skip empty lines
        if not stripped:
            continue

        # ── Half-width punctuation in Chinese context ──
        has_cjk = bool(re.search(r"[一-鿿㐀-䶿]", stripped))
        if has_cjk:
            # Build a "clean" version of the line with markup/notation
            # contexts removed so half-width punctuation inside them is
            # not incorrectly flagged as Chinese punctuation errors.
            clean = stripped

            # Remove markdown images: ![alt](path)
            clean = MD_IMAGE_RE.sub("", clean)
            # Remove markdown links: [text](url)
            clean = MD_LINK_RE.sub("", clean)
            # Remove HTML entities: &quot; &amp; &#160; etc.
            clean = HTML_ENTITY_RE.sub("", clean)
            # Remove numeric ratios like 1:17.00, 17.93:1
            clean = RATIO_RE.sub("[ratio]", clean)

            for half, full in HALF_TO_FULL.items():
                # Skip dots — they appear in too many legitimate contexts
                # (numbers, URLs, abbreviations, etc.)
                if half == ".":
                    continue
                if half in clean:
                    count = clean.count(half)
                    if count > 0:
                        issues.append(Issue(
                            severity="info",
                            category="punctuation",
                            line_num=line_num,
                            message=(
                                f"检测到半角标点「{half}」，"
                                f"建议使用全角「{full}」（出现 {count} 次）"
                            ),
                            snippet=stripped[:80],
                        ))
                        break  # one issue per line for punctuation

            # ── Missing sentence-ending punctuation ──
            # Only check lines that look like complete sentences
            # (at least 15 chars, has Chinese text, not ending with list/item marker)
            if len(stripped) >= 15 and has_cjk:
                if not re.search(r"[。！？……）\)]$", stripped):
                    # Don't flag lines ending with colon or comma (likely clause breaks)
                    if not re.search(r"[：:，,；;、]$", stripped):
                        # Only flag if it ends with a Chinese char (truly missing punctuation)
                        if re.search(r"[一-鿿]$", stripped):
                            issues.append(Issue(
                                severity="info",
                                category="punctuation",
                                line_num=line_num,
                                message="段落末尾缺少句号（。）或相应标点",
                                snippet=stripped[-60:],
                            ))

    return issues


# ── Report generation ─────────────────────────────────────────────────


def _generate_report(doc: DocumentAnalysis, issues: list[Issue]) -> str:
    """Generate the 修改建议.md report content."""
    # Group issues by category
    by_category: dict[str, list[Issue]] = {}
    for iss in issues:
        by_category.setdefault(iss.category, []).append(iss)

    # Sort issues within each category by line_num (None at end)
    for cat_issues in by_category.values():
        cat_issues.sort(key=lambda x: (x.line_num is None, x.line_num or 0))

    # Category display order
    category_order = [
        "heading_hierarchy",
        "required_sections",
        "format_issues",
        "content",
        "list_continuity",
        "punctuation",
    ]

    # Count severities
    error_count = sum(1 for iss in issues if iss.severity == "error")
    warning_count = sum(1 for iss in issues if iss.severity == "warning")
    info_count = sum(1 for iss in issues if iss.severity == "info")

    # Build report
    lines: list[str] = []
    lines.append(f"# 修改建议 — {doc.path.name}")
    lines.append("")
    lines.append(f"> 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"> 分析文件: {doc.path}")
    lines.append(f"> 总行数: {len(doc.lines)}")
    lines.append(
        f"> 发现问题: {len(issues)} 项"
        f"（错误 {error_count} 项，警告 {warning_count} 项，提示 {info_count} 项）"
    )
    lines.append("")
    lines.append("---")
    lines.append("")

    section_num = 0
    for cat_key in category_order:
        cat_issues = by_category.get(cat_key, [])
        if not cat_issues:
            continue

        section_num += 1
        cat_name = CATEGORY_LABELS.get(cat_key, cat_key)
        lines.append(f"## {_cn_number(section_num)}、{cat_name}（{len(cat_issues)}项）")
        lines.append("")
        lines.append("| 行号 | 严重度 | 问题描述 |")
        lines.append("|------|--------|----------|")

        for iss in cat_issues:
            line_display = str(iss.line_num) if iss.line_num is not None else "--"
            sev_display = SEVERITY_LABELS.get(iss.severity, iss.severity)
            msg = iss.message.replace("|", "\\|")  # escape pipes
            lines.append(f"| {line_display} | {sev_display} | {msg} |")

        lines.append("")

    if not any(by_category.get(c, []) for c in category_order):
        lines.append("> ✅ 未发现明显问题，文档质量良好。")
        lines.append("")

    return "\n".join(lines)


def _cn_number(n: int) -> str:
    """Convert integer to Chinese number (1-99)."""
    digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
    if n <= 10:
        return digits[n]
    elif n < 20:
        return f"十{digits[n - 10]}"
    elif n < 100:
        tens = n // 10
        ones = n % 10
        if ones == 0:
            return f"{digits[tens]}十"
        return f"{digits[tens]}十{digits[ones]}"
    return str(n)


# ── Public API ────────────────────────────────────────────────────────


def analyze(input_path: str | Path) -> Path:
    """Analyze a Markdown document and generate modification suggestions.

    Args:
        input_path: Path to the input Markdown file.

    Returns:
        Path to the generated `{stem}-修改建议.md` file.

    Raises:
        FileNotFoundError: input_path does not exist.
        ValueError: input_path is not a .md file or cannot be read.
    """
    input_path = Path(input_path).resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if not input_path.is_file():
        raise FileNotFoundError(f"Input path is not a file: {input_path}")

    # Parse document
    doc = _parse_document(input_path)

    # Run all checkers
    issues: list[Issue] = []
    issues.extend(_check_heading_hierarchy(doc))
    issues.extend(_check_required_sections(doc))
    issues.extend(_check_format_issues(doc))
    issues.extend(_check_content(doc))
    issues.extend(_check_list_continuity(doc))
    issues.extend(_check_punctuation(doc))

    # Generate report
    report = _generate_report(doc, issues)

    # Write to same directory as input
    output_path = input_path.parent / f"{input_path.stem}-修改建议.md"
    output_path.write_text(report, encoding="utf-8")

    return output_path
