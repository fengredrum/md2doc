"""MCP server exposing Markdown-to-Word conversion and analysis tools.

Usage:
    uv run md2doc-mcp

Configuration:
    DEFAULT_FORMAT_SPEC_PATH — path to the default format spec markdown file.
    Modify this constant to change the fallback format used when the user
    does not provide their own format spec or reference template.
"""

from pathlib import Path

from mcp.server.fastmcp import FastMCP

from md2doc.converter import (
    ConversionError,
    InputFileNotFoundError,
    PandocConversionError,
    PandocNotFoundError,
    PandocTimeoutError,
    convert,
)
from md2doc.analyzer import analyze

# ── Configurable default format spec path ─────────────────────────────
# Change this to point to your preferred default format specification.
# Set to None to skip loading any default format file (falls back to
# hardcoded DEFAULT_FORMAT_SPEC in the core module).
DEFAULT_FORMAT_SPEC_PATH: str | None = str(
    Path(__file__).parent.parent.parent / "格式要求.md"
)

# ── MCP server instance ───────────────────────────────────────────────

mcp = FastMCP(
    "md2doc",
    instructions="""Convert Markdown to Word (.docx) with Chinese government document formatting (公文格式规范).

## Conversion workflow

Before converting, ask the user about format requirements:
1. Does the user have a format specification file (格式要求.md style) or a reference DOCX template?
2. Format requirements can be provided as:
   - **format_spec**: A format description markdown file (like 格式要求.md) that specifies fonts, sizes, spacing, and alignment per style
   - **reference_docx**: A .docx template whose styles are used as an overlay on top of the format spec — pandoc reads styles from this template
   - **Both can be combined** — the reference template overlays on top of the format spec
3. If the user has no specific requirements, use the default format spec configured in the MCP source code (DEFAULT_FORMAT_SPEC_PATH)
4. If the user explicitly wants no formatting applied, pass skip_format=True for pure pandoc output

After conversion completes, ask the user whether they also want to generate a modification suggestions report (修改建议). If they say yes, call analyze_and_suggest with the same input path. Alternatively, pass auto_suggest=True to convert_markdown_to_docx to generate suggestions automatically in one call.

## Available tools
- **convert_markdown_to_docx**: Convert .md to .docx with Chinese government document formatting (公文格式). Supports custom format specs, reference templates, media extraction, and optional auto-generation of modification suggestions.
- **analyze_and_suggest**: Analyze a Markdown document's structure and content, generating a structured modification suggestions report (xx-修改建议.md) covering heading hierarchy, required sections, format issues, content problems, list continuity, and Chinese punctuation.
""",
)


# ── Tools ─────────────────────────────────────────────────────────────


@mcp.tool()
def convert_markdown_to_docx(
    input_path: str,
    output_path: str | None = None,
    format_spec: str | None = None,
    reference_docx: str | None = None,
    extract_media: str | None = None,
    skip_format: bool = False,
    no_cache_template: bool = False,
    auto_suggest: bool = False,
) -> str:
    """Convert a Markdown file to Word (.docx) with formatting.

    Applies Chinese government document formatting (公文格式规范) by default:
    - Headings: 黑体 (H1 三号16pt, H2 小三15pt, H3 四号14pt) bold, left-aligned
    - Body: 仿宋_GB2312 四号14pt, justified, 2-char first-line indent
    - Tables: 仿宋_GB2312 五号10.5pt, center-aligned
    - Line spacing: fixed 28pt for body/headings, single for tables

    Args:
        input_path: Absolute path to the input Markdown file.
        output_path: Output .docx path. Defaults to <input_dir>/<input_stem>.docx.
        format_spec: Path to a format description markdown file (like 格式要求.md).
            Overrides the default format. The file defines fonts, sizes, spacing,
            and alignment per Word style (H1/H2/H3/正文/表格).
        reference_docx: Path to a .docx template whose styles overlay on top of
            the format spec. Useful for matching an existing document's styles.
        extract_media: Directory to extract media files to.
        skip_format: If True, bypass ALL formatting (pure pandoc output).
        no_cache_template: If True, force regeneration of the cached reference
            template (needed when format spec changes).
        auto_suggest: If True, automatically generate a modification suggestions
            report (xx-修改建议.md) after conversion. The report covers heading
            hierarchy, required sections, format issues, content problems, list
            continuity, and Chinese punctuation.

    Returns:
        Summary string with input path, output path, file size, and optionally
        the modification suggestions report path and issue count.
    """
    # Resolve default format spec
    resolved_format_spec = format_spec
    if resolved_format_spec is None and not skip_format:
        if DEFAULT_FORMAT_SPEC_PATH and Path(DEFAULT_FORMAT_SPEC_PATH).exists():
            resolved_format_spec = DEFAULT_FORMAT_SPEC_PATH

    try:
        result_path = convert(
            input_path=input_path,
            output_path=output_path,
            format_spec_md=resolved_format_spec,
            reference_docx=reference_docx,
            extract_media=extract_media,
            skip_format=skip_format,
            no_cache_template=no_cache_template,
        )
    except InputFileNotFoundError as e:
        return f"错误：输入文件不存在 — {e}"
    except PandocNotFoundError as e:
        return f"错误：未找到 pandoc — {e}\n请安装 pandoc: https://pandoc.org/installing.html"
    except PandocConversionError as e:
        return f"错误：pandoc 转换失败 — {e}"
    except PandocTimeoutError as e:
        return f"错误：转换超时 — {e}"
    except ConversionError as e:
        return f"错误：转换失败 — {e}"
    except Exception as e:
        return f"未知错误: {type(e).__name__}: {e}"

    # Build summary
    size_bytes = result_path.stat().st_size
    if size_bytes >= 1_000_000:
        size_str = f"{size_bytes / 1_000_000:.1f} MB"
    elif size_bytes >= 1_000:
        size_str = f"{size_bytes / 1_000:.1f} KB"
    else:
        size_str = f"{size_bytes} bytes"

    parts = [
        "✅ 转换成功",
        f"输入文件: {input_path}",
        f"输出文件: {result_path}",
        f"文件大小: {size_str}",
    ]
    if resolved_format_spec:
        parts.append(f"格式规范: {resolved_format_spec}")
    if reference_docx:
        parts.append(f"参考模板: {reference_docx}")
    if skip_format:
        parts.append("格式: 无（纯 pandoc 输出）")

    # ── Auto-generate modification suggestions ──
    if auto_suggest:
        try:
            suggestions_path = analyze(input_path)
            report = suggestions_path.read_text(encoding="utf-8")
            import re as _re
            issue_lines = len(_re.findall(r"^\| \d+ \|", report, _re.MULTILINE))
            dash_lines = len(_re.findall(r"^\| -- \|", report, _re.MULTILINE))
            total_issues = issue_lines + dash_lines
            parts.append("")
            parts.append("---")
            parts.append(f"📋 修改建议已自动生成: {suggestions_path}")
            parts.append(f"发现问题: {total_issues} 项")
            parts.append(f"请查看 {suggestions_path.name} 了解详细修改建议。")
        except Exception as e:
            parts.append("")
            parts.append(f"⚠️ 自动生成修改建议失败: {type(e).__name__}: {e}")

    return "\n".join(parts)


@mcp.tool()
def analyze_and_suggest(
    input_path: str,
) -> str:
    """Analyze a Markdown document and generate modification suggestions.

    Performs semantic analysis across six categories:
    1. Heading hierarchy — skipped levels, missing/multiple H1, long titles
    2. Required 公文 sections — expected elements in Chinese government docs
    3. Format issues — broken image references, HTML table problems
    4. Content issues — empty sections, placeholder text, duplicates, long paragraphs
    5. Numbered list continuity — gaps and restarts in ordered lists
    6. Chinese punctuation — half-width marks, missing sentence endings

    Generates a structured report at <input_dir>/<stem>-修改建议.md with
    findings organized by category, each showing line numbers, severity,
    and specific suggestions.

    Args:
        input_path: Absolute path to the input Markdown file.

    Returns:
        Summary string with input path, suggestions file path, and issue count.
    """
    try:
        suggestions_path = analyze(input_path)
    except FileNotFoundError as e:
        return f"错误：文件不存在 — {e}"
    except UnicodeDecodeError:
        return f"错误：文件编码不是 UTF-8，无法读取 — {input_path}"
    except Exception as e:
        return f"分析失败: {type(e).__name__}: {e}"

    # Count issues from the generated report
    report = suggestions_path.read_text(encoding="utf-8")
    issue_count = report.count("| -- |") + report.count("| ")

    # Count lines starting with issue rows (markdown table rows after header)
    import re
    issue_lines = len(re.findall(r"^\| \d+ \|", report, re.MULTILINE))
    # Also count lines with "--" line numbers
    dash_lines = len(re.findall(r"^\| -- \|", report, re.MULTILINE))

    total = issue_lines + dash_lines

    return (
        f"✅ 分析完成\n"
        f"输入文件:     {input_path}\n"
        f"修改建议文件: {suggestions_path}\n"
        f"发现问题:     {total} 项\n"
        f"\n请查看 {suggestions_path.name} 了解详细修改建议。"
    )


# ── Entry point ───────────────────────────────────────────────────────


def main():
    """Run the MCP server over stdio transport."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
