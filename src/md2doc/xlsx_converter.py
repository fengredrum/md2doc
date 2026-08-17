"""Core wrapper for Markdown-to-Excel (xlsx) conversion.

Loads scripts/md2xlsx.py as a module and exposes an exception-based API
suitable for MCP server use (no sys.exit calls).
"""

import importlib.util
import sys
from pathlib import Path


def _load_core_module():
    """Load scripts/md2xlsx.py as a module via importlib."""
    core_path = Path(__file__).parent.parent.parent / "scripts" / "md2xlsx.py"
    if not core_path.exists():
        raise RuntimeError(
            f"Core script not found: {core_path}. "
            "Please ensure scripts/md2xlsx.py exists."
        )
    spec = importlib.util.spec_from_file_location("_md2xlsx_core", str(core_path))
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load module spec from {core_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules["_md2xlsx_core"] = module
    spec.loader.exec_module(module)
    return module


_core = _load_core_module()

extract_tables = _core.extract_tables
build_workbook = _core.build_workbook


# ── Exception hierarchy ───────────────────────────────────────────────

class XlsxConversionError(RuntimeError):
    """Base exception for all conversion failures."""


class InputFileNotFoundError(XlsxConversionError):
    """The input Markdown file does not exist."""


class NoTableFoundError(XlsxConversionError):
    """No HTML or Markdown pipe table was found in the input."""


# ── Public API ────────────────────────────────────────────────────────

def convert_markdown_to_xlsx(
    input_path: str | Path,
    output_path: str | Path | None = None,
    font_name: str | None = None,
    body_size: int | None = None,
) -> Path:
    """Convert Markdown tables to an Excel (.xlsx) file.

    Extracts every HTML <table> and Markdown pipe table from the document
    (in document order) and writes each to its own worksheet, styled with
    微软雅黑 font, thin borders, bold header (with light fill), vertical
    centering, wrapped text, and content-estimated column widths/row heights.
    colspan/rowspan merged cells, inline <b> bold, and text-align are preserved.

    Args:
        input_path: Path to the input Markdown file.
        output_path: Path for the output .xlsx file. Auto-generated if None
            (same directory and stem as the input).
        font_name: Optional font name override (default 微软雅黑).
        body_size: Optional body font size override (default 11).

    Returns:
        Path to the generated .xlsx file.

    Raises:
        InputFileNotFoundError: input_path does not exist.
        NoTableFoundError: no table was found in the document.
        XlsxConversionError: other conversion failures.
    """
    input_path = Path(input_path).resolve()
    if output_path is not None:
        output_path = Path(output_path).resolve()

    if not input_path.exists():
        raise InputFileNotFoundError(f"Input file not found: {input_path}")
    if not input_path.is_file():
        raise InputFileNotFoundError(f"Input path is not a file: {input_path}")

    if output_path is None:
        output_path = input_path.parent / f"{input_path.stem}.xlsx"

    tables = extract_tables(input_path.read_text(encoding="utf-8"))
    if not tables:
        raise NoTableFoundError(
            "No table found in the document (neither HTML <table> nor "
            "Markdown pipe table)."
        )

    kwargs = {}
    if font_name:
        kwargs["font_name"] = font_name
    if body_size:
        kwargs["body_size"] = body_size

    wb = build_workbook(tables, **kwargs)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
