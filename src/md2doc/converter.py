"""Core conversion wrapper for Markdown-to-DOCX conversion.

Wraps the existing scripts/md2docx.py pipeline with exception-based error
handling suitable for MCP server use (no sys.exit calls).
"""

import importlib.util
import io
import os
import subprocess
import sys
from contextlib import redirect_stdout
from pathlib import Path


# ── Load core module from scripts/md2docx.py ──────────────────────────

def _load_core_module():
    """Load scripts/md2docx.py as a module via importlib.

    This avoids modifying the existing CLI script and isolates any
    module-level side effects.
    """
    core_path = Path(__file__).parent.parent.parent / "scripts" / "md2docx.py"
    if not core_path.exists():
        raise RuntimeError(
            f"Core script not found: {core_path}. "
            "Please ensure scripts/md2docx.py exists."
        )
    spec = importlib.util.spec_from_file_location("_md2docx_core", str(core_path))
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load module spec from {core_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules["_md2docx_core"] = module
    spec.loader.exec_module(module)
    return module


_core = _load_core_module()

# Re-export core functions and constants used by converter
DEFAULT_FORMAT_SPEC = _core.DEFAULT_FORMAT_SPEC
get_format_spec = _core.get_format_spec
create_reference_docx = _core.create_reference_docx
find_pandoc = _core.find_pandoc
extract_html_tables = _core.extract_html_tables
post_process_docx = _core.post_process_docx
_insert_html_tables = _core._insert_html_tables
_get_template_cache_dir = _core._get_template_cache_dir


# ── Exception hierarchy ───────────────────────────────────────────────

class ConversionError(RuntimeError):
    """Base exception for all conversion failures."""


class InputFileNotFoundError(ConversionError):
    """The input Markdown file does not exist."""


class PandocNotFoundError(ConversionError):
    """pandoc executable not found."""


class PandocConversionError(ConversionError):
    """pandoc returned a non-zero exit code."""


class PandocTimeoutError(ConversionError):
    """pandoc conversion exceeded the time limit."""


# ── Public API ────────────────────────────────────────────────────────

def convert(
    input_path: str | Path,
    output_path: str | Path | None = None,
    format_spec_md: str | Path | None = None,
    reference_docx: str | Path | None = None,
    extract_media: str | Path | None = None,
    skip_format: bool = False,
    no_cache_template: bool = False,
) -> Path:
    """Convert a Markdown file to Word (.docx) with formatting.

    Mirrors the pipeline of convert_md_to_docx() from scripts/md2docx.py
    but replaces all sys.exit(1) calls with raised exceptions.

    Args:
        input_path: Path to the input Markdown file.
        output_path: Path for the output .docx file. Auto-generated if None.
        format_spec_md: Path to a format spec markdown file (e.g. 格式要求.md).
            If None, uses the hardcoded DEFAULT_FORMAT_SPEC.
        reference_docx: Path to a .docx template to overlay on top of the
            format spec (pandoc --reference-doc).
        extract_media: Directory to extract media files to.
        skip_format: If True, bypass all formatting (pure pandoc output).
        no_cache_template: If True, force regeneration of the cached
            reference template.

    Returns:
        Path to the generated .docx file.

    Raises:
        InputFileNotFoundError: input_path does not exist.
        PandocNotFoundError: pandoc executable not found.
        PandocConversionError: pandoc returned non-zero exit code.
        PandocTimeoutError: conversion timed out.
        ConversionError: other conversion failures.
    """
    # Resolve all paths to absolute Path objects
    input_path = Path(input_path).resolve()
    if output_path is not None:
        output_path = Path(output_path).resolve()
    if format_spec_md is not None:
        format_spec_md = Path(format_spec_md).resolve()
    if reference_docx is not None:
        reference_docx = Path(reference_docx).resolve()
    if extract_media is not None:
        extract_media = Path(extract_media).resolve()

    # Validate input
    if not input_path.exists():
        raise InputFileNotFoundError(f"Input file not found: {input_path}")
    if not input_path.is_file():
        raise InputFileNotFoundError(f"Input path is not a file: {input_path}")

    # Load format spec
    format_spec = get_format_spec(str(format_spec_md) if format_spec_md else None)

    # Auto-generate output path
    if output_path is None:
        output_path = input_path.parent / f"{input_path.stem}.docx"

    # Ensure output directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # ── HTML table extraction ──
    original_md = input_path.read_text(encoding="utf-8")
    modified_md, html_tables = extract_html_tables(original_md)

    temp_md = None
    if html_tables:
        temp_md = input_path.with_suffix(".tmp.md")
        temp_md.write_text(modified_md, encoding="utf-8")
        actual_input = str(temp_md)
    else:
        actual_input = str(input_path)

    # ── Pandoc discovery ──
    pandoc_bin = find_pandoc()
    import shutil
    if not shutil.which(pandoc_bin):
        raise PandocNotFoundError(
            f"pandoc not found at '{pandoc_bin}' or in PATH. "
            "Install pandoc: https://pandoc.org/installing.html"
        )

    # ── Build pandoc command ──
    from_format = "markdown+pipe_tables+fenced_divs+bracketed_spans+raw_html"
    cmd = [
        pandoc_bin,
        actual_input,
        "-o", str(output_path),
        "--from", from_format,
        "--resource-path", str(Path(actual_input).parent),
        "--standalone",
    ]

    # ── Reference template ──
    use_builtin_ref = not skip_format
    if use_builtin_ref:
        cache_dir = _get_template_cache_dir()
        builtin_ref_path = os.path.join(cache_dir, "reference-template.docx")

        if no_cache_template or not os.path.exists(builtin_ref_path):
            create_reference_docx(builtin_ref_path, format_spec)
        cmd.extend(["--reference-doc", builtin_ref_path])

    # User-provided reference template (overlays on top)
    if reference_docx and reference_docx.exists():
        cmd.extend(["--reference-doc", str(reference_docx)])

    if extract_media:
        cmd.extend(["--extract-media", str(extract_media)])

    # ── Run pandoc ──
    # Suppress print() output from the core module during conversion
    try:
        with redirect_stdout(io.StringIO()):
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
            )
    except FileNotFoundError:
        raise PandocNotFoundError(
            f"pandoc not found at '{pandoc_bin}' or in PATH."
        )
    except subprocess.TimeoutExpired:
        raise PandocTimeoutError(
            "Conversion timed out after 120 seconds."
        )

    if result.returncode != 0:
        stderr = result.stderr.strip()
        raise PandocConversionError(
            f"pandoc exited with code {result.returncode}: {stderr}"
        )

    # ── Post-processing ──
    if not skip_format:
        with redirect_stdout(io.StringIO()):
            post_process_docx(str(output_path), format_spec)

    # ── Insert HTML tables ──
    if html_tables:
        with redirect_stdout(io.StringIO()):
            _insert_html_tables(str(output_path), html_tables, format_spec)

    # ── Cleanup temp file ──
    if temp_md is not None:
        try:
            os.unlink(temp_md)
        except OSError:
            pass

    return output_path
