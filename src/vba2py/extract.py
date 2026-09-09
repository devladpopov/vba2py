"""Extract VBA source code from Office files (.xlsm, .xlsb, .xls, .docm, .xlam)."""

from __future__ import annotations

from pathlib import Path

OFFICE_EXTENSIONS = {".xlsm", ".xlsb", ".xls", ".xlam", ".docm", ".doc", ".xla"}


class ExtractionError(Exception):
    """Raised when VBA macros cannot be extracted from an Office file."""


def is_office_file(path: Path) -> bool:
    return path.suffix.lower() in OFFICE_EXTENSIONS


def extract_vba(path: Path) -> str:
    """Extract all VBA module sources from an Office file into one string.

    Modules are concatenated, each preceded by a `' Module: <name>` comment.
    Attribute lines (module metadata) are stripped since they are not
    executable VBA.
    """
    from oletools.olevba import VBA_Parser

    vba_parser = VBA_Parser(str(path))
    try:
        if not vba_parser.detect_vba_macros():
            raise ExtractionError(f"No VBA macros found in {path.name}")

        chunks: list[str] = []
        for _, _, module_name, code in vba_parser.extract_macros():
            code = _strip_attributes(code)
            if code.strip():
                chunks.append(f"' Module: {module_name}\n{code.strip()}")

        if not chunks:
            raise ExtractionError(f"VBA project in {path.name} contains no code")
        return "\n\n".join(chunks) + "\n"
    finally:
        vba_parser.close()


def _strip_attributes(code: str) -> str:
    lines = [
        line for line in code.splitlines()
        if not line.strip().startswith("Attribute VB_")
    ]
    return "\n".join(lines)
