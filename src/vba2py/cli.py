"""Command-line interface for vba2py."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from vba2py.parser import parse
from vba2py.codegen import generate


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        prog="vba2py",
        description="Convert VBA source code to Python.",
    )
    ap.add_argument("input", help="VBA source file (.bas / .vba / .txt) or Office file (.xlsm / .xls / .docm)")
    ap.add_argument("-o", "--output", help="Output Python file (default: stdout)")
    ap.add_argument("--ast", action="store_true", help="Print AST instead of Python code")
    args = ap.parse_args(argv)

    src_path = Path(args.input)
    if not src_path.exists():
        print(f"Error: file not found: {src_path}", file=sys.stderr)
        return 1

    from vba2py.extract import ExtractionError, extract_vba, is_office_file

    if is_office_file(src_path):
        try:
            source = extract_vba(src_path)
        except ExtractionError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 1
    else:
        source = src_path.read_text(encoding="utf-8-sig")

    try:
        module = parse(source)
    except Exception as exc:
        print(f"Parse error: {exc}", file=sys.stderr)
        return 1

    if args.ast:
        output = _format_ast(module)
    else:
        try:
            output = generate(module)
        except Exception as exc:
            print(f"Code generation error: {exc}", file=sys.stderr)
            return 1

    if args.output:
        Path(args.output).write_text(output, encoding="utf-8")
    else:
        print(output, end="")

    return 0


def _format_ast(node, indent: int = 0) -> str:
    """Pretty-print an AST node tree."""
    from dataclasses import fields, is_dataclass
    lines: list[str] = []
    prefix = "  " * indent
    name = type(node).__name__

    if is_dataclass(node):
        lines.append(f"{prefix}{name}(")
        for f in fields(node):
            if f.name == "lineno":
                continue
            val = getattr(node, f.name)
            if isinstance(val, list):
                if not val:
                    lines.append(f"{prefix}  {f.name}=[]")
                else:
                    lines.append(f"{prefix}  {f.name}=[")
                    for item in val:
                        if is_dataclass(item):
                            lines.append(_format_ast(item, indent + 2))
                        elif isinstance(item, tuple):
                            lines.append(f"{prefix}    (")
                            for t in item:
                                if is_dataclass(t):
                                    lines.append(_format_ast(t, indent + 3))
                                elif isinstance(t, list):
                                    for li in t:
                                        lines.append(_format_ast(li, indent + 3))
                                else:
                                    lines.append(f"{prefix}      {t!r}")
                            lines.append(f"{prefix}    )")
                        else:
                            lines.append(f"{prefix}    {val!r}")
                    lines.append(f"{prefix}  ]")
            elif is_dataclass(val):
                lines.append(f"{prefix}  {f.name}=")
                lines.append(_format_ast(val, indent + 2))
            else:
                lines.append(f"{prefix}  {f.name}={val!r}")
        lines.append(f"{prefix})")
    else:
        lines.append(f"{prefix}{node!r}")

    return "\n".join(lines)


if __name__ == "__main__":
    sys.exit(main())
