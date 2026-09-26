#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Is the code on disk fit to be served?

Run this after editing the server and before expecting the client to respawn
it. A stdio MCP server that dies on import fails *silently* from the client's
point of view -- you get "no such tool" and no traceback anywhere you would
think to look. This turns that into a readable answer in one second.

SolidWorks is not needed and is never contacted: nothing here connects, opens
a document, or touches COM beyond reading the type library.

    python scripts/selftest.py

Exit code 0 = servable, 1 = something is wrong.
"""

import asyncio
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))


def _check_view_sketch_math(ext):
    """Sheet mm <-> drawing-view sketch coordinates (add_section_view /
    add_detail_view). The ArrayData below is the view sketch's
    ModelToSketchTransform read live on 2026-09-22 from a 1:5 Front view
    placed at sheet (150,200) mm: sketch = (sheet - Position) / scale."""
    bad = []
    front = [1, 0, 0, 0, 1, 0, 0, 0, 1, -0.75, -1.0, 0.0, 5.0, 0, 0, 0]
    cases = [((150, 200), (0.0, 0.0)),        # view centre -> sketch origin
             ((150, 340), (0.0, 0.7)),        # +140 sheet mm -> +700 model mm
             ((119, 200), (-0.155, 0.0))]
    for sheet, want in cases:
        got = ext.sheet_to_view_sketch(front, *sheet)
        if max(abs(got[0] - want[0]), abs(got[1] - want[1])) > 1e-9:
            bad.append(f"sheet_to_view_sketch{sheet} = {got}, want {want}")
    # a rotated (30 deg) and scaled view must round-trip too
    c, s = 0.8660254037844387, 0.5
    rot = [c, s, 0, -s, c, 0, 0, 0, 1, 0.12, -0.34, 0.0, 2.0, 0, 0, 0]
    for sheet in ((0, 0), (150, 200), (-40.5, 610)):
        sk = ext.sheet_to_view_sketch(rot, *sheet)
        back = ext.view_sketch_to_sheet(rot, *sk)
        if max(abs(back[0] - sheet[0]), abs(back[1] - sheet[1])) > 1e-9:
            bad.append(f"view sketch round trip {sheet} -> {sk} -> {back}")
    return bad


_FORBIDDEN = ("gencache", "EnsureDispatch", "EnsureModule", "makepy",
              "win32com.client.Dispatch", "win32com.client.GetObject",
              "win32com.client.GetActiveObject", "CastTo")


def _check_late_bound_only():
    """The server is late-bound only (utils/com_helpers.py). Two guards:
    no forbidden API in code (comments/strings ignored), and no makepy module
    for a SolidWorks typelib imported by loading the server."""
    import io
    import tokenize
    bad = []
    for f in sorted((REPO / "solidworks_mcp").rglob("*.py")):
        if "backup_" in str(f):
            continue
        toks = list(tokenize.generate_tokens(io.StringIO(f.read_text(encoding="utf-8")).readline))
        code = [t for t in toks if t.type == tokenize.NAME or (t.type == tokenize.OP and t.string == ".")]
        for i, t in enumerate(code):
            if t.type != tokenize.NAME:
                continue
            dotted = t.string
            j = i
            while j + 2 < len(code) and code[j + 1].string == "." and code[j + 2].type == tokenize.NAME:
                dotted += "." + code[j + 2].string
                j += 2
            for word in _FORBIDDEN:
                if dotted == word or dotted.endswith("." + word) or dotted.startswith(word + "."):
                    bad.append(f"{f.relative_to(REPO)}:{t.start[0]}: forbidden '{word}' (late-bound only)")
    typed = [m for m in sys.modules if m.startswith("win32com.gen_py.")]
    if typed:
        bad.append(f"makepy modules imported by the server: {typed}")
    return sorted(set(bad))


def main():
    problems = []
    notes = []

    try:
        from solidworks_mcp import server, ext, toolsets
    except Exception:
        import traceback
        print("IMPORT FAILED -- the client cannot start this server:\n")
        traceback.print_exc()
        return 1

    try:
        tools = asyncio.run(server.list_tools())
    except Exception:
        import traceback
        print("list_tools() FAILED -- the client would see an empty tool list:\n")
        traceback.print_exc()
        return 1

    names = [t.name for t in tools]
    dupes = sorted({n for n in names if names.count(n) > 1})
    if dupes:
        problems.append(f"duplicate tool names: {dupes}")

    # Every ext schema needs a handler behind it, or the tool exists in the
    # list and dies on call. reload_api is handled in server.py, not ext.
    orphan_schemas = [n for n, _, _ in ext.TOOL_SCHEMAS
                      if n not in ext.HANDLERS and n != "reload_api"]
    if orphan_schemas:
        problems.append(f"schema without a handler (would fail on call): {orphan_schemas}")

    orphan_handlers = [n for n in ext.HANDLERS
                       if n not in {s[0] for s in ext.TOOL_SCHEMAS}]
    if orphan_handlers:
        notes.append(f"handler with no schema (unreachable as a tool): {orphan_handlers}")

    classified = set().union(*toolsets.TOOLSETS.values())
    unclassified = sorted(set(names) - classified)
    if unclassified:
        notes.append(f"not in toolsets.py, always visible: {unclassified}")
    stale = sorted(classified - set(names))
    if stale:
        notes.append(f"in toolsets.py but no such tool: {stale}")

    for t in tools:
        if not t.description:
            problems.append(f"{t.name}: empty description (the model reads this)")
        schema = t.inputSchema or {}
        for req in schema.get("required", []):
            if req not in (schema.get("properties") or {}):
                problems.append(f"{t.name}: required '{req}' is not among its properties")

    problems += _check_view_sketch_math(ext)
    problems += _check_late_bound_only()

    print(f"tools served: {len(tools)}")
    print(f"  ext handlers: {len(ext.HANDLERS)}   ext schemas: {len(ext.TOOL_SCHEMAS)}")
    print(f"  toolsets: " + ", ".join(
        f"{k}={len(v)}" for k, v in sorted(toolsets.TOOLSETS.items())))

    for n in notes:
        print(f"  note: {n}")
    if problems:
        print("\nPROBLEMS:")
        for p in problems:
            print(f"  - {p}")
        return 1

    print("\nOK -- servable.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
