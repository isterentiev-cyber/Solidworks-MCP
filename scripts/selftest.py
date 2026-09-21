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
