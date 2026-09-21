#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
swapi -- offline search over the SolidWorks API.

Why this exists: `lookup_api_signature` answers "what are the parameters of
IFeatureManager.FeatureRevolve2" -- but only if you already know both the
interface and the exact member name. Out of ~9800 methods that is a hard
starting condition. The real question is usually the other way round:
"which method makes a shell?", "is there anything for drawings?".

This builds a flat, greppable index of the whole typelib and searches it.

SolidWorks does NOT have to be running -- only installed. Everything here
reads the makepy cache (see solidworks_mcp/utils/typelib.py), which is
generated from SolidWorks' own type library, i.e. ground truth for THIS
install rather than documentation or memory.

Usage:
    python scripts/swapi.py find shell            # search everything
    python scripts/swapi.py find drawing -k m     # methods only
    python scripts/swapi.py sig IFeatureManager FeatureRevolve2
    python scripts/swapi.py const swEndCondBlind
    python scripts/swapi.py build                 # force index rebuild
    python scripts/swapi.py stats

Index files land in api-index/ (generated, gitignored).
"""

import argparse
import inspect
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

INDEX_DIR = REPO / "api-index"
METHODS = INDEX_DIR / "methods.txt"
PROPS = INDEX_DIR / "properties.txt"
CONSTS = INDEX_DIR / "constants.txt"

KIND_FILES = {
    "m": ("method", METHODS),
    "p": ("property", PROPS),
    "c": ("constant", CONSTS),
}


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def _iter_interfaces(mod):
    """Yield (name, class) for generated COM interface wrappers only.

    The module also holds CoClasses and helper types; those carry neither a
    property map nor generated methods, so they get skipped.
    """
    for name in dir(mod):
        if name.startswith("_"):
            continue
        obj = getattr(mod, name, None)
        if not inspect.isclass(obj):
            continue
        own = vars(obj)
        has_props = "_prop_map_get_" in own or "_prop_map_put_" in own
        has_methods = any(callable(v) and not k.startswith("_")
                          for k, v in own.items())
        if has_props or has_methods:
            yield name, obj


def _params(fn):
    """Parameter names of a generated method, without self."""
    try:
        sig = inspect.signature(fn)
    except (ValueError, TypeError):
        return None
    return [p for p in sig.parameters if p != "self"]


def build(verbose=True):
    """Regenerate the index files from the makepy cache."""
    from solidworks_mcp.utils.typelib import get_sw_module, get_const_module

    INDEX_DIR.mkdir(exist_ok=True)

    mod = get_sw_module()
    method_lines, prop_lines = [], []

    for iface, cls in _iter_interfaces(mod):
        own = vars(cls)
        for member, val in own.items():
            if member.startswith("_") or not callable(val):
                continue
            # CoClass shim, present on every coclass and never callable API
            if member == "default_interface":
                continue
            names = _params(val)
            args = ", ".join(names) if names is not None else "?"
            method_lines.append("%s.%s(%s)" % (iface, member, args))

        getters = set(own.get("_prop_map_get_", {}))
        setters = set(own.get("_prop_map_put_", {}))
        for member in sorted(getters | setters):
            if member in getters and member in setters:
                mode = "get/put"
            elif member in getters:
                mode = "get"
            else:
                mode = "put"
            prop_lines.append("%s.%s  [%s]" % (iface, member, mode))

    constants = get_const_module().constants
    const_lines = []
    for name in dir(constants):
        if name.startswith("_"):
            continue
        value = getattr(constants, name)
        if isinstance(value, str):
            const_lines.append("%s = %r" % (name, value))
        elif isinstance(value, (int, float)):
            const_lines.append("%s = %s" % (name, value))

    written = []
    for path, lines in ((METHODS, method_lines),
                        (PROPS, prop_lines),
                        (CONSTS, const_lines)):
        lines.sort()
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        written.append((path, len(lines)))
        if verbose:
            size = path.stat().st_size / 1024
            print("  %-16s %6d lines  %7.0f KB" % (path.name, len(lines), size))

    return written


def _ensure_index():
    if not (METHODS.exists() and PROPS.exists() and CONSTS.exists()):
        print("index missing -- building (a few seconds, SolidWorks not needed)",
              file=sys.stderr)
        build(verbose=False)


# ---------------------------------------------------------------------------
# Search
# ---------------------------------------------------------------------------

def _member_of(line):
    """The member name alone from an index line.

    "IModelDoc2.InsertFeatureShell(Thickness, Outward)" -> "InsertFeatureShell"
    "swEndCondBlind = 0"                                -> "swEndCondBlind"
    """
    head = line.split("(", 1)[0].split("  [", 1)[0].split(" = ", 1)[0]
    return head.rsplit(".", 1)[-1]


def find(query, kinds="mpc", limit=40, regex=False):
    _ensure_index()
    pattern = re.compile(query if regex else re.escape(query), re.I)

    total = 0
    for key in kinds:
        if key not in KIND_FILES:
            continue
        label, path = KIND_FILES[key]
        hits = [ln for ln in path.read_text(encoding="utf-8").splitlines()
                if pattern.search(ln)]
        if not hits:
            continue
        # Rank by where the match landed. A hit in the member name is what
        # you are looking for; a hit in a parameter list is usually noise
        # (searching "shell" must not bury InsertFeatureShell under
        # GetShellType and every method taking a "shell" argument).
        hits.sort(key=lambda ln: (not pattern.search(_member_of(ln)), len(ln)))
        shown = hits[:limit]
        tail = ", showing %d" % len(shown) if len(hits) > len(shown) else ""
        print("\n=== %s (%d hits%s) ===" % (label, len(hits), tail))
        for ln in shown:
            print("  " + ln)
        total += len(hits)

    if not total:
        print("nothing matches %r" % query)
        return 1
    return 0


def sig(interface, member):
    from solidworks_mcp.utils.typelib import get_signature
    try:
        print(get_signature(interface, member))
    except AttributeError as exc:
        print(exc, file=sys.stderr)
        return 1
    return 0


def const(name):
    from solidworks_mcp.utils.typelib import get_constant
    try:
        print("%s = %s" % (name, get_constant(name)))
    except AttributeError as exc:
        print(exc, file=sys.stderr)
        return 1
    return 0


def stats():
    _ensure_index()
    for key in "mpc":
        label, path = KIND_FILES[key]
        n = sum(1 for _ in path.open(encoding="utf-8"))
        print("%-10s %6d   %s" % (label, n, path))
    ifaces = {ln.split(".", 1)[0]
              for ln in METHODS.read_text(encoding="utf-8").splitlines()}
    print("%-10s %6d" % ("interfaces", len(ifaces)))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="swapi",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = ap.add_subparsers(dest="cmd", required=True)

    f = sub.add_parser("find", help="search the index")
    f.add_argument("query")
    f.add_argument("-k", "--kinds", default="mpc",
                   help="which files to search: m=methods p=properties c=constants")
    f.add_argument("-n", "--limit", type=int, default=40, help="max hits per kind")
    f.add_argument("-e", "--regex", action="store_true", help="treat query as a regex")

    s = sub.add_parser("sig", help="full generated source of one member")
    s.add_argument("interface")
    s.add_argument("member")

    c = sub.add_parser("const", help="value of one constant")
    c.add_argument("name")

    sub.add_parser("build", help="rebuild the index")
    sub.add_parser("stats", help="index size")

    a = ap.parse_args(argv)
    if a.cmd == "find":
        return find(a.query, a.kinds, a.limit, a.regex)
    if a.cmd == "sig":
        return sig(a.interface, a.member)
    if a.cmd == "const":
        return const(a.name)
    if a.cmd == "build":
        build()
        return 0
    if a.cmd == "stats":
        return stats()
    return 1


if __name__ == "__main__":
    sys.exit(main())
