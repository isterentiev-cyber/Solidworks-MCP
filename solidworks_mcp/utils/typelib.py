"""
SolidWorks Typelib Introspection
---------------------------------
Look up the REAL signature of a SolidWorks COM method (exact parameter
count, order, and names) instead of guessing from memory or from official
docs that don't match this install. See CLAUDE.md "Не угадывай сигнатуры"
for why this matters -- several bugs this project shipped (FeatureRevolve2,
InsertFeatureChamfer) were exactly "the doc/memory said N params, the real
COM interface wants a different N".

How it works: generates a win32com "makepy" cache from SolidWorks' own
type library (a one-time, ~1-2s operation per Python installation --
cached under %TEMP%\\gen_py after that, survives SW restarts but not a SW
reinstall/upgrade or a Python version change). Once generated, the
resulting Python module has one class per COM interface, and each
generated method is a real Python function whose source literally shows
the parameter list and the underlying InvokeTypes() call -- ground truth,
not documentation.

Typical use (or just call the `lookup_api_signature` / `lookup_api_constant`
MCP tools instead of writing this by hand):

    from solidworks_mcp.utils.typelib import get_signature, get_constant
    print(get_signature("IFeatureManager", "FeatureRevolve2"))
    print(get_constant("swRefPlaneReferenceConstraint_Distance"))
"""

import os
import sys
import inspect
import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Where the generated makepy cache lives
# ---------------------------------------------------------------------------
# pywin32 defaults to %TEMP%\gen_py\<pyver>, which any disk cleanup wipes and
# which a Python version bump orphans. The cache is ~7 MB and takes a couple of
# seconds to rebuild, but it rebuilds *silently and at the worst moment* -- the
# first lookup_api_signature of a session. Pin it inside the repo instead.
#
# Override with SW_MCP_GEN_PY if you want it somewhere else.
# The directory is generated, machine-specific and gitignored -- never commit it.

_GEN_PATH_DONE = False


def _stable_gen_path() -> str:
    """Repo-local makepy cache directory (absolute)."""
    override = os.environ.get("SW_MCP_GEN_PY", "").strip()
    if override:
        return os.path.abspath(override)
    return str(Path(__file__).resolve().parents[2] / ".gen_py")


def _relocate_gen_cache() -> str:
    """Point win32com's generated-code cache at a stable directory.

    Must run before gencache generates or loads anything. gencache reads
    win32com.__gen_path__ on every call (not via `from ... import`), so
    reassigning the attribute is enough -- except for its dicts.dat index,
    which it loaded at import time from the OLD path. That one gets reloaded
    here explicitly.
    """
    global _GEN_PATH_DONE
    if _GEN_PATH_DONE:
        return _stable_gen_path()

    import win32com

    target = _stable_gen_path()
    current = getattr(win32com, "__gen_path__", "")
    if os.path.normcase(os.path.abspath(current)) == os.path.normcase(target):
        _GEN_PATH_DONE = True
        return target

    os.makedirs(target, exist_ok=True)
    win32com.__gen_path__ = target
    gen_py_mod = sys.modules.get("win32com.gen_py")
    if gen_py_mod is not None:
        gen_py_mod.__path__ = [target]

    from win32com.client import gencache

    gencache.clsidToTypelib.clear()
    gencache.versionRedirectMap.clear()
    try:
        gencache._LoadDicts()
    except OSError:
        # No dicts.dat in the fresh directory yet -- build one. verbose=0 is
        # not optional: this process speaks MCP over stdout, and Rebuild()
        # prints progress there by default, which corrupts the protocol.
        gencache.Rebuild(verbose=0)

    logger.info(f"makepy cache: {current or '(default)'} -> {target}")
    _GEN_PATH_DONE = True
    return target

# Typelib descriptions to search for (see selecttlb.EnumTlbs()). Matched by
# substring so this doesn't need updating every SW version bump (e.g. covers
# "SldWorks 2026 Type Library", "SldWorks 2027 Type Library", ...).
_SW_TYPELIB_NAME_HINTS = ("sldworks", "type library")
_CONST_TYPELIB_NAME_HINTS = ("solidworks", "constant type library")

_sw_module = None
_const_module = None


def _find_typelib(name_hints):
    """Find a registered typelib whose description contains all of
    name_hints (case-insensitive). Returns the selecttlb TypelibSpec."""
    import win32com.client.selecttlb as selecttlb

    for t in selecttlb.EnumTlbs():
        desc = (t.desc or "").lower()
        if all(hint in desc for hint in name_hints):
            return t
    raise RuntimeError(
        f"No registered typelib matches hints {name_hints}. "
        f"Is SolidWorks actually installed on this machine?"
    )


def _ensure_module(name_hints, cache_attr):
    """Generate (if needed) and return the makepy module for a typelib
    matching name_hints, memoized in this process by cache_attr."""
    global _sw_module, _const_module
    cached = globals()[cache_attr]
    if cached is not None:
        return cached

    _relocate_gen_cache()

    from win32com.client import makepy
    import win32com.client.gencache as gencache

    target = _find_typelib(name_hints)
    makepy.GenerateFromTypeLibSpec(target)
    mod = gencache.EnsureModule(
        target.clsid, target.lcid, int(target.major), int(target.minor)
    )
    globals()[cache_attr] = mod
    logger.info(f"Loaded typelib module: {target.desc}")
    return mod


def get_sw_module():
    """The main SolidWorks API typelib module (IFeatureManager, IModelDoc2,
    ISketchManager, IBody2, ... one class per COM interface)."""
    return _ensure_module(_SW_TYPELIB_NAME_HINTS, "_sw_module")


def get_const_module():
    """The SolidWorks constants typelib module -- values live under
    `.constants`, e.g. get_const_module().constants.swEndCondBlind."""
    return _ensure_module(_CONST_TYPELIB_NAME_HINTS, "_const_module")


def get_signature(interface: str, member: str) -> str:
    """
    Return the real Python source of a generated COM method/property,
    which shows its exact parameter list and dispid. Raises AttributeError
    with a helpful message if the interface or member doesn't exist.

    Args:
        interface: COM interface name, e.g. "IFeatureManager", "IModelDoc2"
        member: method or property name, e.g. "FeatureRevolve2"
    """
    mod = get_sw_module()
    cls = getattr(mod, interface, None)
    if cls is None:
        raise AttributeError(
            f"No interface '{interface}' in the SolidWorks typelib. "
            f"Check spelling/capitalization (COM interfaces are case-sensitive here)."
        )
    member_obj = getattr(cls, member, None)
    if member_obj is not None and callable(member_obj):
        return inspect.getsource(member_obj)

    # Not a generated method -- check if it's a property instead (properties
    # don't show up via getattr as functions; they're in _prop_map_get_).
    get_props = getattr(cls, "_prop_map_get_", {})
    put_props = getattr(cls, "_prop_map_put_", {})
    if member in get_props or member in put_props:
        return (
            f"{interface}.{member} is a PROPERTY, not a method "
            f"(get entry: {get_props.get(member)!r}, "
            f"put entry: {put_props.get(member)!r}). "
            f"Access it as `obj.{member}` (read) or `obj.{member} = value` "
            f"(write), never with parentheses."
        )

    available = sorted(set(get_props) | set(put_props) |
                        {m for m in dir(cls) if not m.startswith("_")})
    raise AttributeError(
        f"No member '{member}' on {interface}. Closest available members: "
        f"{[m for m in available if member.lower()[:4] in m.lower()][:10] or available[:20]}"
    )


def get_constant(name: str) -> int:
    """Return the integer value of a SolidWorks API constant, e.g.
    get_constant('swEndCondBlind') -> 0. Raises AttributeError if unknown --
    the error message suggests near-matches so you can catch typos."""
    mod = get_const_module()
    try:
        return getattr(mod.constants, name)
    except AttributeError:
        candidates = [n for n in dir(mod.constants) if name.lower() in n.lower()]
        raise AttributeError(
            f"No constant '{name}'. Closest matches: {candidates[:10]}"
        )
