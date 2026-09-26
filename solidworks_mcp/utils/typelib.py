"""
SolidWorks Typelib Introspection (no makepy, no gencache)
---------------------------------------------------------
Look up the REAL signature of a SolidWorks COM method (exact parameter
count, order, names, in/out flags) instead of guessing from memory or from
official docs that don't match this install. See CLAUDE.md "Не угадывай
сигнатуры" -- FeatureRevolve2 / InsertFeatureChamfer shipped broken exactly
because "the doc said N params, the COM interface wants a different N".

How it works: the registered type library is opened directly with
`pythoncom.LoadRegTypeLib` and read through ITypeLib / ITypeInfo. Nothing is
generated on disk and nothing is imported into win32com.gen_py -- this is on
purpose. The server is late-bound only (`dynamic.Dispatch`): once a makepy
module for SldWorks is imported into the process, `win32com.client.Dispatch`
and `GetObject` start handing out typed wrappers whose out-parameter and
property/method conventions differ, and the same tool then behaves
differently from session to session. gencache / EnsureDispatch / makepy must
not be used anywhere in the server or in execute_python code.

Typical use (or call the `lookup_api_signature` / `lookup_api_constant` tools):

    from solidworks_mcp.utils.typelib import get_signature, get_constant
    print(get_signature("IFeatureManager", "FeatureRevolve2"))
    print(get_constant("swRefPlaneReferenceConstraint_Distance"))
"""

import logging
import threading
from typing import Dict, List, Optional, Tuple

import pythoncom

logger = logging.getLogger(__name__)

# Typelib descriptions to search for (selecttlb.EnumTlbs()). Matched by
# substring so it survives version bumps ("SldWorks 2026 Type Library", ...).
_SW_TYPELIB_NAME_HINTS = ("sldworks", "type library")
_CONST_TYPELIB_NAME_HINTS = ("solidworks", "constant type library")

_lock = threading.RLock()
_libs: Dict[str, object] = {}          # "sw" / "const" -> ITypeLib
_iface_index: Optional[Dict[str, int]] = None
_const_values: Optional[Dict[str, object]] = None

# VARTYPE -> readable name
_VT_NAMES = {
    pythoncom.VT_EMPTY: "void", pythoncom.VT_VOID: "void", pythoncom.VT_HRESULT: "HRESULT",
    pythoncom.VT_I2: "short", pythoncom.VT_I4: "long", pythoncom.VT_INT: "int",
    pythoncom.VT_UI1: "byte", pythoncom.VT_UI2: "ushort", pythoncom.VT_UI4: "ulong",
    pythoncom.VT_UINT: "uint", pythoncom.VT_I8: "longlong", pythoncom.VT_UI8: "ulonglong",
    pythoncom.VT_R4: "float", pythoncom.VT_R8: "double", pythoncom.VT_BOOL: "bool",
    pythoncom.VT_BSTR: "str", pythoncom.VT_DISPATCH: "IDispatch", pythoncom.VT_UNKNOWN: "IUnknown",
    pythoncom.VT_VARIANT: "VARIANT", pythoncom.VT_DATE: "date", pythoncom.VT_CY: "currency",
    pythoncom.VT_LPSTR: "char*", pythoncom.VT_LPWSTR: "wchar*",
}
_INVKIND = {pythoncom.INVOKE_FUNC: "method", pythoncom.INVOKE_PROPERTYGET: "get",
            pythoncom.INVOKE_PROPERTYPUT: "put", pythoncom.INVOKE_PROPERTYPUTREF: "putref"}
_PARAMFLAG_FIN, _PARAMFLAG_FOUT, _PARAMFLAG_FRETVAL, _PARAMFLAG_FOPT = 0x1, 0x2, 0x8, 0x10
_FUNCFLAG_FRESTRICTED = 0x1


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

def _find_typelib(name_hints):
    """Registered typelib whose description contains all hints; highest
    version wins if several SW versions are registered."""
    import win32com.client.selecttlb as selecttlb  # registry scan, no codegen

    found = [t for t in selecttlb.EnumTlbs()
             if all(h in (t.desc or "").lower() for h in name_hints)]
    if not found:
        raise RuntimeError(f"No registered typelib matches {name_hints}. "
                           f"Is SolidWorks installed on this machine?")

    def ver(t):
        try:
            return (int(t.major, 16) if isinstance(t.major, str) else int(t.major),
                    int(t.minor, 16) if isinstance(t.minor, str) else int(t.minor))
        except (TypeError, ValueError):
            return (0, 0)
    return max(found, key=ver)


def _load(key: str, hints) -> object:
    lib = _libs.get(key)
    if lib is not None:
        return lib
    with _lock:
        lib = _libs.get(key)
        if lib is None:
            spec = _find_typelib(hints)
            major = int(spec.major, 16) if isinstance(spec.major, str) else int(spec.major)
            minor = int(spec.minor, 16) if isinstance(spec.minor, str) else int(spec.minor)
            lib = pythoncom.LoadRegTypeLib(spec.clsid, major, minor, spec.lcid or 0)
            _libs[key] = lib
            logger.info(f"Loaded typelib: {spec.desc}")
    return lib


def sw_typelib():
    """ITypeLib of the main SolidWorks API (IModelDoc2, IFeatureManager, ...)."""
    return _load("sw", _SW_TYPELIB_NAME_HINTS)


def const_typelib():
    """ITypeLib of the SolidWorks constants (sw*_e enums)."""
    return _load("const", _CONST_TYPELIB_NAME_HINTS)


def _index() -> Dict[str, int]:
    global _iface_index
    if _iface_index is None:
        lib = sw_typelib()
        _iface_index = {lib.GetDocumentation(i)[0]: i for i in range(lib.GetTypeInfoCount())}
    return _iface_index


def interface_info(interface: str):
    """ITypeInfo for an interface name, dispatch view for dual interfaces
    (so inherited IDispatch plumbing is flagged restricted and skipped)."""
    idx = _index().get(interface)
    if idx is None:
        raise AttributeError(
            f"No interface '{interface}' in the SolidWorks typelib. Check spelling/"
            f"capitalization. Similar: {[n for n in _index() if interface.lower()[1:6] in n.lower()][:10]}")
    ti = sw_typelib().GetTypeInfo(idx)
    attr = ti.GetTypeAttr()
    if attr.typekind == pythoncom.TKIND_INTERFACE and attr.wTypeFlags & pythoncom.TYPEFLAG_FDUAL:
        try:
            ti = ti.GetRefTypeInfo(ti.GetRefTypeOfImplType(-1))
        except pythoncom.com_error:
            pass
    return ti


def get_iid(interface: str):
    """IID of an interface (for QueryInterface in ext.T)."""
    return interface_info(interface).GetTypeAttr().iid


# ---------------------------------------------------------------------------
# Members
# ---------------------------------------------------------------------------

def _type_name(ti, tdesc) -> str:
    if isinstance(tdesc, int):
        base = tdesc & ~(pythoncom.VT_BYREF | pythoncom.VT_ARRAY)
        name = _VT_NAMES.get(base, f"vt{base}")
        if tdesc & pythoncom.VT_ARRAY:
            name += "[]"
        if tdesc & pythoncom.VT_BYREF:
            name += "*"
        return name
    vt, inner = tdesc[0], tdesc[1]
    if vt == pythoncom.VT_PTR:
        return _type_name(ti, inner) + "*"
    if vt == pythoncom.VT_SAFEARRAY:
        return _type_name(ti, inner) + "[]"
    if vt == pythoncom.VT_CARRAY:
        return _type_name(ti, inner[0] if isinstance(inner, tuple) else inner) + "[n]"
    if vt == pythoncom.VT_USERDEFINED:
        try:
            return ti.GetRefTypeInfo(inner).GetDocumentation(-1)[0]
        except pythoncom.com_error:
            return "userdefined"
    return f"vt{vt}"


def members(interface: str) -> List[dict]:
    """Every API member of an interface: name, kind (method/get/put/putref),
    dispid, params [(name, type, flags)], return type. Restricted IDispatch /
    IUnknown plumbing is dropped."""
    ti = interface_info(interface)
    attr = ti.GetTypeAttr()
    out = []
    for k in range(attr.cFuncs):
        fd = ti.GetFuncDesc(k)
        if fd.wFuncFlags & _FUNCFLAG_FRESTRICTED:
            continue
        names = ti.GetNames(fd.memid)
        params = []
        for j, (tdesc, flags, _default) in enumerate(fd.args):
            pname = names[j + 1] if j + 1 < len(names) else f"arg{j}"
            if fd.invkind in (pythoncom.INVOKE_PROPERTYPUT, pythoncom.INVOKE_PROPERTYPUTREF) \
                    and j == len(fd.args) - 1 and j + 1 >= len(names):
                pname = "value"
            params.append((pname, _type_name(ti, tdesc), flags))
        out.append({"name": names[0], "kind": _INVKIND.get(fd.invkind, str(fd.invkind)),
                    "dispid": fd.memid, "params": params,
                    "ret": _type_name(ti, fd.rettype[0]).rstrip("*")})
    return out


def out_params(interface: str, member: str) -> List[str]:
    """Names of [out] (by-ref) parameters of a method -- in late-bound calls
    each of them needs a com_helpers.out_*() VARIANT."""
    return [p[0] for m in members(interface) if m["name"] == member
            for p in m["params"] if p[2] & _PARAMFLAG_FOUT and not p[2] & _PARAMFLAG_FRETVAL]


def _flag_text(flags: int) -> str:
    parts = []
    if flags & _PARAMFLAG_FOUT:
        parts.append("out, byref")
    elif flags & _PARAMFLAG_FIN or not flags:
        parts.append("in")
    if flags & _PARAMFLAG_FOPT:
        parts.append("optional")
    return ", ".join(parts)


def get_signature(interface: str, member: str) -> str:
    """Readable signature of a method or property from the typelib, plus
    how to call it late-bound (out params -> com.out_*(), null objects ->
    com.nothing())."""
    ms = [m for m in members(interface) if m["name"] == member]
    if not ms:
        avail = sorted({m["name"] for m in members(interface)})
        near = [n for n in avail if member.lower()[:4] in n.lower()]
        raise AttributeError(f"No member '{member}' on {interface}. "
                             f"Closest: {near[:10] or avail[:20]}")
    lines = []
    methods = [m for m in ms if m["kind"] == "method"]
    props = [m for m in ms if m["kind"] != "method"]
    for m in methods:
        lines.append(f"{interface}.{member}  [method, dispid {m['dispid']}]")
        call_args = []
        for pname, ptype, flags in m["params"]:
            lines.append(f"  {pname}: {ptype}  ({_flag_text(flags)})")
            if flags & _PARAMFLAG_FOUT:
                call_args.append(_out_hint(ptype))
            elif ptype in ("IDispatch", "IUnknown") or ptype.startswith("I") and not ptype.endswith("[]"):
                call_args.append(f"{pname} | com.nothing()")
            else:
                call_args.append(pname)
        lines.append(f"  -> {m['ret']}")
        if not m["params"]:
            lines.append(f"Late-bound: zero-arg method -- read it with v(obj, '{member}') "
                         f"(obj.{member}() breaks: attribute access already invokes it).")
        else:
            lines.append(f"Late-bound: obj.{member}({', '.join(call_args)})")
            if any(p[2] & _PARAMFLAG_FOUT for p in m["params"]):
                lines.append("  [out] params: pass the VARIANT, read .value after the call "
                             "(or ext.com.call_out(obj, name, ...) -> (ret, *outs)).")
    if props:
        kinds = sorted({m["kind"] for m in props})
        t = next((m["ret"] for m in props if m["kind"] == "get"),
                 props[0]["params"][-1][1] if props[0]["params"] else "?")
        idx = [p for p in props[0]["params"] if p[0] != "value"]
        lines.append(f"{interface}.{member}  [PROPERTY {'/'.join(kinds)}: {t}]"
                     + (f"  indexed by {', '.join(p[0] + ': ' + p[1] for p in idx)}" if idx else ""))
        lines.append(f"Access as obj.{member} (read) / obj.{member} = value (write), no parentheses.")
    return "\n".join(lines)


def _out_hint(ptype: str) -> str:
    base = ptype.rstrip("*")
    return {"long": "com.out_int()", "int": "com.out_int()", "short": "com.out_int()",
            "bool": "com.out_bool()", "double": "com.out_float()", "float": "com.out_float()",
            "str": "com.out_str()", "VARIANT": "com.out_variant()"}.get(base, "com.out_dispatch()")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

def constants() -> Dict[str, object]:
    """All sw* enum values from the constants typelib, name -> value."""
    global _const_values
    if _const_values is None:
        with _lock:
            if _const_values is None:
                lib = const_typelib()
                vals = {}
                for i in range(lib.GetTypeInfoCount()):
                    ti = lib.GetTypeInfo(i)
                    attr = ti.GetTypeAttr()
                    if attr.typekind not in (pythoncom.TKIND_ENUM, pythoncom.TKIND_MODULE):
                        continue
                    for j in range(attr.cVars):
                        vd = ti.GetVarDesc(j)
                        if vd.varkind != pythoncom.VAR_CONST:
                            continue
                        vals[ti.GetNames(vd.memid)[0]] = vd.value
                _const_values = vals
    return _const_values


def get_constant(name: str):
    """Integer value of a SolidWorks API constant, e.g. 'swEndCondBlind' -> 0.
    Raises AttributeError with near-matches on a typo."""
    vals = constants()
    try:
        return vals[name]
    except KeyError:
        cand = [n for n in vals if name.lower() in n.lower()]
        raise AttributeError(f"No constant '{name}'. Closest matches: {cand[:10]}")


_enum_cache: Dict[str, Dict[str, int]] = {}


def enum_values(enum_name: str) -> Dict[str, int]:
    """Members of one enum from the constants typelib, e.g.
    enum_values('swFileLoadError_e') -> {'swGenericError': 1, ...}."""
    vals = _enum_cache.get(enum_name)
    if vals is None:
        lib = const_typelib()
        vals = {}
        for i in range(lib.GetTypeInfoCount()):
            if lib.GetDocumentation(i)[0] != enum_name:
                continue
            ti = lib.GetTypeInfo(i)
            for j in range(ti.GetTypeAttr().cVars):
                vd = ti.GetVarDesc(j)
                vals[ti.GetNames(vd.memid)[0]] = vd.value
            break
        _enum_cache[enum_name] = vals
    return vals


def enum_flags(enum_name: str, value) -> List[str]:
    """Names of the bits set in `value` for a bitmask enum (file load/save
    errors and warnings). Exact match first; unknown bits are reported as
    numbers. [] for 0 / None."""
    try:
        value = int(value or 0)
    except (TypeError, ValueError):
        return [str(value)]
    if not value:
        return []
    try:
        vals = enum_values(enum_name)
    except Exception:
        return [str(value)]
    exact = [n for n, x in vals.items() if x == value]
    if exact:
        return exact[:1]
    names, rest = [], value
    for n, x in sorted(vals.items(), key=lambda kv: kv[1]):
        if x and x & (x - 1) == 0 and value & x:
            names.append(n)
            rest &= ~x
    if rest:
        names.append(str(rest))
    return names


def interfaces() -> List[str]:
    """All type names in the main typelib (for the offline index)."""
    return list(_index())
