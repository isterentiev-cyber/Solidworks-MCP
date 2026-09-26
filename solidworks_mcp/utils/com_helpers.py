"""
SolidWorks COM Access Helpers (late-bound only)
-----------------------------------------------
The whole server talks to SolidWorks through `win32com.client.dynamic`
(late binding). Never `win32com.client.Dispatch`, `GetObject`,
`GetActiveObject`, `gencache.EnsureDispatch` or makepy: once a makepy module
for SldWorks is imported into the process, those hand out *typed* wrappers,
and typed wrappers have different conventions for out-parameters (tuple
returns instead of by-ref VARIANTs) and for zero-arg methods (call with `()`
instead of plain attribute read). The same tool then worked in one session
and failed in the next (open_document, 2026-09-26).

SolidWorks' IDispatch does not expose type info (GetTypeInfo -> "Element not
found"), so a dynamic wrapper never knows whether a member is a property or
a method. Rules that follow from that:

- Zero-arg member (property OR method): read it -- `v(obj, "GetTitle")` or
  `obj.GetTitle`. Attribute access already invokes it; `obj.GetTitle()` then
  tries to call the returned str/None and fails (or runs the method twice).
- Member with args: `v(obj, "Name", a, b)` or `obj.Name(a, b)`.
- [out] / by-ref parameter: pass `out_int()` / `out_bool()` / ... and read
  `.value` after the call, or use `call_out()` which returns (ret, *outs).
- Null object argument (Callout, Data, ...): `nothing()` -- a plain `None`
  arrives as VT_EMPTY and SolidWorks answers "Type mismatch".
- Arrays of doubles for properties like IView.Position: `double_array()`.

`com_get` / `com_call` are kept for old call sites; new code uses `v`.
"""

from typing import Any

import pythoncom
import win32com.client
from win32com.client import dynamic

PROGID = "SldWorks.Application"
_GET_OR_CALL = pythoncom.DISPATCH_METHOD | pythoncom.DISPATCH_PROPERTYGET


# ---------------------------------------------------------------------------
# Wrapping
# ---------------------------------------------------------------------------

def dispatch(obj: Any):
    """Late-bound wrapper for any COM object / IDispatch. Always a
    `dynamic.CDispatch`, whatever sits in the gen_py cache."""
    if obj is None:
        return None
    if isinstance(obj, dynamic.CDispatch):
        return obj
    return dynamic.Dispatch(getattr(obj, "_oleobj_", obj))


def connect_app():
    """The SolidWorks application, late-bound. Attaches to the running
    instance if there is one (SolidWorks registers as a single-instance
    server)."""
    return dynamic.Dispatch(PROGID)


# ---------------------------------------------------------------------------
# Members
# ---------------------------------------------------------------------------

def v(obj: Any, name: str, *args) -> Any:
    """Read a property / call a method by name, deterministically.

    Goes straight to IDispatch::Invoke with METHOD|PROPERTYGET, so it neither
    depends on win32com's property-vs-method guess nor caches it on the
    wrapper. Works for zero-arg methods, properties, indexed properties and
    methods with args (including out_*() VARIANTs). Returned IDispatch values
    come back as dynamic wrappers."""
    ole = getattr(obj, "_oleobj_", None)
    if ole is None:
        raise TypeError(f"v(): not a COM object: {type(obj).__name__}")
    dispid = ole.GetIDsOfNames(0, name)
    ret = ole.Invoke(dispid, 0, _GET_OR_CALL, True, *args)
    return _good(ret, name)


def setp(obj: Any, name: str, value: Any) -> None:
    """Write a property (PROPERTYPUT)."""
    ole = obj._oleobj_
    ole.Invoke(ole.GetIDsOfNames(0, name), 0, pythoncom.DISPATCH_PROPERTYPUT, False, value)


def _good(ret, name=None):
    if type(ret).__name__ == "PyIDispatch":
        return dynamic.Dispatch(ret, name)
    if isinstance(ret, tuple):
        return tuple(_good(x) for x in ret)
    return ret


# ---------------------------------------------------------------------------
# Out-parameters and empty objects
# ---------------------------------------------------------------------------

def byref(vt: int, value: Any = None):
    """By-ref VARIANT for an [out] parameter; read `.value` after the call."""
    return win32com.client.VARIANT(pythoncom.VT_BYREF | vt, value)


def out_int(value: int = 0):
    return byref(pythoncom.VT_I4, value)


def out_bool(value: bool = False):
    return byref(pythoncom.VT_BOOL, value)


def out_float(value: float = 0.0):
    return byref(pythoncom.VT_R8, value)


def out_str(value: str = ""):
    return byref(pythoncom.VT_BSTR, value)


def out_variant(value: Any = None):
    return byref(pythoncom.VT_VARIANT, value)


def out_dispatch():
    return byref(pythoncom.VT_DISPATCH, None)


def nothing():
    """VB `Nothing` for an object argument (Callout, Data, ...)."""
    return win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)


def double_array(values):
    """SAFEARRAY of doubles (IView.Position, ...)."""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x) for x in values])


class Out:
    """Placeholder for call_out(): `OUT` = long, `Out(pythoncom.VT_BOOL)`..."""

    _DEFAULTS = {pythoncom.VT_I4: 0, pythoncom.VT_BOOL: False,
                 pythoncom.VT_R8: 0.0, pythoncom.VT_BSTR: ""}

    def __init__(self, vt: int = pythoncom.VT_I4, value: Any = None):
        self.vt = vt
        self.value = value

    def make(self):
        return byref(self.vt, self._DEFAULTS.get(self.vt) if self.value is None else self.value)


OUT = Out()  # the common case: a long out-param (Errors, Warnings, ...)


def call_out(obj: Any, name: str, *args):
    """Call a method with [out] params marked by OUT / Out(vt).
    Returns (ret, out1, out2, ...), e.g.

        doc, err, warn = call_out(sw, "OpenDoc6", path, 2, 1, "", OUT, OUT)
    """
    real, outs = [], []
    for a in args:
        if isinstance(a, Out):
            var = a.make()
            outs.append(var)
            real.append(var)
        else:
            real.append(a)
    ret = v(obj, name, *real)
    return (ret, *[_good(o.value) for o in outs])


# ---------------------------------------------------------------------------
# Legacy names
# ---------------------------------------------------------------------------

def com_get(obj: Any, name: str) -> Any:
    """Read a zero-argument member (property or method); same as v(obj, name)."""
    return v(obj, name)


def com_call(obj: Any, name: str, *args) -> Any:
    """Call a member that takes arguments; same as v(obj, name, *args)."""
    return v(obj, name, *args)
