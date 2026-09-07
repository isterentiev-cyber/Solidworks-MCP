"""
SolidWorks COM Access Helpers
-----------------------------
Utilities for reading zero-argument SolidWorks API members whose
dispatch kind (property vs. method) is inconsistent across SolidWorks
versions and cannot be told apart with `callable()`.

Why this exists
----------------
win32com's dynamic dispatch resolves some zero-arg COM members as plain
Python values the moment you access the attribute (e.g. `doc.GetTitle`
already returns the string), while others resolve to a bound
`win32com.client.CDispatch` wrapper that still needs to be invoked with
`()` to get the actual value.

The obvious-looking fix, `if callable(x): x = x()`, is wrong: EVERY
`CDispatch` object is callable (it implements `__call__` to support a
COM object's default member, e.g. VB's `obj(0)`), including ones that
are already the fully-resolved result (e.g. the next feature in
`feat.GetNextFeature`). Calling those raises a COM error such as
"Member not found" instead of a Python TypeError, and unguarded code
mistakes that for "no more items" and stops early -- this is what made
`list_features` (and friends) stop after the very first feature.

`com_get()` sidesteps the ambiguity: it only trusts the call attempt
when it actually succeeds, and otherwise falls back to the value that
attribute access already gave it.
"""

from typing import Any


def com_get(obj: Any, name: str) -> Any:
    """
    Read a zero-argument SolidWorks COM member regardless of whether
    this SolidWorks version exposes it as a property or a method.

    Args:
        obj: The COM object (e.g. a document or feature dispatch).
        name: Attribute name to read (e.g. "GetTypeName2").

    Returns:
        The resolved value. If attribute access alone already produced
        a plain Python value (str/int/float/bool/tuple/None), that
        value is returned as-is. If it produced a callable COM
        sub-object, calling it is attempted; if the call raises
        (typically because the attribute access had already fully
        resolved the member), the pre-call value is returned instead.
    """
    value = getattr(obj, name)
    if not callable(value):
        return value
    try:
        return value()
    except Exception:
        return value


def com_call(obj: Any, name: str, *args) -> Any:
    """
    Call a SolidWorks COM member that takes arguments. Plain passthrough
    (`getattr(obj, name)(*args)`) -- kept alongside `com_get` so call
    sites can express "this one always needs a real call" explicitly.
    """
    return getattr(obj, name)(*args)
