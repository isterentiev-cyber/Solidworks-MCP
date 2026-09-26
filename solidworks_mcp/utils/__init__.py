"""
SolidWorks MCP Utilities
------------------------
Utility modules for unit conversion, SolidWorks detection, and validation.
"""

from .units import (
    Unit,
    UnitConverter,
    mm, cm, inch, ft,
    to_mm, to_inch,
    get_converter,
    set_default_unit,
)

from .sw_finder import (
    SolidWorksFinder,
    find_solidworks,
    find_template,
    get_solidworks_info,
)

from .com_helpers import com_get, com_call
from . import com_helpers as com

from .typelib import get_signature, get_constant, get_iid, out_params

__all__ = [
    # Units
    "Unit",
    "UnitConverter", 
    "mm", "cm", "inch", "ft",
    "to_mm", "to_inch",
    "get_converter",
    "set_default_unit",
    
    # SolidWorks Finder
    "SolidWorksFinder",
    "find_solidworks",
    "find_template",
    "get_solidworks_info",

    # COM access helpers
    "com_get",
    "com_call",
    "com",

    # Typelib introspection (real method signatures / constants)
    "get_signature",
    "get_constant",
    "get_iid",
    "out_params",
]
