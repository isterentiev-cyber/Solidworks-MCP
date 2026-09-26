"""
SolidWorks MCP Server
---------------------
Main MCP server entry point with all tools.

Version: 4.0.0 (Fixed for SolidWorks 2025)
Author: Samsaam Ali Baig

Fixes v4.0.0:
- execute_python now captures stdout/stderr
- FeatureExtrusion2 with correct 23 params for SW 2025
- list_features: property access instead of method calls
- extrude_sketch: proper sketch close + select before extrude
- cut_extrude: proper sketch handling
- NEW: close_sketch tool
- NEW: get_sketch_status tool for diagnostics
"""

import io
import os
import sys
import json
import importlib
import logging
import traceback
from typing import Dict
from pathlib import Path

# MCP imports
from mcp.server import Server, NotificationOptions
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# Local imports
from .automation import SolidWorksAutomation
from .constants import SwErrors
from .config import get_config, save_config
# Modules, not names: `from .utils import get_signature` pins the function
# object that existed when server.py was imported, and server.py is never
# reloaded -- reload_api then left lookup_api_signature on the old code.
from .utils import sw_finder, units, com_helpers, typelib
from . import ext
from . import toolsets

# Configure logging
config = get_config()
LOG_FILE = Path(__file__).parent / config.log_file

logging.basicConfig(
    level=config.get_log_level_int(),
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler(LOG_FILE, encoding='utf-8')]
)
logger = logging.getLogger("SolidWorksMCP")

# ============================================================================
# Global Instances
# ============================================================================

sw_automation = SolidWorksAutomation()
server = Server("solidworks-mcp-server")


# ============================================================================
# Tool Definitions
# ============================================================================

@server.list_tools()
async def list_tools() -> list[Tool]:
    """List all available SolidWorks tools"""
    base = [
        # Connection Tools
        Tool(
            name="connect_solidworks",
            description="Connect to SolidWorks. Launches if not running.",
            inputSchema={"type": "object", "properties": {}, "required": []}
        ),
        Tool(
            name="get_solidworks_info",
            description="Get SolidWorks installation information.",
            inputSchema={"type": "object", "properties": {}, "required": []}
        ),
        
        # Document Tools
        Tool(
            name="create_new_part",
            description="Create a new part document.",
            inputSchema={"type": "object", "properties": {}, "required": []}
        ),
        Tool(
            name="create_new_assembly",
            description="Create a new assembly document.",
            inputSchema={"type": "object", "properties": {}, "required": []}
        ),
        Tool(
            name="open_document",
            description="Open an existing SolidWorks document (silent). Reports decoded load "
                        "errors/warnings; for an assembly also component counts, how many came in "
                        "lightweight (Large Assembly Mode) and resolves them unless told not to.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to file"},
                    "resolve_lightweight": {"type": "boolean", "default": True,
                                            "description": "Assemblies: resolve lightweight components after opening"},
                    "read_only": {"type": "boolean", "default": False}
                },
                "required": ["filepath"]
            }
        ),
        Tool(
            name="save_document",
            description="Save the active document.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {"type": "string", "description": "Path to save (optional for Save As)"}
                },
                "required": []
            }
        ),
        Tool(
            name="close_document",
            description="Close the active document.",
            inputSchema={
                "type": "object",
                "properties": {
                    "save": {"type": "boolean", "default": False, "description": "Save before closing"}
                },
                "required": []
            }
        ),
        Tool(
            name="get_document_info",
            description="Get information about the active document.",
            inputSchema={"type": "object", "properties": {}, "required": []}
        ),
        Tool(
            name="list_open_documents",
            description="List all open documents.",
            inputSchema={"type": "object", "properties": {}, "required": []}
        ),
        Tool(
            name="capture_view",
            description=(
                "Screenshot the active document's current view to a PNG file. "
                "Use this to get an image of a part/assembly for visual "
                "inspection (read the returned path with the Read tool) before "
                "cross-checking your visual read against list_features / "
                "execute_python geometry data."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "output_path": {"type": "string", "description": "Where to save the PNG (default: <part folder>/_screenshots/<title>_<timestamp>.png)"},
                    "view": {
                        "type": "string",
                        "enum": ["isometric", "front", "back", "left", "right", "top", "bottom", "trimetric", "dimetric"],
                        "description": "Named view to switch to before capturing (default: current view)"
                    },
                    "width": {"type": "integer", "description": "Capture width in pixels (default from config, 1920)"},
                    "height": {"type": "integer", "description": "Capture height in pixels (default from config, 1080)"},
                    "zoom_to_fit": {"type": "boolean", "default": True, "description": "Zoom to fit the model before capturing"}
                },
                "required": []
            }
        ),

        # Sketch Tools
        Tool(
            name="create_sketch",
            description=(
                "Create a new sketch on a default plane (optionally on a parallel plane at `offset`). "
                "The result reports the sketch frame: where sketch X/Y point in model space."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "plane": {
                        "type": "string",
                        "enum": ["Front", "Top", "Right"],
                        "default": "Front",
                        "description": "Plane to sketch on"
                    },
                    "offset": {"type": "number", "default": 0, "description": "Offset along the plane normal (creates a reference plane); negative = other side"},
                    "unit": {"type": "string", "description": "Unit for offset"}
                },
                "required": []
            }
        ),
        Tool(
            name="create_sketch_on_face",
            description=(
                "Create a new sketch on a planar body face, by face_index (from list_faces) or by a 3D point on it. "
                "Sketch coordinates are LOCAL to the face -- read the reported sketch frame before drawing."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "face_index": {"type": "integer", "default": 0, "description": "Face index from list_faces (overrides x/y/z)"},
                    "x": {"type": "number", "default": 0, "description": "X coordinate on the face"},
                    "y": {"type": "number", "default": 0, "description": "Y coordinate on the face"},
                    "z": {"type": "number", "default": 0, "description": "Z coordinate on the face"},
                    "unit": {"type": "string", "description": "Unit (mm, inch, m)"}
                },
                "required": []
            }
        ),
        Tool(
            name="draw_line",
            description="Draw a line in the active sketch (construction=true -> centerline, e.g. revolve axis).",
            inputSchema={
                "type": "object",
                "properties": {
                    "construction": {"type": "boolean", "default": False, "description": "Centerline instead of a profile line"},
                    "x1": {"type": "number", "default": 0, "description": "Start X"},
                    "y1": {"type": "number", "default": 0, "description": "Start Y"},
                    "x2": {"type": "number", "default": 100, "description": "End X"},
                    "y2": {"type": "number", "default": 0, "description": "End Y"},
                    "unit": {"type": "string", "description": "Unit (mm, inch, m)"}
                },
                "required": []
            }
        ),
        Tool(
            name="draw_circle",
            description="Draw a circle in the active sketch.",
            inputSchema={
                "type": "object",
                "properties": {
                    "x": {"type": "number", "default": 0, "description": "Center X"},
                    "y": {"type": "number", "default": 0, "description": "Center Y"},
                    "radius": {"type": "number", "default": 25, "description": "Radius"},
                    "unit": {"type": "string", "description": "Unit (mm, inch, m)"}
                },
                "required": []
            }
        ),
        Tool(
            name="draw_rectangle",
            description="Draw a rectangle in the active sketch.",
            inputSchema={
                "type": "object",
                "properties": {
                    "x1": {"type": "number", "default": -50, "description": "First corner X"},
                    "y1": {"type": "number", "default": -25, "description": "First corner Y"},
                    "x2": {"type": "number", "default": 50, "description": "Second corner X"},
                    "y2": {"type": "number", "default": 25, "description": "Second corner Y"},
                    "unit": {"type": "string", "description": "Unit (mm, inch, m)"}
                },
                "required": []
            }
        ),
        Tool(
            name="draw_arc",
            description="Draw an arc by center and angles.",
            inputSchema={
                "type": "object",
                "properties": {
                    "cx": {"type": "number", "default": 0, "description": "Center X"},
                    "cy": {"type": "number", "default": 0, "description": "Center Y"},
                    "radius": {"type": "number", "default": 25, "description": "Radius"},
                    "start_angle": {"type": "number", "default": 0, "description": "Start angle (degrees)"},
                    "end_angle": {"type": "number", "default": 90, "description": "End angle (degrees)"},
                    "unit": {"type": "string", "description": "Unit for radius"}
                },
                "required": []
            }
        ),
        Tool(
            name="draw_polygon",
            description="Draw a regular polygon.",
            inputSchema={
                "type": "object",
                "properties": {
                    "cx": {"type": "number", "default": 0, "description": "Center X"},
                    "cy": {"type": "number", "default": 0, "description": "Center Y"},
                    "radius": {"type": "number", "default": 25, "description": "Radius"},
                    "sides": {"type": "integer", "default": 6, "description": "Number of sides (3-100)"},
                    "unit": {"type": "string", "description": "Unit"}
                },
                "required": []
            }
        ),
        
        # Feature Tools
        Tool(
            name="extrude_sketch",
            description="Extrude the active sketch (Boss-Extrude).",
            inputSchema={
                "type": "object",
                "properties": {
                    "depth": {"type": "number", "default": 10, "description": "Extrusion depth"},
                    "both_directions": {"type": "boolean", "default": False, "description": "Extrude in both directions"},
                    "reverse": {"type": "boolean", "default": False, "description": "Extrude opposite to the sketch normal"},
                    "unit": {"type": "string", "description": "Unit"}
                },
                "required": []
            }
        ),
        Tool(
            name="revolve_sketch",
            description=(
                "Revolve the active sketch (Boss/Cut-Revolve) around its single centerline, "
                "or around `axis` (world X/Y/Z, axis feature, edge) -- then no centerline is needed. "
                "Profile must be closed and on one side of the axis."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "angle": {"type": "number", "default": 360, "description": "Revolve angle in degrees"},
                    "cut": {"type": "boolean", "default": False, "description": "Cut-revolve instead of boss-revolve"},
                    "reverse": {"type": "boolean", "default": False, "description": "Revolve the other way (for angles < 360)"},
                    "axis": {"type": "string", "description": "Axis outside the sketch instead of a centerline: 'X'/'Y'/'Z' (world, through origin), 'edge:N', or an axis feature name"}
                },
                "required": []
            }
        ),
        Tool(
            name="cut_extrude",
            description="Cut extrude to remove material. Direction is verified by volume and retried the other way if nothing was removed.",
            inputSchema={
                "type": "object",
                "properties": {
                    "depth": {"type": "number", "default": 10, "description": "Cut depth"},
                    "through_all": {"type": "boolean", "default": False, "description": "Cut through all"},
                    "both_directions": {"type": "boolean", "default": False, "description": "Cut both directions"},
                    "reverse": {"type": "boolean", "default": False, "description": "Try the opposite direction first"},
                    "unit": {"type": "string", "description": "Unit"}
                },
                "required": []
            }
        ),
        Tool(
            name="fillet_edges",
            description="Add fillet to edges (edge_indices from list_edges, or the current selection).",
            inputSchema={
                "type": "object",
                "properties": {
                    "edge_indices": {"type": "string", "description": "e.g. '1,2,5' (from list_edges); omit to use current selection"},
                    "radius": {"type": "number", "default": 2, "description": "Fillet radius"},
                    "unit": {"type": "string", "description": "Unit"}
                },
                "required": []
            }
        ),
        Tool(
            name="chamfer_edges",
            description="Add chamfer (distance x angle) to edges (edge_indices from list_edges, or the current selection).",
            inputSchema={
                "type": "object",
                "properties": {
                    "edge_indices": {"type": "string", "description": "e.g. '1,2,5' (from list_edges); omit to use current selection"},
                    "distance": {"type": "number", "default": 2, "description": "Chamfer distance"},
                    "angle": {"type": "number", "default": 45, "description": "Chamfer angle (degrees)"},
                    "unit": {"type": "string", "description": "Unit"}
                },
                "required": []
            }
        ),
        Tool(
            name="list_features",
            description="List all features in the model.",
            inputSchema={"type": "object", "properties": {}, "required": []}
        ),
        
        # Sketch Management Tools
        Tool(
            name="close_sketch",
            description="Close/exit the active sketch. Call this before extrude if sketch is still open.",
            inputSchema={"type": "object", "properties": {}, "required": []}
        ),
        Tool(
            name="get_sketch_status",
            description="Get diagnostic info: active sketch state, sketch count, sketch names in feature tree.",
            inputSchema={"type": "object", "properties": {}, "required": []}
        ),
        
        # Utility Tools
        Tool(
            name="set_units",
            description="Set default unit for dimensions.",
            inputSchema={
                "type": "object",
                "properties": {
                    "unit": {
                        "type": "string",
                        "enum": ["mm", "inch", "m", "cm"],
                        "description": "Default unit"
                    }
                },
                "required": ["unit"]
            }
        ),
        Tool(
            name="execute_python",
            description=(
                "Execute Python against the live SolidWorks connection; print() your results. "
                "Everything is late-bound (dynamic dispatch; never gencache/EnsureDispatch). "
                "Pre-bound (fresh each call): sw (app), doc (ActiveDoc), md (doc as IModelDoc2), "
                "T(obj,'IFace2') = QueryInterface, v(obj,'Member',*args) read/call -- use it for zero-arg "
                "methods (obj.GetTitle() fails), call_out(obj,'M',..,OUT,OUT) -> (ret,*outs) for by-ref "
                "out params, nothing() for a null object arg, com (out_int/out_bool/... helpers), "
                "ext (face_rows, edge_rows, select_entities, find_feature, snapshot, delta...), math, json. "
                "Variables you assign persist between calls. SW API units are metres/radians."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "code": {"type": "string", "description": "Python code to execute"}
                },
                "required": ["code"]
            }
        ),
        Tool(
            name="lookup_api_signature",
            description=(
                "Get the REAL parameter list of a SolidWorks COM method, generated "
                "from SolidWorks' own type library (not memory/docs, which have been "
                "wrong before -- see CLAUDE.md). Use this BEFORE guessing a param "
                "count for any FeatureManager/SketchManager/etc. method you haven't "
                "already confirmed working. First call generates a one-time cache "
                "(~1-2s); later calls are instant."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "interface": {"type": "string", "description": "COM interface name, e.g. IFeatureManager, IModelDoc2, ISketchManager"},
                    "member": {"type": "string", "description": "Method or property name, e.g. FeatureRevolve2"}
                },
                "required": ["interface", "member"]
            }
        ),
        Tool(
            name="lookup_api_constant",
            description="Get the integer value of a SolidWorks API constant (e.g. swEndCondBlind, swRefPlaneReferenceConstraint_Distance) from SolidWorks' own constant type library, instead of guessing.",
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Constant name, e.g. swRefPlaneReferenceConstraint_Distance"}
                },
                "required": ["name"]
            }
        ),
    ]
    all_tools = base + [Tool(name=n, description=d, inputSchema=s)
                        for n, d, s in ext.TOOL_SCHEMAS]
    # Trim the list to the sections asked for in SW_MCP_TOOLSETS. Every
    # schema here is re-sent to the model on every request, so an unused
    # half of the list is pure context cost. See solidworks_mcp/toolsets.py.
    return toolsets.filter_tools(all_tools)


# ============================================================================
# Result Formatter
# ============================================================================

def format_result(r: Dict) -> str:
    """Format result dictionary as readable text"""
    # Compact on purpose: this is read by a model, and the upstream
    # indent=2 JSON dump often repeated the message (execute_python
    # printed its whole stdout twice).
    status = "OK" if r["success"] else "ERROR"
    text = f"[{status}] {r['message']}"
    data = r.get("data")
    if data and not r.get("_no_data"):
        text += "\n" + json.dumps(data, ensure_ascii=False, separators=(",", ":"), default=str)
    return text


# ============================================================================
# Tool Handlers
# ============================================================================

# Tools that must answer even while SolidWorks is wedged: they either do not
# touch the live document at all, or they are how you find out what is wrong.
BUSY_EXEMPT = {
    "connect_solidworks",
    "get_solidworks_info",
    "lookup_api_signature",
    "lookup_api_constant",
    "reload_api",
    "set_units",
}


def _busy_timeout_ms() -> int:
    """SW_MCP_BUSY_TIMEOUT_MS overrides the default; 0 disables the probe."""
    raw = os.environ.get("SW_MCP_BUSY_TIMEOUT_MS", "").strip()
    if not raw:
        return ext.BUSY_TIMEOUT_MS
    try:
        return int(raw)
    except ValueError:
        logger.warning(f"SW_MCP_BUSY_TIMEOUT_MS={raw!r} is not a number, using default")
        return ext.BUSY_TIMEOUT_MS


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    """Handle MCP tool calls"""
    try:
        logger.info(f"Tool: {name}, Args: {arguments}")
        arguments = arguments or {}

        # Refuse rather than block. Everything below marshals COM into
        # SLDWORKS.exe synchronously on this thread, so a modal dialog or a
        # long rebuild over there does not slow us down -- it freezes the
        # whole server. One ~0.3 ms window ping turns that into an answer.
        # Tools that do not touch the live document are exempt: they must
        # keep working precisely when SolidWorks is stuck.
        if name not in BUSY_EXEMPT and sw_automation.is_connected:
            busy = ext.sw_busy(sw_automation.app, _busy_timeout_ms())
            if busy:
                logger.warning(f"{name} refused: {busy}")
                return [TextContent(type="text", text=f"[ERROR] {busy}")]

        # Feature tools get a volume/topology delta appended (ext.delta),
        # so every step is self-verifying without extra inspect calls.
        before = None
        if name in ext.REPORTED and sw_automation.is_connected:
            before = ext.snapshot(sw_automation.app)
            if name in ("fillet_edges", "chamfer_edges") and arguments.get("edge_indices"):
                ext.select_entities(ext.model(sw_automation.app), "edge",
                                    arguments["edge_indices"], mark=1)

        # Connection Tools
        if name == "connect_solidworks":
            result = sw_automation.connect()
        
        elif name == "get_solidworks_info":
            info = sw_finder.get_solidworks_info()
            result = {
                "success": info["found"],
                "message": f"SolidWorks {'found' if info['found'] else 'not found'}",
                "error_code": 0 if info["found"] else 105,
                "error_name": "swSuccess" if info["found"] else "swSolidWorksNotFound",
                "data": info
            }
        
        # Document Tools
        elif name == "create_new_part":
            result = sw_automation.create_new_part()
        
        elif name == "create_new_assembly":
            result = sw_automation.create_new_assembly()
        
        elif name == "open_document":
            result = sw_automation.open_document(
                arguments.get("filepath", ""),
                resolve_lightweight=bool(arguments.get("resolve_lightweight", True)),
                read_only=bool(arguments.get("read_only", False)))
        
        elif name == "save_document":
            result = sw_automation.save_document(arguments.get("filepath"))
        
        elif name == "close_document":
            result = sw_automation.close_document(arguments.get("save", False))
        
        elif name == "get_document_info":
            result = sw_automation.get_document_info()

        elif name == "capture_view":
            result = sw_automation.capture_view(
                output_path=arguments.get("output_path"),
                view=arguments.get("view"),
                width=arguments.get("width"),
                height=arguments.get("height"),
                zoom_to_fit=arguments.get("zoom_to_fit", True),
            )
        
        elif name == "list_open_documents":
            result = sw_automation.list_open_documents()
        
        # Sketch Tools
        elif name == "create_sketch":
            result = sw_automation.create_sketch(
                arguments.get("plane", "Front"),
                arguments.get("offset", 0),
                arguments.get("unit")
            )

        elif name == "create_sketch_on_face":
            result = sw_automation.create_sketch_on_face(
                arguments.get("x", 0),
                arguments.get("y", 0),
                arguments.get("z", 0),
                arguments.get("unit"),
                arguments.get("face_index", 0)
            )

        elif name == "draw_line":
            result = sw_automation.draw_line(
                arguments.get("x1", 0),
                arguments.get("y1", 0),
                arguments.get("x2", 100),
                arguments.get("y2", 0),
                arguments.get("unit"),
                arguments.get("construction", False)
            )
        
        elif name == "draw_circle":
            result = sw_automation.draw_circle(
                arguments.get("x", 0),
                arguments.get("y", 0),
                arguments.get("radius", 25),
                arguments.get("unit")
            )
        
        elif name == "draw_rectangle":
            result = sw_automation.draw_rectangle(
                arguments.get("x1", -50),
                arguments.get("y1", -25),
                arguments.get("x2", 50),
                arguments.get("y2", 25),
                arguments.get("unit")
            )
        
        elif name == "draw_arc":
            result = sw_automation.draw_arc_center(
                arguments.get("cx", 0),
                arguments.get("cy", 0),
                arguments.get("radius", 25),
                arguments.get("start_angle", 0),
                arguments.get("end_angle", 90),
                arguments.get("unit")
            )
        
        elif name == "draw_polygon":
            result = sw_automation.draw_polygon(
                arguments.get("cx", 0),
                arguments.get("cy", 0),
                arguments.get("radius", 25),
                arguments.get("sides", 6),
                arguments.get("unit")
            )
        
        # Feature Tools
        elif name == "extrude_sketch":
            result = sw_automation.extrude_sketch(
                arguments.get("depth", 10),
                arguments.get("both_directions", False),
                arguments.get("unit"),
                arguments.get("reverse", False)
            )
        
        elif name == "revolve_sketch":
            result = sw_automation.revolve_sketch(
                arguments.get("angle", 360),
                arguments.get("cut", False),
                arguments.get("reverse", False),
                arguments.get("axis")
            )

        elif name == "cut_extrude":
            result = sw_automation.cut_extrude(
                arguments.get("depth", 10),
                arguments.get("through_all", False),
                arguments.get("both_directions", False),
                arguments.get("unit"),
                arguments.get("reverse", False)
            )
        
        elif name == "fillet_edges":
            result = sw_automation.fillet_edges(
                arguments.get("radius", 2),
                arguments.get("unit")
            )
        
        elif name == "chamfer_edges":
            result = sw_automation.chamfer_edges(
                arguments.get("distance", 2),
                arguments.get("angle", 45),
                arguments.get("unit")
            )
        
        elif name == "list_features":
            result = _list_features_fixed()
        
        # Sketch Management Tools
        elif name == "close_sketch":
            result = _close_sketch_handler()
        
        elif name == "get_sketch_status":
            result = _get_sketch_status_handler()
        
        # Utility Tools
        elif name == "set_units":
            unit = arguments.get("unit", "mm")
            units.set_default_unit(unit)
            sw_automation._units.default_unit = unit
            result = {
                "success": True,
                "message": f"Default unit set to: {unit}",
                "error_code": 0,
                "error_name": "swSuccess",
                "data": {"unit": unit}
            }
        
        elif name == "execute_python":
            code = arguments.get("code", "")
            if not code:
                result = sw_automation._result(False, "Code is required", SwErrors.swInvalidInput)
            else:
                result = _execute_python_fixed(code)

        elif name == "lookup_api_signature":
            result = _lookup_api_signature_handler(
                arguments.get("interface", ""), arguments.get("member", "")
            )

        elif name == "lookup_api_constant":
            result = _lookup_api_constant_handler(arguments.get("name", ""))

        elif name == "reload_api":
            result = _reload_api()
            if result.get("success"):
                # list_tools() reads ext.TOOL_SCHEMAS at call time and
                # importlib.reload mutates the module in place, so the fresh
                # list is already what the next tools/list would return --
                # the client just has to be told to ask again.
                try:
                    await server.request_context.session.send_tool_list_changed()
                    result["message"] += " | tools/list_changed sent"
                except Exception as e:
                    logger.warning(f"send_tool_list_changed failed: {e}")
                    result["message"] += (
                        f" | could NOT notify the client ({e}) -- "
                        f"restart the MCP client to see new tools"
                    )

        elif name in ext.HANDLERS:
            result = ext.dispatch(name, arguments, sw_automation)

        else:
            result = sw_automation._result(False, f"Unknown tool: {name}", SwErrors.swUnknownError)

        if before is not None and result.get("success"):
            d = ext.delta(sw_automation.app, before, ext.snapshot(sw_automation.app),
                          ext.expect_for(name, arguments))
            if d:
                result["message"] += " | " + d
            result.pop("data", None)

        logger.info(f"Result: success={result['success']}")
        return [TextContent(type="text", text=format_result(result))]
        
    except Exception as e:
        logger.error(f"Tool error: {e}\n{traceback.format_exc()}")
        return [TextContent(type="text", text=f"[ERROR] {e}")]


# ============================================================================
# FIXED: Execute Python with stdout capture
# ============================================================================

# One namespace for the whole server process: variables assigned in one
# execute_python call are visible in the next (like a REPL).
_PY_NS: Dict = {}


def _execute_python_fixed(code: str) -> Dict:
    """
    Execute Python against the live SolidWorks connection.
    stdout/stderr are captured and returned even when the code raises --
    the upstream version dropped everything printed before the error.
    """
    if not sw_automation.is_connected:
        r = sw_automation.connect()
        if not r["success"]:
            return r

    import math
    import os as os_module
    import win32com.client
    import pythoncom
    from .utils import com_helpers as com

    app = sw_automation.app
    doc = com.v(app, "ActiveDoc") if app else None
    _PY_NS.update({
        "sw": app,
        "doc": doc,
        "md": ext.T(doc, "IModelDoc2") if doc is not None else None,
        "T": ext.T,
        "v": com.v,
        "com": com,
        "nothing": com.nothing,
        "OUT": com.OUT,
        "call_out": com.call_out,
        "ext": ext,
        "g": com_helpers.com_get,
        "automation": sw_automation,
        "win32com": win32com,
        "pythoncom": pythoncom,
        "math": math,
        "os": os_module,
        "json": json,
        "result": None,
    })

    out, errs = io.StringIO(), io.StringIO()
    old_out, old_err = sys.stdout, sys.stderr
    failure = None
    try:
        sys.stdout, sys.stderr = out, errs
        exec(code, _PY_NS)
    except Exception:
        failure = traceback.format_exc(limit=-3)
    finally:
        sys.stdout, sys.stderr = old_out, old_err

    parts = []
    if out.getvalue():
        parts.append(out.getvalue().rstrip())
    if errs.getvalue():
        parts.append("[stderr] " + errs.getvalue().rstrip())
    if _PY_NS.get("result") is not None:
        parts.append(f"[result] {_PY_NS['result']}")
    if failure:
        parts.append("[exception]\n" + failure.rstrip())
    msg = "\n".join(parts) or "(no output)"
    return {
        "success": failure is None,
        "message": msg,
        "error_code": 0 if failure is None else 999,
        "error_name": "swSuccess" if failure is None else "swUnknownError",
    }


# ============================================================================
# reload_api: hot-reload automation mixins + ext without restarting
# ============================================================================

# Dependency order: each module imports only from modules above it. The
# utils package comes after its submodules -- reloading it re-binds the names
# it re-exports (automation.base does `from ..utils import com_get, ...`).
# Any other solidworks_mcp.utils.* found in sys.modules is reloaded too,
# right before the package, so a new utils module can't silently stay old
# (2026-09-26: typelib was missing here, the fresh ext.py failed on
# `from .utils.typelib import get_iid` and left the server half-reloaded).
_RELOAD_ORDER = [
    "solidworks_mcp.utils.units",
    "solidworks_mcp.utils.sw_finder",
    "solidworks_mcp.utils.com_helpers",
    "solidworks_mcp.utils.typelib",
    "solidworks_mcp.utils",
    "solidworks_mcp.toolsets",
    "solidworks_mcp.ext",
    "solidworks_mcp.automation.base",
    "solidworks_mcp.automation.documents",
    "solidworks_mcp.automation.sketches",
    "solidworks_mcp.automation.features",
    "solidworks_mcp.automation.capture",
    "solidworks_mcp.automation",
]
_UTILS_PKG = "solidworks_mcp.utils"


def _short(name: str) -> str:
    return name[len("solidworks_mcp."):] if name.startswith("solidworks_mcp.") else name


def _reload_plan() -> list:
    order = list(_RELOAD_ORDER)
    extra = sorted(n for n in sys.modules
                   if n.startswith(_UTILS_PKG + ".") and n not in order and sys.modules[n] is not None)
    i = order.index(_UTILS_PKG)
    order[i:i] = extra
    return [n for n in order if sys.modules.get(n) is not None]


def _reload_api() -> Dict:
    """Reload modules in dependency order, then swap the live automation
    instance onto the fresh class -- the COM connection (instance state)
    survives. server.py itself is not reloaded (tool schemas are fixed at
    MCP startup anyway), which is why it reaches library code only through
    module attributes (typelib.get_signature, not a from-imported name).

    All or nothing: every file is compiled first (a syntax error reloads
    nothing), and if a module then fails to execute (e.g. an ImportError on
    a name its dependency doesn't have yet), every module already reloaded
    in this call is put back to its previous namespace -- the server stays
    consistently on the old code instead of half on each."""
    plan = _reload_plan()

    bad = []
    for name in plan:
        path = getattr(sys.modules[name], "__file__", None)
        if not path:
            continue
        try:
            with open(path, "rb") as fh:
                compile(fh.read(), path, "exec")
        except Exception as e:
            bad.append(f"{_short(name)}: {type(e).__name__}: {e}")
    if bad:
        return {"success": False,
                "message": "reload NOT started, nothing changed -- the server keeps running the "
                           "previous code. Fix and call reload_api again:\n  " + "\n  ".join(bad),
                "error_code": 999, "error_name": "swUnknownError"}

    # set_units state: units.py re-creates its converter at mm. Carried as the
    # string -- the Unit enum of the old module is a foreign class to the new one.
    unit = getattr(units.get_converter().default_unit, "value", "mm")
    saved: Dict[str, dict] = {}
    current = None
    try:
        for name in plan:
            current = name
            mod = sys.modules[name]
            saved[name] = dict(mod.__dict__)
            importlib.reload(mod)
        sw_automation.__class__ = sys.modules["solidworks_mcp.automation"].SolidWorksAutomation
    except Exception as e:
        tb = traceback.format_exc(limit=-3)
        restored, broken = [], []
        for name in reversed(list(saved)):  # includes the module that failed
            try:
                ns = sys.modules[name].__dict__
                ns.clear()
                ns.update(saved[name])
                restored.append(_short(name))
            except Exception as re:
                broken.append(f"{_short(name)} ({re})")
        ok_before = [_short(n) for n in list(saved)[:-1]]
        rest = [_short(n) for n in plan[len(saved):]]
        state = ("rolled back, the server is consistently on the PREVIOUS code"
                 if not broken else
                 f"⚠ rollback failed for {', '.join(broken)} -- state is mixed, "
                 f"restart the server (scripts/restart_mcp.ps1)")
        return {"success": False,
                "message": (f"reload FAILED in {_short(current)}: {type(e).__name__}: {e}\n"
                            f"reloaded before it: {', '.join(ok_before) or '-'}; "
                            f"not attempted: {', '.join(rest) or '-'}\n"
                            f"{state} (restored {len(restored)} module(s)).\n{tb}"),
                "error_code": 999, "error_name": "swUnknownError"}
    finally:
        units.get_converter().default_unit = unit

    extra = [_short(n) for n in plan if n not in _RELOAD_ORDER]
    return {"success": True,
            "message": (f"reloaded {len(plan)} modules: {', '.join(_short(n) for n in plan)}"
                        + (f" (not in _RELOAD_ORDER, reloaded before utils: {', '.join(extra)})"
                           if extra else "")),
            "error_code": 0, "error_name": "swSuccess"}


# ============================================================================
# FIXED: list_features
# ============================================================================

def _list_features_fixed() -> Dict:
    """
    List all features in the active document.
    FIXED v4.1: Property access for SW 2025 (FirstFeature, GetNextFeature, GetTypeName2).
    """
    try:
        doc, err = sw_automation.get_active_doc()
        if err:
            return err
        
        features = []

        # FirstFeature: property on some SW versions, zero-arg method on
        # others -- com_get() reads it correctly either way.
        feat = com_helpers.com_get(doc, "FirstFeature")

        while feat is not None:
            try:
                try:
                    name = feat.Name
                except:
                    name = "<unknown>"

                try:
                    feat_type = com_helpers.com_get(feat, "GetTypeName2")
                except:
                    try:
                        feat_type = com_helpers.com_get(feat, "GetTypeName")
                    except:
                        feat_type = "<unknown>"

                try:
                    suppressed = com_helpers.com_get(feat, "IsSuppressed")
                except:
                    suppressed = False

                features.append({
                    "name": name,
                    "type": feat_type,
                    "suppressed": bool(suppressed)
                })

            except Exception as e:
                features.append({
                    "name": "<error>",
                    "type": str(e),
                    "suppressed": False
                })

            # NOTE: do NOT use `if callable(feat): feat = feat()` here.
            # win32com.client.CDispatch objects (including the already-
            # resolved next feature) are always callable, so that check
            # can't tell "still need a call" from "already have the
            # feature" -- calling the latter raises a COM error that
            # looks like "no more items" and stops the walk after one
            # feature. com_get() tries the call and falls back to the
            # pre-call value when it fails, which handles both cases.
            try:
                feat = com_helpers.com_get(feat, "GetNextFeature")
            except:
                break
        
        return {
            "success": True,
            "message": f"{len(features)} features found",
            "error_code": 0,
            "error_name": "swSuccess",
            "data": {"features": features, "count": len(features)}
        }
        
    except Exception as e:
        logger.error(f"List features error: {e}\n{traceback.format_exc()}")
        return {
            "success": False,
            "message": f"Error: {e}",
            "error_code": 999,
            "error_name": "swUnknownError",
            "data": {}
        }

# ============================================================================
# NEW: API typelib lookup handlers (see utils/typelib.py, CLAUDE.md)
# ============================================================================

def _lookup_api_signature_handler(interface: str, member: str) -> Dict:
    """Look up a COM method/property's real signature via the SW typelib."""
    if not interface or not member:
        return {
            "success": False,
            "message": "Both 'interface' and 'member' are required",
            "error_code": SwErrors.swInvalidInput,
            "error_name": "swInvalidInput",
            "data": {},
        }
    try:
        sig = typelib.get_signature(interface, member)
        return {
            "success": True,
            "message": f"{interface}.{member}",
            "error_code": 0,
            "error_name": "swSuccess",
            "data": {"signature": sig},
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e),
            "error_code": SwErrors.swUnknownError,
            "error_name": "swUnknownError",
            "data": {},
        }


def _lookup_api_constant_handler(name: str) -> Dict:
    """Look up a SolidWorks API constant's value via the constant typelib."""
    if not name:
        return {
            "success": False,
            "message": "'name' is required",
            "error_code": SwErrors.swInvalidInput,
            "error_name": "swInvalidInput",
            "data": {},
        }
    try:
        value = typelib.get_constant(name)
        return {
            "success": True,
            "message": f"{name} = {value}",
            "error_code": 0,
            "error_name": "swSuccess",
            "data": {"name": name, "value": value},
        }
    except Exception as e:
        return {
            "success": False,
            "message": str(e),
            "error_code": SwErrors.swUnknownError,
            "error_name": "swUnknownError",
            "data": {},
        }


# ============================================================================
# NEW: close_sketch handler
# ============================================================================

def _close_sketch_handler() -> Dict:
    """
    Close the active sketch if one is open.
    Returns sketch status info.
    """
    try:
        doc, err = sw_automation.get_active_doc()
        if err:
            return err
        
        # Check if sketch is active
        had_active = False
        try:
            active_sketch = doc.SketchManager.ActiveSketch
            had_active = active_sketch is not None
        except:
            pass
        
        if had_active:
            try:
                doc.SketchManager.InsertSketch(True)
            except:
                try:
                    doc.InsertSketch2(True)
                except:
                    pass
            
            return {
                "success": True,
                "message": "Sketch closed successfully",
                "error_code": 0,
                "error_name": "swSuccess",
                "data": {"had_active_sketch": True, "action": "closed"}
            }
        else:
            return {
                "success": True,
                "message": "No active sketch to close",
                "error_code": 0,
                "error_name": "swSuccess",
                "data": {"had_active_sketch": False, "action": "none"}
            }
        
    except Exception as e:
        logger.error(f"Close sketch error: {e}\n{traceback.format_exc()}")
        return {
            "success": False,
            "message": f"Error: {e}",
            "error_code": 999,
            "error_name": "swUnknownError",
            "data": {"traceback": traceback.format_exc()}
        }


# ============================================================================
# NEW: get_sketch_status handler
# ============================================================================

def _get_sketch_status_handler() -> Dict:
    """
    Get diagnostic info about the current sketch state.
    Useful for debugging sketch/extrude issues.
    """
    try:
        doc, err = sw_automation.get_active_doc()
        if err:
            return err
        
        info = {
            "has_active_sketch": False,
            "active_sketch_name": None,
            "sketch_count": 0,
            "sketch_names": [],
            "feature_count": 0,
            "extrusion_count": 0,
        }
        
        # Check active sketch
        try:
            active_sketch = doc.SketchManager.ActiveSketch
            if active_sketch is not None:
                info["has_active_sketch"] = True
                try:
                    info["active_sketch_name"] = active_sketch.Name
                except:
                    info["active_sketch_name"] = "<unknown>"
        except:
            pass
        
        # Walk the feature tree through ext (late-bound, v() for zero-arg members).
        try:
            for f in ext.user_features(ext.T(doc, "IModelDoc2")):
                feat_type = ext.v(f, "GetTypeName2")
                info["feature_count"] += 1
                if feat_type == "ProfileFeature":
                    info["sketch_count"] += 1
                    info["sketch_names"].append(f.Name)
                elif feat_type == "Extrusion":
                    info["extrusion_count"] += 1
        except Exception:
            pass
        
        # Build readable message
        status = "OPEN" if info["has_active_sketch"] else "CLOSED"
        msg = (f"Sketch status: {status}. "
               f"Sketches: {info['sketch_count']} {info['sketch_names']}. "
               f"Extrusions: {info['extrusion_count']}. "
               f"Total features: {info['feature_count']}")
        
        return {
            "success": True,
            "message": msg,
            "error_code": 0,
            "error_name": "swSuccess",
            "data": info
        }
        
    except Exception as e:
        logger.error(f"Sketch status error: {e}\n{traceback.format_exc()}")
        return {
            "success": False,
            "message": f"Error: {e}",
            "error_code": 999,
            "error_name": "swUnknownError",
            "data": {"traceback": traceback.format_exc()}
        }


# ============================================================================
# Main Entry Point
# ============================================================================

async def main():
    """Main entry point for MCP server"""
    logger.info("Starting SolidWorks MCP Server v4.0.0 (Fixed)...")
    logger.info(f"Log file: {LOG_FILE}")
    # No typelib preload: the server is late-bound only (utils/com_helpers.py).
    # utils/typelib.py reads the registered typelib directly on first use,
    # without generating or importing any makepy module.

    async with stdio_server() as (read_stream, write_stream):
        # tools_changed advertises tools.listChanged, which is what lets
        # reload_api add a NEW tool without restarting the MCP client.
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(
                NotificationOptions(tools_changed=True)
            ),
        )


def run():
    """Run the server"""
    import asyncio
    asyncio.run(main())


if __name__ == "__main__":
    run()
