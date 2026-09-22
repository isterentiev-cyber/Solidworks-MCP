"""
Toolset filtering
-----------------
Every tool schema is sent to the model on every single request, so a long
list is a standing tax on context -- and measured usage says most of it is
idle: across the recorded sessions only about half the tools were ever
called, and a typical modelling session touches 10-13 of them.

Set SW_MCP_TOOLSETS to the sections you actually want:

    SW_MCP_TOOLSETS=sketch,features          # + core, always on
    SW_MCP_TOOLSETS=analysis                 # read-mostly poking around
    SW_MCP_TOOLSETS=                         # unset/empty -> everything

`core` is implicit and cannot be switched off: without connect_solidworks,
execute_python and the lookup tools there is no way to recover from a
missing tool, which would turn a narrow toolset into a dead end.

Unknown names (a tool added later and not classified here) are kept, not
dropped -- a mistake in this table must not silently hide a working tool.

Changing this needs an MCP restart: tool schemas are fixed at startup, and
reload_api does not touch server.py.
"""

import logging
import os
from typing import Dict, Iterable, List, Set

logger = logging.getLogger(__name__)

ENV_VAR = "SW_MCP_TOOLSETS"

# Section -> tools. Mirrors how the work actually splits up, not how the
# code is organised.
TOOLSETS: Dict[str, Set[str]] = {
    # Session, documents, units, screenshots, and the escape hatches.
    "core": {
        "connect_solidworks",
        "get_solidworks_info",
        "create_new_part",
        "create_new_assembly",
        "open_document",
        "save_document",
        "close_document",
        "get_document_info",
        "list_open_documents",
        "set_units",
        "capture_view",
        "execute_python",
        "reload_api",
        "lookup_api_signature",
        "lookup_api_constant",
        "transaction",
    },
    # Finding and selecting geometry to act on.
    "select": {
        "list_faces",
        "list_edges",
        "find_face",
        "find_edge",
        "select_entities",
    },
    # Everything that happens inside an open sketch.
    "sketch": {
        "create_sketch",
        "create_sketch_on_face",
        "close_sketch",
        "get_sketch_status",
        "draw_line",
        "draw_circle",
        "draw_rectangle",
        "draw_arc",
        "draw_polygon",
        "draw_profile",
        "sketch_entities",
        "add_sketch_relation",
        "add_sketch_dimension",
    },
    # Solid operations and the feature tree.
    "features": {
        "extrude_sketch",
        "revolve_sketch",
        "cut_extrude",
        "fillet_edges",
        "chamfer_edges",
        "hole",
        "linear_pattern",
        "circular_pattern",
        "mirror_features",
        "delete_feature",
        "suppress_feature",
        "edit_feature",
        "create_reference_plane",
        "create_reference_axis",
    },
    # Reading the model back: topology, tree, parameters and equations.
    "analysis": {
        "inspect",
        "list_features",
        "get_parameters",
        "set_parameter",
        "add_parameter",
        "mass_properties",
        "get_rebuild_errors",
    },
    # Drawing documents: sheets, views and annotations of a saved part/assembly.
    "drawings": {
        "create_drawing",
        "add_standard_views",
        "add_drawing_view",
        "add_section_view",
        "add_detail_view",
        "add_broken_out_section",
        "insert_model_dimensions",
        "add_note",
        "export_pdf",
        "move_drawing_view",
        "delete_drawing_view",
        "add_drawing_dimension",
        "add_gtol",
        "add_datum",
        "add_surface_finish",
        "delete_annotation",
    },
}

ALWAYS_ON = "core"


def parse_env(raw: str = None) -> List[str]:
    """Requested section names, lowercased, `core` first and deduplicated.

    An empty or unset variable means "no filtering at all" and returns [].
    """
    if raw is None:
        raw = os.environ.get(ENV_VAR, "")
    names = [n.strip().lower() for n in raw.replace(";", ",").split(",")]
    names = [n for n in names if n]
    if not names:
        return []

    unknown = [n for n in names if n not in TOOLSETS]
    if unknown:
        logger.warning(
            f"{ENV_VAR}: unknown section(s) {unknown}; "
            f"known sections are {sorted(TOOLSETS)}"
        )
    known = [n for n in names if n in TOOLSETS]
    if not known:
        logger.warning(f"{ENV_VAR}: nothing usable in {raw!r}, keeping all tools")
        return []

    ordered = [ALWAYS_ON] + [n for n in known if n != ALWAYS_ON]
    return list(dict.fromkeys(ordered))


def allowed_names(sections: Iterable[str]) -> Set[str]:
    """Tool names covered by the given sections."""
    allowed: Set[str] = set()
    for name in sections:
        allowed |= TOOLSETS.get(name, set())
    return allowed


def filter_tools(tools: list, raw_env: str = None) -> list:
    """Filter a list of MCP Tool objects by the configured toolsets.

    Keeps anything this table does not know about, so adding a tool without
    classifying it degrades to "always visible" rather than "invisible".
    """
    sections = parse_env(raw_env)
    if not sections:
        return tools

    allowed = allowed_names(sections)
    classified = set().union(*TOOLSETS.values())

    kept, dropped, unclassified = [], [], []
    for tool in tools:
        if tool.name in allowed:
            kept.append(tool)
        elif tool.name not in classified:
            unclassified.append(tool.name)
            kept.append(tool)
        else:
            dropped.append(tool.name)

    if unclassified:
        logger.warning(
            f"tools missing from toolsets.py, kept by default: {sorted(unclassified)}"
        )
    logger.info(
        f"{ENV_VAR}={','.join(sections)} -> {len(kept)}/{len(tools)} tools "
        f"(dropped {len(dropped)})"
    )
    return kept
