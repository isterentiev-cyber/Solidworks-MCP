#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Live smoke test: real tools through server.call_tool against a running
SolidWorks. Complements selftest.py (which never touches COM).

Covers the late-bound calling conventions end to end: connect, open (part and
assembly, lightweight resolve), Save As, close, get_assembly_tree, a feature
chain on a fresh part (sketch -> extrude -> hole -> fillet -> chamfer ->
edit_feature), read-back tools and a drawing with PDF export.

    python scripts/live_smoke.py --out D:/tmp/smoke --part X.sldprt --asm Y.sldasm
    python scripts/live_smoke.py ... --preload-typed

--preload-typed imports the makepy module for the SolidWorks typelib into this
process BEFORE the server connects, if it exists in the gen_py cache (it never
generates one). That is the worst case the late-bound design must survive:
win32com.client.Dispatch would hand out typed wrappers from then on, the
server must not care.

Only writes into --out (Save As copies with a _test suffix, a PDF, JSON).
Closes every document it opened; never saves over an input file.
Exit code 0 = all steps passed.
"""

import argparse
import asyncio
import json
import os
import shutil
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))


def gen_py_state():
    """Where the win32com cache is and which SolidWorks modules are in it /
    imported into this process."""
    import win32com
    path = Path(win32com.__gen_path__)
    on_disk = sorted(p.name for p in path.glob("*.py") if p.name != "__init__.py") if path.exists() else []
    loaded = sorted(m for m in sys.modules if m.startswith("win32com.gen_py."))
    return {"gen_path": str(path), "on_disk": on_disk, "loaded": loaded}


def preload_typed():
    """Import (not generate) the cached makepy module for SldWorks, if any."""
    import importlib
    st = gen_py_state()
    mods = [m[:-3] for m in st["on_disk"] if m.upper().startswith("83A33D31")]
    for m in mods:
        importlib.import_module("win32com.gen_py." + m)
    return mods


class Run:
    def __init__(self):
        self.rows = []

    def step(self, label, tool, args=None, expect_ok=True, check=None):
        from solidworks_mcp import server
        t0 = time.time()
        try:
            out = asyncio.run(server.call_tool(tool, args or {}))
            text = "\n".join(c.text for c in out)
        except Exception as e:  # call_tool itself should never raise
            text = f"[RAISED] {e!r}"
        dt = time.time() - t0
        ok = text.startswith("[OK]") == expect_ok
        if ok and check:
            try:
                why = check(text)
                if why:
                    ok, text = False, f"{text}\n  CHECK FAILED: {why}"
            except Exception as e:
                ok, text = False, f"{text}\n  CHECK RAISED: {e!r}"
        self.rows.append((label, ok, dt, text))
        first = text.splitlines()[0] if text else ""
        print(f"{'PASS' if ok else 'FAIL'}  {dt:6.1f}s  {label:<28} {first[:150]}")
        if not ok:
            for ln in text.splitlines()[1:8]:
                print(f"{'':42}{ln[:150]}")
        return text

    def active_title(self):
        from solidworks_mcp import server, ext
        try:
            return ext.v(ext.model(server.sw_automation.app), "GetTitle")
        except Exception:
            return None

    def close(self, label, expected_title):
        """close_document only if the active document is the one this test
        opened -- never someone else's."""
        act = self.active_title()
        if not expected_title or act != expected_title:
            self.rows.append((label, False, 0.0, f"skipped: active is {act!r}, expected {expected_title!r}"))
            print(f"FAIL  {0:6.1f}s  {label:<28} skipped close: active {act!r} != {expected_title!r}")
            return
        self.step(label, "close_document")

    def extra(self, label, fn):
        t0 = time.time()
        try:
            why = fn()
            ok = not why
            text = why or "ok"
        except Exception as e:
            ok, text = False, repr(e)
        dt = time.time() - t0
        self.rows.append((label, ok, dt, text))
        print(f"{'PASS' if ok else 'FAIL'}  {dt:6.1f}s  {label:<28} {text[:150]}")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, help="output folder (created)")
    ap.add_argument("--part", help="an existing .sldprt to open (read-only use, Save As copy only)")
    ap.add_argument("--asm", help="an existing .sldasm to open")
    ap.add_argument("--preload-typed", action="store_true")
    ap.add_argument("--tag", default="", help="suffix for output file names")
    a = ap.parse_args(argv)

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    tag = f"_{a.tag}" if a.tag else ""

    if a.preload_typed:
        mods = preload_typed()
        print(f"preloaded typed modules: {mods or 'NONE (cache empty)'}")
        # control: with the module loaded, the NON-dynamic API really does hand
        # out typed wrappers in this process (what broke open_document before)
        import win32com.client
        probe = win32com.client.GetObject(Class="SldWorks.Application")
        print(f"control: win32com.client.GetObject(Class=...) -> {type(probe).__module__}.{type(probe).__name__}")
        del probe
    print("gen_py before server import:", json.dumps(gen_py_state()))

    from solidworks_mcp import server, ext
    from win32com.client import dynamic
    r = Run()

    r.step("connect", "connect_solidworks")
    r.extra("app is late-bound", lambda: None if isinstance(server.sw_automation.app, dynamic.CDispatch)
            else f"app wrapper is {type(server.sw_automation.app)}")
    before_docs = r.step("list_open_documents", "list_open_documents")

    # ---- new part: feature chain ---------------------------------------
    r.step("create_new_part", "create_new_part")
    t_part = r.active_title()
    r.step("create_sketch Top", "create_sketch", {"plane": "Top"})
    r.step("draw_rectangle", "draw_rectangle", {"x1": -50, "y1": -25, "x2": 50, "y2": 25})
    r.step("sketch_entities", "sketch_entities")
    r.step("extrude_sketch 20", "extrude_sketch", {"depth": 20})

    def top_z():
        md = ext.model(server.sw_automation.app)
        box = ext.v(ext.T(md, "IPartDoc"), "GetPartBox", True)
        return box[5] * 1000 if abs(box[5]) >= abs(box[2]) else box[2] * 1000
    z = top_z()
    r.step("hole drill 6", "hole", {"kind": "drill", "x": 20, "y": 0, "z": z, "diameter": 6})
    r.step("list_edges", "list_edges", {"max_count": 8, "type": "LINE"})
    r.step("fillet_edges 1", "fillet_edges", {"edge_indices": "1", "radius": 2})
    r.step("chamfer_edges 1", "chamfer_edges", {"edge_indices": "1", "distance": 1})
    r.step("list_faces", "list_faces", {"max_count": 5})
    r.step("get_parameters", "get_parameters")
    feat = None
    try:
        md = ext.model(server.sw_automation.app)
        names = [f.Name for f in ext.user_features(md)]
        feat = next((n for n in names if "xtru" in n or "拉伸" in n), names[0] if names else None)
    except Exception as e:
        print("feature lookup failed:", e)
    if feat:
        r.step(f"edit_feature {feat}", "edit_feature", {"feature": feat, "dimensions": "D1=25"})
    r.step("get_rebuild_errors", "get_rebuild_errors")
    r.step("get_sketch_status", "get_sketch_status")
    r.step("list_features", "list_features")
    r.step("inspect", "inspect")
    r.step("mass_properties", "mass_properties")
    r.step("capture_view", "capture_view", {"view": "isometric", "output_path": str(out / f"smoke_part{tag}.png"),
                                            "width": 640, "height": 400})
    new_part = out / f"smoke_part{tag}_test.SLDPRT"
    r.step("save_document (Save As)", "save_document", {"filepath": str(new_part)},
           check=lambda t: None if new_part.exists() else "file not written")
    r.step("save_document (in place)", "save_document")
    t_part = r.active_title()
    r.close("close_document", t_part)
    r.step("open saved part", "open_document", {"filepath": str(new_part)})
    t_part = r.active_title()

    # ---- drawing of the saved part --------------------------------------
    pdf = out / f"smoke_drawing{tag}.pdf"
    r.close("close_document", t_part)
    r.step("create_drawing", "create_drawing")
    t_drw = r.active_title()
    r.step("add_standard_views", "add_standard_views", {"model_path": str(new_part)},
           check=lambda t: None if r.active_title() and "Sheet" in r.active_title()
           else f"focus not on the drawing afterwards: {r.active_title()!r}")
    # SW renames the drawing after its model once views exist -- follow the
    # rename only if the active document is still a drawing sheet
    if "Sheet" in (r.active_title() or ""):
        t_drw = r.active_title()
    r.step("add_note", "add_note", {"text": "late-bound smoke", "x": 30, "y": 30})
    r.step("export_pdf", "export_pdf", {"path": str(pdf)},
           check=lambda t: None if pdf.exists() else "pdf not written")
    r.close("close drawing", t_drw)

    # ---- existing part --------------------------------------------------
    if a.part:
        src = Path(a.part)
        r.step("open part", "open_document", {"filepath": str(src)})
        t_src = r.active_title()
        r.step("inspect part", "inspect")
        r.step("mass_properties part", "mass_properties")
        dst = out / f"{src.stem}{tag}_test{src.suffix}"
        r.step("Save As part _test", "save_document", {"filepath": str(dst)},
               check=lambda t: None if dst.exists() else "file not written")
        r.close("close part", r.active_title() if (r.active_title() or "").startswith(dst.stem) else t_src)

    # ---- assembly -------------------------------------------------------
    if a.asm:
        r.step("open assembly", "open_document", {"filepath": a.asm},
               check=lambda t: None if "components" in t else "no component counts in answer")
        t_asm = r.active_title()
        r.step("get_document_info", "get_document_info")
        r.step("tree depth 2", "get_assembly_tree", {"max_depth": 2})
        r.step("tree flat", "get_assembly_tree", {"mode": "flat", "max_lines": 12})
        js = out / f"smoke_tree{tag}.json"
        r.step("tree details + json", "get_assembly_tree",
               {"details": True, "max_lines": 5, "output_path": str(js)},
               check=lambda t: None if js.exists() else "json not written")
        r.step("get_rebuild_errors asm", "get_rebuild_errors")
        r.close("close assembly", t_asm)

    after_docs = r.step("list_open_documents", "list_open_documents")
    print("gen_py after run:", json.dumps(gen_py_state()))

    failed = [x for x in r.rows if not x[1]]
    total = sum(x[2] for x in r.rows)
    print(f"\n{len(r.rows) - len(failed)}/{len(r.rows)} passed, {total:.1f}s")
    with open(out / f"smoke{tag}.json", "w", encoding="utf-8") as fh:
        json.dump({"gen_py": gen_py_state(), "preload_typed": a.preload_typed,
                   "steps": [{"label": l, "ok": ok, "s": round(dt, 2), "text": t} for l, ok, dt, t in r.rows]},
                  fh, ensure_ascii=False, indent=1)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
