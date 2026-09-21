"""
SolidWorks MCP -- extended tools (hot-reloadable)
-------------------------------------------------
Everything added on top of the upstream server lives here so that
`reload_api` can pick up edits without restarting the MCP process
(tool *schemas* still need a client restart; handler code does not).

Design rules (see CLAUDE.md / NOTES.md):
- All COM objects go through `T(obj, "IInterface")` -- the makepy-typed
  wrapper from utils/typelib.py. With typed wrappers every member is
  deterministically a method or a property (per the typelib), which
  removes the property-vs-method guessing that broke dynamic dispatch.
- All user-facing numbers are millimetres / degrees. SW internals are
  metres / radians.
- Results are compact single-line text: this output is read by a model,
  every token counts.
"""

import math
import time
import types
import logging
import traceback
from typing import Dict, List, Optional

import pythoncom
import win32com.client

from .utils.typelib import get_sw_module, get_constant

logger = logging.getLogger(__name__)


# ============================================================================
# Typed COM access
# ============================================================================

def T(obj, iface: str):
    """Wrap a COM object into its makepy-typed interface class.

    QueryInterface first: SW objects implement several interfaces (a sketch
    object is also its IFeature), and invoking one interface's dispids on
    another interface's IDispatch returns garbage (an int where a dispatch
    was expected). Falls back to a plain wrap if the object refuses the QI."""
    if obj is None:
        return None
    cls = getattr(get_sw_module(), iface)
    ole = getattr(obj, "_oleobj_", obj)
    try:
        ole = ole.QueryInterface(cls.CLSID, pythoncom.IID_IDispatch)
    except pythoncom.com_error:
        pass
    return cls(ole)


def v(obj, name: str, *args):
    """Read a member of a *typed* wrapper: call it if the typelib says it is
    a method, return it as-is if it is a property."""
    val = getattr(obj, name)
    return val(*args) if isinstance(val, types.MethodType) else val


def empty_dispatch():
    return win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)


def select_by_id(ext_obj, name, typ, x=0.0, y=0.0, z=0.0, append=False, mark=0) -> bool:
    """IModelDocExtension.SelectByID2 that works on both wrapper kinds: the
    typed one wants Callout=None, the dynamic one a VT_DISPATCH VARIANT
    (each rejects the other's form)."""
    last = None
    for callout in (None, empty_dispatch()):
        try:
            return bool(ext_obj.SelectByID2(name, typ, float(x), float(y), float(z),
                                            append, mark, callout, 0))
        except (TypeError, pythoncom.com_error) as e:
            # typed: TypeError on the VARIANT; dynamic: DISP_E_TYPEMISMATCH on None
            last = e
    raise ExtError(f"SelectByID2({name!r}, {typ}) failed for both callout forms: {last}")


class ExtError(Exception):
    pass


def _ok(message: str, data: Optional[Dict] = None) -> Dict:
    r = {"success": True, "message": message, "error_code": 0, "error_name": "swSuccess"}
    if data:
        r["data"] = data
    return r


def _err(message: str, code: int = 999, name: str = "swUnknownError") -> Dict:
    return {"success": False, "message": message, "error_code": code, "error_name": name}


def model(app):
    """Active document as typed IModelDoc2 (raises ExtError if none)."""
    if app is None:
        raise ExtError("Not connected to SolidWorks")
    doc = app.ActiveDoc
    if doc is None:
        raise ExtError("No active document")
    return T(doc, "IModelDoc2")


def save_in_place(md):
    """IModelDoc2.Save3 -> (ok, errors, warnings). Errors/Warnings are
    by-ref out params: the typed wrapper returns them in a tuple; the
    upstream dynamic call passed plain ints -> 'Type mismatch' (arg 2)."""
    r = md.Save3(1, 0, 0)  # swSaveAsOptions_Silent
    if isinstance(r, tuple):
        return bool(r[0]), int(r[1]), int(r[2])
    return bool(r), 0, 0


def is_part(md) -> bool:
    return md.GetType() == 1


# ============================================================================
# Formatting
# ============================================================================

def fnum(x: float, nd: int = 2) -> str:
    s = f"{x:.{nd}f}".rstrip("0").rstrip(".")
    return "0" if s in ("-0", "") else s


def fpt(p, scale: float = 1000.0, nd: int = 2) -> str:
    return "(" + ",".join(fnum(c * scale, nd) for c in p[:3]) + ")"


def fdir(d) -> str:
    return "(" + ",".join(fnum(c, 3) for c in d[:3]) + ")"


def parse_indices(s) -> List[int]:
    if isinstance(s, (list, tuple)):
        return [int(x) for x in s]
    return [int(x) for x in str(s).replace(";", ",").split(",") if x.strip()]


def parse_names(s) -> List[str]:
    if isinstance(s, (list, tuple)):
        return [str(x).strip() for x in s]
    return [x.strip() for x in str(s).split(",") if x.strip()]


# ============================================================================
# Features / bodies
# ============================================================================

def iter_features(md):
    f = md.FirstFeature()
    while f is not None:
        f = T(f, "IFeature")
        yield f
        f = f.GetNextFeature()


def user_features(md):
    """Features after the Origin (skips the fixed folders at the top)."""
    started = False
    for f in iter_features(md):
        if started:
            yield f
        elif v(f, "GetTypeName2") == "OriginProfileFeature":
            started = True


def find_feature(md, name: str):
    for f in iter_features(md):
        if f.Name == name:
            return f
    raise ExtError(f"Feature '{name}' not found")


def default_planes(md) -> Dict[str, str]:
    """Names of the three default planes, locale-independent: they are the
    first three RefPlane features, always in Front/Top/Right order."""
    names = [f.Name for f in iter_features(md) if v(f, "GetTypeName2") == "RefPlane"][:3]
    if len(names) < 3:
        raise ExtError("Default planes not found in feature tree")
    return {"front": names[0], "top": names[1], "right": names[2]}


def resolve_plane(md, plane: str) -> str:
    key = (plane or "front").strip().lower()
    for suffix in (" plane", "plane"):
        if key.endswith(suffix):
            key = key[: -len(suffix)].strip()
    dp = default_planes(md)
    return dp.get(key, plane)


def bodies(md) -> list:
    if not is_part(md):
        return []
    arr = T(md, "IPartDoc").GetBodies2(0, True) or ()
    return [T(b, "IBody2") for b in arr]


def volume_mm3(md) -> float:
    if not bodies(md):
        return 0.0
    mp = v(md.Extension, "CreateMassProperty")
    if mp is None:
        return 0.0
    return T(mp, "IMassProperty").Volume * 1e9


# ============================================================================
# Snapshot / auto-report
# ============================================================================

def snapshot(app) -> Optional[Dict]:
    """Cheap state fingerprint of the active part (None if not a part)."""
    try:
        md = model(app)
        if not is_part(md):
            return None
        bs = bodies(md)
        return {
            "V": volume_mm3(md),
            "B": len(bs),
            "F": sum(b.GetFaceCount() for b in bs),
            "E": sum(b.GetEdgeCount() for b in bs),
            "feats": [f.Name for f in user_features(md)],
        }
    except Exception as e:
        logger.debug(f"snapshot failed: {e}")
        return None


def feature_errors(md, names: List[str]) -> List[str]:
    out = []
    for n in names:
        try:
            r = find_feature(md, n).GetErrorCode2()
            code, warn = (r if isinstance(r, tuple) else (r, False))
            if code:
                out.append(f"{'warning' if warn else 'ERROR'} in {n}: code {code}")
        except Exception:
            pass
    return out


_STATUS = {1: "unknown", 2: "UNDER-defined", 3: "fully defined", 4: "OVER-defined",
           5: "no solution", 6: "invalid solution"}


def sketch_status(sk) -> int:
    return T(sk, "ISketch").GetConstrainedStatus()


def underdefined_parents(md, names: List[str]) -> List[str]:
    """Sketches that new features consume but that are not fully defined.
    A parametric model needs every driving sketch fully defined; building on
    'drawn where it happens to be' geometry is the mistake this flags."""
    out, seen = [], set()
    for n in names:
        try:
            f = find_feature(md, n)
        except ExtError:
            continue
        subs = []
        sub = f.GetFirstSubFeature()
        while sub is not None:
            subs.append(T(sub, "IFeature"))
            sub = T(sub, "IFeature").GetNextSubFeature()
        for sf in subs + [T(x, "IFeature") for x in (f.GetParents() or ())]:
            if v(sf, "GetTypeName2") != "ProfileFeature" or sf.Name in seen:
                continue
            seen.add(sf.Name)
            st = sketch_status(sf.GetSpecificFeature2())
            if st != 3:
                out.append(f"{sf.Name} is {_STATUS.get(st, st)} (sketch_entities / add_sketch_dimension / add_sketch_relation)")
    return out


def delta(app, before: Optional[Dict], after: Optional[Dict], expect: Optional[str] = None) -> str:
    """'Cut-Extrude1 | V 7091.8 mm³ (-503.4) | B1 F9 E15' plus warnings.
    expect: 'add' / 'remove' / None -- what the volume is supposed to do."""
    if before is None or after is None:
        return ""
    new = [n for n in after["feats"] if n not in before["feats"]]
    gone = [n for n in before["feats"] if n not in after["feats"]]
    dv = after["V"] - before["V"]
    parts = [",".join(new) if new else "no new feature"]
    if gone:
        parts[0] += f" (removed: {','.join(gone)})"
    parts.append(f"V {fnum(after['V'], 1)} mm³ ({dv:+.1f})")
    parts.append(f"B{after['B']} F{after['F']} E{after['E']}")
    warn = []
    eps = max(1e-3, 1e-6 * abs(before["V"]))
    if expect == "add" and dv <= eps:
        warn.append("volume did not grow")
    if expect == "remove" and dv >= -eps:
        warn.append("nothing was removed")
    if after["B"] > before["B"] and before["B"] > 0:
        warn.append(f"body count grew {before['B']}->{after['B']} (split or disjoint?)")
    try:
        md_ = model(app)
        warn += feature_errors(md_, new)
        warn += underdefined_parents(md_, new)
    except Exception:
        pass
    s = " | ".join(parts)
    if warn:
        s += " | ⚠ " + "; ".join(warn)
    return s


# ============================================================================
# Topology tables
# ============================================================================

_SURF = {4001: "PLANE", 4002: "CYL", 4003: "CONE", 4004: "SPHERE", 4005: "TORUS",
         4006: "BSURF", 4007: "BLEND", 4008: "OFFSET", 4009: "EXTRU", 4010: "SREV"}
_CURVE = {3001: "LINE", 3002: "CIRC", 3003: "ELLIPSE", 3004: "INTERSECT", 3005: "BCURVE",
          3006: "SPCURVE", 3008: "CONSTPARAM", 3009: "TRIMMED"}


def face_rows(md) -> List[Dict]:
    rows, i = [], 0
    for bi, b in enumerate(bodies(md)):
        for fo in b.GetFaces() or ():
            i += 1
            f = T(fo, "IFace2")
            s = T(f.GetSurface(), "ISurface")
            st = s.Identity()
            box = v(f, "GetBox")
            c = [(box[k] + box[k + 3]) / 2 for k in range(3)]
            p = list(f.GetClosestPointOn(*c)[:3])
            extra = ""
            if st == 4001:
                extra = "n=" + fdir(v(f, "Normal"))
            elif st == 4002:
                cp = s.CylinderParams
                extra = f"r={fnum(cp[6] * 1000)} axis={fdir(cp[3:6])} through {fpt(cp[0:3])}"
            elif st == 4003:
                cp = s.ConeParams
                extra = f"r={fnum(cp[6] * 1000)} half-angle={fnum(math.degrees(cp[7]), 1)}° axis={fdir(cp[3:6])}"
            elif st == 4004:
                sp = s.SphereParams
                extra = f"r={fnum(sp[3] * 1000)} c={fpt(sp[0:3])}"
            elif st == 4005:
                tp = s.TorusParams
                extra = f"R={fnum(tp[6] * 1000)} r={fnum(tp[7] * 1000)} c={fpt(tp[0:3])} axis={fdir(tp[3:6])}"
            rows.append({"i": i, "body": bi + 1, "type": _SURF.get(st, str(st)),
                         "area": f.GetArea() * 1e6, "p": p, "extra": extra, "obj": f})
    return rows


def edge_rows(md) -> List[Dict]:
    rows, i = [], 0
    for bi, b in enumerate(bodies(md)):
        for eo in b.GetEdges() or ():
            i += 1
            e = T(eo, "IEdge")
            c = T(e.GetCurve(), "ICurve")
            cp = T(e.GetCurveParams3(), "ICurveParamData")
            u0, u1 = cp.UMinValue, cp.UMaxValue
            length = c.GetLength3(u0, u1) * 1000
            ct = c.Identity()
            a, bpt = cp.StartPoint, cp.EndPoint
            if ct == 3001:
                mid = [(a[k] + bpt[k]) / 2 for k in range(3)]
            else:
                # the curve's parameter range does not always match the
                # trimmed edge (reversed sense) -- project back onto the edge
                guess = c.Evaluate2((u0 + u1) / 2, 0)[:3]
                mid = list(e.GetClosestPointOn(*guess)[:3])
            extra = ""
            if ct == 3001:
                d = [bpt[k] - a[k] for k in range(3)]
                n = math.sqrt(sum(x * x for x in d)) or 1
                extra = "dir=" + fdir([x / n for x in d])
            elif ct == 3002:
                cc = c.CircleParams
                full = abs(length - 2 * math.pi * cc[6] * 1000) < 1e-3
                extra = (f"r={fnum(cc[6] * 1000)} c={fpt(cc[0:3])} n={fdir(cc[3:6])}"
                         + (" full" if full else ""))
            rows.append({"i": i, "body": bi + 1, "type": _CURVE.get(ct, str(ct)),
                         "len": length, "mid": mid, "extra": extra, "obj": e})
    return rows


def fmt_face(r) -> str:
    return f"[{r['i']}] {r['type']} A={fnum(r['area'], 1)} p={fpt(r['p'])} {r['extra']}".rstrip()


def fmt_edge(r) -> str:
    return f"[{r['i']}] {r['type']} L={fnum(r['len'])} mid={fpt(r['mid'])} {r['extra']}".rstrip()


def _match(pt, x, y, z, use, tol):
    target = (x, y, z)
    return all(abs(pt[k] * 1000 - target[k]) <= tol for k in range(3) if use[k])


def select_face_at(md, p_mm, append: bool = False, mark: int = 0) -> Dict:
    """Select the face that contains model point p (mm), keeping p as the
    selection point (Hole Wizard places the hole there).

    SelectByID2("", "FACE", x, y, z) returns False in this install even for
    a point exactly on a planar face, while SelectByRay works -- so: find the
    face whose closest point is p, shoot a short ray at p along -normal."""
    p = [c / 1000 for c in p_mm]
    best, best_d = None, 1e9
    for r in face_rows(md):
        q = r["obj"].GetClosestPointOn(*p)[:3]
        d = math.dist(p, q)
        if d < best_d:
            best, best_d = r, d
    if best is None or best_d > 1e-5:
        raise ExtError(f"No face through {p_mm} mm (nearest is {fnum(best_d * 1000, 3)} mm away)")
    if best["type"] == "PLANE":
        n = list(v(best["obj"], "Normal"))
    else:
        ev = T(best["obj"].GetSurface(), "ISurface").EvaluateAtPoint(*p)
        n = list(ev[3:6]) if ev else [0.0, 0.0, 1.0]
    o = [p[k] + n[k] * 0.001 for k in range(3)]
    if not append:
        md.ClearSelection2(True)
    ok = md.Extension.SelectByRay(o[0], o[1], o[2], -n[0], -n[1], -n[2], 0.0002, 2, append, mark, 0)
    if not ok:
        raise ExtError(f"SelectByRay missed face [{best['i']}] at {p_mm}")
    return best


def select_entities(md, kind: str, indices, append: bool = False, mark: int = 0) -> List[int]:
    rows = edge_rows(md) if kind == "edge" else face_rows(md)
    by_i = {r["i"]: r for r in rows}
    idx = parse_indices(indices)
    missing = [i for i in idx if i not in by_i]
    if missing:
        raise ExtError(f"No {kind}(s) {missing}; valid 1..{len(rows)}. Indices change after every feature -- re-list.")
    sm = T(md.SelectionManager, "ISelectionMgr")
    if not append:
        md.ClearSelection2(True)
    for n, i in enumerate(idx):
        data = T(sm.CreateSelectData(), "ISelectData")
        data.Mark = mark
        ok = T(by_i[i]["obj"], "IEntity").Select4(True if (append or n > 0) else False, data)
        if not ok:
            raise ExtError(f"Could not select {kind} {i}")
    return idx


# ============================================================================
# Sketch helpers
# ============================================================================

def xform(a, p) -> List[float]:
    """Apply an IMathTransform.ArrayData to a point. SW convention is a row
    vector: p' = (p · R) * scale + t, R = a[0:9] row-major, t = a[9:12].
    (IMathUtility.CreatePoint with a Python list gets garbage through the
    typed wrapper -- 1e-311 values -- so the math is done here.)"""
    s = a[12] if len(a) > 12 and a[12] else 1.0
    return [sum(p[i] * a[3 * i + j] for i in range(3)) * s + a[9 + j] for j in range(3)]


def sketch_to_model(sk) -> List[float]:
    """ArrayData of the active sketch's sketch->model transform."""
    return list(T(T(sk.ModelToSketchTransform, "IMathTransform").Inverse(), "IMathTransform").ArrayData)


def sketch_frame(app) -> str:
    """Where the active sketch's 2D axes point in model space (mm)."""
    try:
        md = model(app)
        sk = md.SketchManager.ActiveSketch
        if sk is None:
            return ""
        a = sketch_to_model(T(sk, "ISketch"))
        o, x, y = xform(a, [0, 0, 0]), xform(a, [1, 0, 0]), xform(a, [0, 1, 0])
        dx = [x[k] - o[k] for k in range(3)]
        dy = [y[k] - o[k] for k in range(3)]
        n = [dx[1] * dy[2] - dx[2] * dy[1], dx[2] * dy[0] - dx[0] * dy[2], dx[0] * dy[1] - dx[1] * dy[0]]
        return f"sketch frame: origin={fpt(o)} X→{fdir(dx)} Y→{fdir(dy)} normal→{fdir(n)}"
    except Exception as e:
        logger.debug(f"sketch_frame failed: {e}")
        return ""


class NoInference:
    """Disable sketch auto-relations while drawing via API. Without this,
    SW snaps new points to existing geometry/axes and coordinates silently
    drift; CreateCircle/CreateCornerRectangle even return None (NOTES.md)."""

    def __init__(self, sm):
        self.sm = sm

    def __enter__(self):
        self.sm.AddToDB = True
        self.sm.AutoInference = False
        return self.sm

    def __exit__(self, *exc):
        self.sm.AddToDB = False
        self.sm.AutoInference = True
        return False


def sketch_manager(md):
    sm = T(md.SketchManager, "ISketchManager")
    if sm.ActiveSketch is None:
        raise ExtError("No active sketch -- call create_sketch first")
    return sm


def draw_profile(md, points, closed: bool = True, construction: bool = False) -> int:
    pts = points
    if isinstance(pts, str):
        import json as _json
        pts = _json.loads(pts)
    if len(pts) < 2:
        raise ExtError("Need at least 2 points")
    sm = sketch_manager(md)
    segs = list(zip(pts, pts[1:]))
    if closed and len(pts) > 2 and list(pts[0]) != list(pts[-1]):
        segs.append((pts[-1], pts[0]))
    made = 0
    with NoInference(sm):
        for a, b in segs:
            fn = sm.CreateCenterLine if construction else sm.CreateLine
            if fn(a[0] / 1000, a[1] / 1000, 0, b[0] / 1000, b[1] / 1000, 0) is not None:
                made += 1
    if made != len(segs):
        raise ExtError(f"Only {made}/{len(segs)} segments created (zero-length or duplicate segment?)")
    return made


# ============================================================================
# Reference geometry
# ============================================================================

def _plane_normal(md, plane_feat) -> List[float]:
    rp = T(plane_feat.GetSpecificFeature2(), "IRefPlane")
    a = T(rp.Transform, "IMathTransform").ArrayData
    return [a[6], a[7], a[8]]


def world_axis(md, axis: str) -> str:
    """Name of a (hidden) reference axis along world X/Y/Z through the
    origin, created on first use from two default planes."""
    axis = axis.upper()
    name = f"MCP_Axis_{axis}"
    for f in iter_features(md):
        if f.Name == name:
            return name
    k = "XYZ".index(axis)
    planes = [f for f in iter_features(md) if v(f, "GetTypeName2") == "RefPlane"][:3]
    # the axis lies in both planes whose normal is perpendicular to it
    pick = [p for p in planes if abs(_plane_normal(md, p)[k]) < 1e-6]
    if len(pick) != 2:
        raise ExtError(f"Could not find two default planes containing world {axis}")
    md.ClearSelection2(True)
    pick[0].Select2(False, 0)
    pick[1].Select2(True, 0)
    if not md.InsertAxis2(True):
        raise ExtError("InsertAxis2 failed")
    last = list(user_features(md))[-1]
    last.Name = name
    md.ClearSelection2(True)
    last.Select2(False, 0)
    md.BlankRefGeom()
    md.ClearSelection2(True)
    return name


def select_direction(md, spec: str, mark: int, append: bool = True):
    """spec: 'X'/'Y'/'Z' (world axis), 'edge:N', 'face:N' (cyl face -> its
    axis), or any feature name (axis, plane, sketch line...)."""
    s = str(spec).strip()
    if s.upper() in ("X", "Y", "Z"):
        find_feature(md, world_axis(md, s)).Select2(append, mark)
        return
    if ":" in s:
        kind, idx = s.split(":", 1)
        select_entities(md, kind.strip().lower(), idx, append=append, mark=mark)
        return
    find_feature(md, s).Select2(append, mark)


def select_features(md, names, mark: int, append: bool = True):
    for n, name in enumerate(parse_names(names)):
        if not find_feature(md, name).Select2(append or n > 0, mark):
            raise ExtError(f"Could not select feature '{name}'")


# ============================================================================
# Tool implementations
# ============================================================================

def t_inspect(app, args):
    md = model(app)
    title = md.GetTitle()
    lines = [f"{title} | {md.GetPathName() or '<unsaved>'}"]
    if is_part(md):
        bs = bodies(md)
        vol = volume_mm3(md)
        area = 0.0
        if bs:
            area = T(v(md.Extension, "CreateMassProperty"), "IMassProperty").SurfaceArea * 1e6
        lines.append(f"bodies {len(bs)} | V {fnum(vol, 1)} mm³ | A {fnum(area, 1)} mm²")
        for bi, b in enumerate(bs):
            bx = b.GetBodyBox()
            size = [bx[k + 3] - bx[k] for k in range(3)]
            fr = [T(f, "IFace2") for f in b.GetFaces() or ()]
            kinds = {}
            for f in fr:
                t = _SURF.get(T(f.GetSurface(), "ISurface").Identity(), "?")
                kinds[t] = kinds.get(t, 0) + 1
            lines.append(
                f"body{bi + 1}: bbox X[{fnum(bx[0]*1000)}..{fnum(bx[3]*1000)}] "
                f"Y[{fnum(bx[1]*1000)}..{fnum(bx[4]*1000)}] Z[{fnum(bx[2]*1000)}..{fnum(bx[5]*1000)}] "
                f"size {'×'.join(fnum(s*1000) for s in size)} | faces {b.GetFaceCount()} "
                f"({', '.join(f'{k}×{n}' for k, n in sorted(kinds.items()))}) | edges {b.GetEdgeCount()}")
        try:
            mat = T(md, "IPartDoc").GetMaterialPropertyName2("", "")
            mat = mat[1] if isinstance(mat, tuple) else mat
            if mat:
                lines.append(f"material: {mat}")
        except Exception:
            pass
    feats = []
    for f in user_features(md):
        tag = f"{f.Name}[{v(f, "GetTypeName2")}]"
        try:
            if v(f, "IsSuppressed"):
                tag += "(suppressed)"
        except Exception:
            pass
        feats.append(tag)
    lines.append(f"features ({len(feats)}): " + ", ".join(feats))
    sk = md.SketchManager.ActiveSketch
    if sk is not None:
        lines.append("ACTIVE SKETCH open; " + sketch_frame(app))
    try:
        em = T(md.GetEquationMgr(), "IEquationMgr")
        if em.GetCount():
            lines.append(f"equations: {em.GetCount()} (get_parameters for details)")
    except Exception:
        pass
    return _ok("\n".join(lines))


def t_list_faces(app, args):
    md = model(app)
    rows = [r for r in face_rows(md) if r["area"] >= float(args.get("min_area", 0))]
    t = args.get("type")
    if t:
        rows = [r for r in rows if r["type"] == t.upper()]
    n = int(args.get("max_count", 60))
    out = [fmt_face(r) for r in rows[:n]]
    if len(rows) > n:
        out.append(f"... {len(rows) - n} more (raise max_count / min_area)")
    return _ok(f"{len(rows)} faces (p = point ON the face, mm)\n" + "\n".join(out))


def t_list_edges(app, args):
    md = model(app)
    rows = [r for r in edge_rows(md) if r["len"] >= float(args.get("min_length", 0))]
    t = args.get("type")
    if t:
        rows = [r for r in rows if r["type"] == t.upper()]
    n = int(args.get("max_count", 60))
    out = [fmt_edge(r) for r in rows[:n]]
    if len(rows) > n:
        out.append(f"... {len(rows) - n} more (raise max_count / min_length)")
    return _ok(f"{len(rows)} edges\n" + "\n".join(out))


def _find(app, args, kind):
    md = model(app)
    x, y, z = (float(args.get(k, 0)) for k in "xyz")
    use = [bool(args.get(f"use_{k}", False)) for k in "xyz"]
    if not any(use):
        raise ExtError("Set at least one of use_x/use_y/use_z")
    tol = float(args.get("tolerance", 0.1))
    if kind == "face":
        rows = [r for r in face_rows(md) if _match(r["p"], x, y, z, use, tol)]
        t = args.get("type")
        if t:
            rows = [r for r in rows if r["type"] == t.upper()]
        return _ok(f"{len(rows)} match: " + ("; ".join(fmt_face(r) for r in rows) or "none"),
                   {"indices": [r["i"] for r in rows]})
    rows = [r for r in edge_rows(md) if _match(r["mid"], x, y, z, use, tol)]
    t = args.get("type")
    if t:
        rows = [r for r in rows if r["type"] == t.upper()]
    return _ok(f"{len(rows)} match: " + ("; ".join(fmt_edge(r) for r in rows) or "none"),
               {"indices": [r["i"] for r in rows]})


def t_find_face(app, args):
    return _find(app, args, "face")


def t_find_edge(app, args):
    return _find(app, args, "edge")


def t_select(app, args):
    md = model(app)
    kind = args.get("kind", "edge").lower()
    if kind not in ("edge", "face"):
        raise ExtError("kind must be 'edge' or 'face'")
    idx = select_entities(md, kind, args["indices"], bool(args.get("append", False)),
                          int(args.get("mark", 0)))
    return _ok(f"selected {kind}s {idx}")


def t_draw_profile(app, args):
    md = model(app)
    n = draw_profile(md, args["points"], bool(args.get("closed", True)),
                     bool(args.get("construction", False)))
    return _ok(f"{n} segments")


# ISO 261 coarse pitch -> tap drill diameter (ISO 2306), mm
_TAP_DRILL = {"M2": 1.6, "M2.5": 2.05, "M3": 2.5, "M4": 3.3, "M5": 4.2, "M6": 5.0,
              "M8": 6.8, "M10": 8.5, "M12": 10.2, "M14": 12.0, "M16": 14.0, "M18": 15.5,
              "M20": 17.5, "M22": 19.5, "M24": 21.0, "M27": 24.0, "M30": 26.5}


def t_hole(app, args):
    """Hole Wizard hole at a point on a planar face.

    HoleWizard5's Value1..12 meaning depends on the hole type and passing
    -1 ("default") is NOT safe everywhere -- found empirically on SW 2026:
      drill (swWzdHole):  size 'Ø6.0', Diameter, Depth, Value2 = drill angle,
                          other Values 0 (with -1 through-all returns None)
      tap   (swWzdTap):   Diameter = tap drill Ø, Depth = drill depth,
                          Value1 = THREAD DEPTH (0/-1 -> inch-default 25.4 mm
                          profile + warning; Value6 must stay -1), the thread
                          Ø comes from the size. HoleWizard5 always builds a
                          modeled Ø-major step ("remove thread", Type 31);
                          cosmetic thread = set Type 46/48 afterwards via
                          IWizardHoleFeatureData2 + ModifyDefinition. No
                          Value slot sets the drill-point angle for taps (it
                          comes out ~0.3°, a 7 mm needle cone) -- also fixed
                          there via DrillAngle.
      cbore (swWzdCounterBore): all -1 gives correct ISO 4762 seats.
    """
    md = model(app)
    pts = args.get("points")
    if isinstance(pts, str):
        import json as _json
        pts = _json.loads(pts)
    if not pts:
        pts = [[float(args["x"]), float(args["y"]), float(args["z"])]]
    kind = (args.get("kind") or "drill").lower()
    dia = float(args.get("diameter", 0) or 0)
    depth = float(args.get("depth", 0) or 0)
    through = depth <= 0
    size = (args.get("size") or "").strip()
    std = get_constant("swStandardISO")
    end = 1 if through else 0  # swEndCondThroughAll / swEndCondBlind

    if kind == "drill":
        if dia <= 0:
            raise ExtError("diameter required for kind=drill")
        htype, fast = get_constant("swWzdHole"), get_constant("swStandardISODrillSizes")
        ssize = size or f"Ø{dia:.1f}"
        d_arg, depth_arg = dia / 1000, (depth / 1000 if not through else 0.0)
        V = [0.0, math.radians(118)] + [0.0] * 10
    elif kind == "tap":
        size = size.upper()
        if size not in _TAP_DRILL:
            raise ExtError(f"kind=tap needs an ISO coarse size, one of {', '.join(_TAP_DRILL)}")
        htype, fast = get_constant("swWzdTap"), get_constant("swStandardISOTappedHole")
        ssize = size
        thread_d = float(size[1:])
        thread_depth = depth if not through else 0.0
        d_arg = _TAP_DRILL[size] / 1000
        depth_arg = (thread_depth + 0.5 * thread_d) / 1000 if not through else 0.0
        V = [(thread_depth if not through else 2 * thread_d) / 1000] + [-1.0] * 11
    elif kind == "cbore":
        if not size:
            raise ExtError("size required for kind=cbore, e.g. 'M6' (ISO 4762 socket head)")
        htype, fast = get_constant("swWzdCounterBore"), get_constant("swStandardISOSocketHeadCap")
        ssize = size.upper()
        d_arg, depth_arg = -1.0, (depth / 1000 if not through else 0.01)
        V = [-1.0] * 12
    else:
        raise ExtError("kind must be drill / tap / cbore")

    before = snapshot(app)
    select_face_at(md, pts[0])
    fm = T(md.FeatureManager, "IFeatureManager")
    feat = fm.HoleWizard5(htype, std, fast, ssize, end, d_arg, depth_arg, -1.0,
                          *V, "", False, True, True, True, True, False)
    if feat is None:
        raise ExtError(f"HoleWizard5 returned None (kind={kind}, size={ssize!r}). "
                       f"Drill sizes must exist in the ISO drill table ('Ø6.0', 'Ø6.6', ...).")
    feat = T(feat, "IFeature")
    if kind == "tap":
        # two separate edits: changing Type resets DrillAngle
        edits = []
        if (args.get("thread") or "cosmetic").lower() == "cosmetic":
            edits.append({"Type": get_constant("swTapThruCosmeticThread" if through else "swTapBlindCosmeticThread"),
                          "CosmeticThreadType": get_constant("swCosmeticThreadWithCallout")})
        edits.append({"DrillAngle": math.radians(118)})
        for ed in edits:
            data = T(feat.GetDefinition(), "IWizardHoleFeatureData2")
            data.AccessSelections(md, None)
            for k, val in ed.items():
                setattr(data, k, val)
            if not feat.ModifyDefinition(data, md, None):
                data.ReleaseSelectionAccess()
                logger.warning(f"could not update tapped hole definition {list(ed)}")
    msg = [f"{kind} {ssize}" + (" through all" if through else f" depth {fnum(depth)}")]
    if len(pts) > 1:
        added = _add_hole_points(md, feat, pts[1:])
        msg.append(f"+{added} positions")
    msg.append(delta(app, before, snapshot(app), "remove"))
    return _ok(" | ".join(m for m in msg if m))


def _add_hole_points(md, feat, pts_mm) -> int:
    """Add placement points to a Hole Wizard feature. Its FIRST sub-sketch
    holds the positions; the second one is the hole's revolve profile."""
    sub = feat.GetFirstSubFeature()
    sketch_feat = None
    while sub is not None:
        sub = T(sub, "IFeature")
        if "Profile" in v(sub, "GetTypeName2"):
            sketch_feat = sub
            break
        sub = sub.GetNextSubFeature()
    if sketch_feat is None:
        raise ExtError("Hole position sketch not found")
    md.ClearSelection2(True)
    sketch_feat.Select2(False, 0)
    md.EditSketch()
    sm = T(md.SketchManager, "ISketchManager")
    sk = T(sm.ActiveSketch, "ISketch")
    m2s = list(T(sk.ModelToSketchTransform, "IMathTransform").ArrayData)
    added = 0
    with NoInference(sm):
        for p in pts_mm:
            q = xform(m2s, [c / 1000 for c in p])
            if abs(q[2]) > 1e-6:
                logger.warning(f"hole point {p} is {q[2] * 1000:.3f} mm off the face plane")
            if sm.CreatePoint(q[0], q[1], 0) is not None:
                added += 1
    sm.InsertSketch(True)
    md.EditRebuild3()
    return added


def t_circular_pattern(app, args):
    md = model(app)
    before = snapshot(app)
    md.ClearSelection2(True)
    select_direction(md, args.get("axis", "Z"), mark=1, append=False)
    select_features(md, args["features"], mark=4, append=True)
    count = int(args.get("count", 4))
    angle = math.radians(float(args.get("angle", 360)))
    fm = T(md.FeatureManager, "IFeatureManager")
    feat = fm.FeatureCircularPattern5(count, angle, bool(args.get("reverse", False)), "NULL",
                                      bool(args.get("geometry_pattern", False)), True, False, False,
                                      False, False, 1, 0.0, "NULL", False)
    md.ClearSelection2(True)
    if feat is None:
        raise ExtError("FeatureCircularPattern5 returned None. Selection was fine in testing; the usual "
                       "cause is geometry: an instance landing exactly in/over an existing cut. "
                       "Try another count, or check the seed position.")
    return _ok(delta(app, before, snapshot(app)))


def t_linear_pattern(app, args):
    md = model(app)
    before = snapshot(app)
    md.ClearSelection2(True)
    select_direction(md, args.get("direction", "X"), mark=1, append=False)
    n2 = int(args.get("count2", 1))
    if n2 > 1:
        select_direction(md, args["direction2"], mark=2, append=True)
    select_features(md, args["features"], mark=4, append=True)
    fm = T(md.FeatureManager, "IFeatureManager")
    feat = fm.FeatureLinearPattern5(
        int(args.get("count", 2)), float(args.get("spacing", 10)) / 1000,
        n2, float(args.get("spacing2", 10)) / 1000,
        bool(args.get("reverse", False)), bool(args.get("reverse2", False)),
        "NULL", "NULL", bool(args.get("geometry_pattern", False)), False,
        False, False, False, False, False, False, False, False, 0.0, 0.0, False, False)
    md.ClearSelection2(True)
    if feat is None:
        raise ExtError("FeatureLinearPattern5 returned None (direction/feature selection?)")
    return _ok(delta(app, before, snapshot(app)))


def t_mirror(app, args):
    md = model(app)
    before = snapshot(app)
    md.ClearSelection2(True)
    plane = str(args.get("plane", "Right"))
    if plane.lower().startswith("face:"):
        select_entities(md, "face", plane.split(":", 1)[1], append=False, mark=2)
    else:
        find_feature(md, resolve_plane(md, plane)).Select2(False, 2)
    select_features(md, args["features"], mark=1, append=True)
    feat = T(md.FeatureManager, "IFeatureManager").InsertMirrorFeature2(False, False, False, False, 0)
    md.ClearSelection2(True)
    if feat is None:
        raise ExtError("InsertMirrorFeature2 returned None")
    return _ok(delta(app, before, snapshot(app)))


def t_get_parameters(app, args):
    md = model(app)
    lines = []
    em = T(md.GetEquationMgr(), "IEquationMgr")
    for i in range(em.GetCount()):
        eq = v(em, "Equation", i)
        val = v(em, "Value", i)
        gv = v(em, "GlobalVariable", i)
        lines.append(f"{'var' if gv else 'eq '} [{i}] {eq}  = {fnum(val, 4)}")
    flt = args.get("feature")
    for f in user_features(md):
        if flt and f.Name != flt:
            continue
        dims = []
        for dd in display_dims(f):
            dim = T(dd.GetDimension2(0), "IDimension")
            val = dim.SystemValue
            if dim.GetType() == 1:
                s = f"{fnum(math.degrees(val), 3)}°"
            else:
                s = fnum(val * 1000, 4)
            dims.append(f"{dim.FullName}={s}")
        if dims:
            lines.append(f"{f.Name}: " + ", ".join(dims))
    return _ok("\n".join(lines) if lines else "no equations / dimensions")


def _eq_add(em, eq: str) -> int:
    """Add an equation. IEquationMgr.Add3 returns -1 for every input in this
    install (even '"h" = 30mm'); the older Add2 works."""
    i = em.Add3(-1, eq, True, get_constant("swAllConfiguration"), None)
    if i < 0:
        i = em.Add2(-1, eq, True)
    return i


def _eq_set(em, i: int, eq: str):
    """Replace equation i. SetEquationAndConfigurationOption returns -1 and
    changes nothing in this install; the plain Equation setter works."""
    em.SetEquation(i, eq)
    if v(em, "Equation", i).replace(" ", "") != eq.replace(" ", ""):
        raise ExtError(f"Equation rejected: {eq} (still: {v(em, 'Equation', i)})")


def _eq_index(em, name: str) -> int:
    key = f'"{name}"'
    for i in range(em.GetCount()):
        if v(em, "Equation", i).strip().startswith(key):
            return i
    return -1


def t_set_parameter(app, args):
    md = model(app)
    name, expr = args["name"], str(args["expression"]).strip()
    em = T(md.GetEquationMgr(), "IEquationMgr")
    all_cfg = get_constant("swAllConfiguration")
    before = snapshot(app)
    if "@" in name:
        try:
            num = float(expr.replace(",", "."))
        except ValueError:
            num = None
        if num is not None:
            dim = md.Parameter(name)
            if dim is None:
                raise ExtError(f"Dimension '{name}' not found (get_parameters shows names)")
            dim = T(dim, "IDimension")
            val = math.radians(num) if dim.GetType() == 1 else num / 1000
            dim.SetSystemValue3(val, all_cfg, None)
        else:
            i = _eq_index(em, name)
            eq = f'"{name}" = {expr}'
            if i >= 0:
                _eq_set(em, i, eq)
            elif _eq_add(em, eq) < 0:
                raise ExtError(f"Equation rejected: {eq}")
    else:
        i = _eq_index(em, name)
        if i < 0:
            raise ExtError(f"Global variable '{name}' not found -- use add_parameter")
        _eq_set(em, i, f'"{name}" = {expr}')
    em.EvaluateAll()
    md.EditRebuild3()
    return _ok(f"{name} = {expr} | " + delta(app, before, snapshot(app)))


def t_add_parameter(app, args):
    md = model(app)
    name, expr = args["name"], str(args["expression"]).strip()
    em = T(md.GetEquationMgr(), "IEquationMgr")
    if _eq_index(em, name) >= 0:
        raise ExtError(f"'{name}' already exists -- use set_parameter")
    eq = f'"{name}" = {expr}'
    if _eq_add(em, eq) < 0:
        raise ExtError(f"Equation rejected: {eq}")
    em.EvaluateAll()
    i = _eq_index(em, name)
    return _ok(f"{eq} -> {fnum(v(em, 'Value', i), 4)}")


def t_delete_feature(app, args):
    md = model(app)
    before = snapshot(app)
    names = parse_names(args["names"])
    for n in reversed(names):
        md.ClearSelection2(True)
        find_feature(md, n).Select2(False, 0)
        if not md.Extension.DeleteSelection2(get_constant("swDelete_Absorbed")):
            raise ExtError(f"Could not delete '{n}'")
    md.ClearSelection2(True)
    return _ok(delta(app, before, snapshot(app)))


def t_suppress_feature(app, args):
    md = model(app)
    before = snapshot(app)
    state = get_constant("swSuppressFeature" if args.get("suppressed", True) else "swUnSuppressFeature")
    for n in parse_names(args["names"]):
        if not find_feature(md, n).SetSuppression2(state, get_constant("swThisConfiguration"), None):
            raise ExtError(f"SetSuppression2 failed on '{n}'")
    md.EditRebuild3()
    return _ok(delta(app, before, snapshot(app)))


# Transaction state survives reload_api via the server module (see server.py),
# but a module-level dict is enough within one process lifetime.
_TX: Dict = {}


def sketch_feature(md, sk):
    """The IFeature of a sketch. Casting via T(sk, "IFeature") does NOT work:
    the QI succeeds but the returned IDispatch still routes to ISketch's
    dispid table (IFeature.Name came back as a transform tuple). Match COM
    identity (IUnknown) against the tree instead."""
    u = sk._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
    for f in user_features(md):
        if v(f, "GetTypeName2") in ("ProfileFeature", "3DProfileFeature"):
            s2 = T(f.GetSpecificFeature2(), "ISketch")
            if s2._oleobj_.QueryInterface(pythoncom.IID_IUnknown) == u:
                return f
    raise ExtError("Sketch feature not found in tree")


def _active_sketch(md):
    sk = md.SketchManager.ActiveSketch
    if sk is None:
        raise ExtError("No active sketch (create_sketch first, or edit one)")
    return T(sk, "ISketch")


_SEG_TYPES = {0: "LINE", 1: "ARC", 2: "ELLIPSE", 3: "SPLINE", 4: "TEXT", 5: "PARABOLA"}


def sketch_items(md):
    """Addressable entities of the active sketch: S1.. segments, P1.. points,
    plus 'O' = the part origin. Coordinates are sketch 2D, mm."""
    sk = _active_sketch(md)
    segs = [T(x, "ISketchSegment") for x in (sk.GetSketchSegments() or ())]
    pts = [T(x, "ISketchPoint") for x in (sk.GetSketchPoints2() or ())]
    return sk, segs, pts


def _pt(q) -> str:
    return f"({fnum(q.X * 1000)},{fnum(q.Y * 1000)})"


def display_dims(f):
    """Iterate IFeature display dimensions. With no dimensions SW returns an
    int 0 instead of Nothing and the makepy wrapper crashes on it."""
    try:
        dd = f.GetFirstDisplayDimension()
    except AttributeError:
        return
    while dd is not None:
        dd = T(dd, "IDisplayDimension")
        yield dd
        try:
            dd = f.GetNextDisplayDimension(dd)
        except AttributeError:
            return


def _describe_seg(sg) -> str:
    t = sg.GetType()
    kind = _SEG_TYPES.get(t, str(t))
    c = " construction" if sg.ConstructionGeometry else ""
    if t == 0:
        ln = T(sg, "ISketchLine")
        return f"LINE {_pt(T(ln.GetStartPoint2(), 'ISketchPoint'))}->{_pt(T(ln.GetEndPoint2(), 'ISketchPoint'))}{c}"
    if t == 1:
        arc = T(sg, "ISketchArc")
        cen = T(arc.GetCenterPoint2(), "ISketchPoint")
        full = arc.IsCircle()
        return f"{'CIRCLE' if full else 'ARC'} c={_pt(cen)} r={fnum(arc.GetRadius() * 1000, 4)}{c}"
    return kind + c


def _sketch_dims(md, sk) -> List[str]:
    out = []
    for dd in display_dims(sketch_feature(md, sk)):
        dim = T(dd.GetDimension2(0), "IDimension")
        val = dim.SystemValue
        val_s = f"{fnum(math.degrees(val), 3)}°" if dim.GetType() == 1 else fnum(val * 1000, 4)
        out.append(f"{dim.Name}={val_s}")
    return out


def t_sketch_entities(app, args):
    md = model(app)
    sk, segs, pts = sketch_items(md)
    lines = [f"{sketch_feature(md, sk).Name}: {_STATUS.get(sketch_status(sk), '?')}"]
    lines += [f"S{i + 1} {_describe_seg(sg)}" for i, sg in enumerate(segs)]
    lines += [f"P{i + 1} {_pt(q)}" for i, q in enumerate(pts)]
    lines.append("O = part origin")
    dims = _sketch_dims(md, sk)
    if dims:
        lines.append("dims: " + ", ".join(dims))
    return _ok("\n".join(lines))


def _select_sketch_items(md, tokens, mark: int = 0):
    sk, segs, pts = sketch_items(md)
    sm = T(md.SelectionManager, "ISelectionMgr")
    md.ClearSelection2(True)
    kinds = []
    for n, tok in enumerate(parse_names(tokens)):
        t = tok.upper()
        append = n > 0
        if t in ("O", "ORIGIN"):
            if not select_by_id(md.Extension, "Point1@Origin", "EXTSKETCHPOINT", append=append, mark=mark):
                raise ExtError("Could not select the origin")
            kinds.append("P")
            continue
        if t[:1] not in ("S", "P") or not t[1:].isdigit():
            raise ExtError(f"Bad entity '{tok}': use S<n>, P<n> or O (see sketch_entities)")
        pool = segs if t[0] == "S" else pts
        i = int(t[1:]) - 1
        if not 0 <= i < len(pool):
            raise ExtError(f"No {tok} (have S1..S{len(segs)}, P1..P{len(pts)})")
        data = T(sm.CreateSelectData(), "ISelectData")
        data.Mark = mark
        if not pool[i].Select4(append, data):
            raise ExtError(f"Could not select {tok}")
        kinds.append(t[0] if t[0] == "P" else ("L" if pool[i].GetType() == 0 else "A"))
    return sk, kinds


_REL = {
    "coincident": "sgCOINCIDENT", "concentric": "sgCONCENTRIC", "coradial": "sgCORADIAL",
    "tangent": "sgTANGENT", "equal": "sgSAMELENGTH", "fix": "sgFIXED", "merge": "sgMERGEPOINTS",
    "collinear": "sgCOLINEAR", "parallel": "sgPARALLEL", "perpendicular": "sgPERPENDICULAR",
    "midpoint": "sgATMIDDLE", "symmetric": "sgSYMMETRIC",
}


def t_add_sketch_relation(app, args):
    md = model(app)
    rel = str(args.get("relation", "")).lower()
    sk, kinds = _select_sketch_items(md, args["entities"])
    if rel in ("horizontal", "vertical"):
        pts_only = all(k == "P" for k in kinds)
        code = {"horizontal": ("sgHORIZONTAL2D", "sgHORIZONTALPOINTS2D"),
                "vertical": ("sgVERTICAL2D", "sgVERTICALPOINTS2D")}[rel][1 if pts_only else 0]
    elif rel in _REL:
        code = _REL[rel]
    else:
        raise ExtError(f"relation must be one of horizontal, vertical, {', '.join(_REL)}")
    md.SketchAddConstraints(code)
    md.ClearSelection2(True)
    return _ok(f"{rel} {args['entities']} -> sketch {_STATUS.get(sketch_status(sk), '?')}")


def t_add_sketch_dimension(app, args):
    """Dimension selected sketch entities, set its value, optionally link it
    to a global variable. Text position is in sketch coordinates (mm)."""
    md = model(app)
    kind = (args.get("kind") or "auto").lower()
    sk, kinds = _select_sketch_items(md, args["entities"])
    sk_name = sketch_feature(md, sk).Name
    # text position: given, or next to the first entity
    a = sketch_to_model(sk)
    tx, ty = args.get("x"), args.get("y")
    if tx is None or ty is None:
        tx, ty = 5.0, 5.0
    p = xform(a, [float(tx) / 1000, float(ty) / 1000, 0.0])
    app_t = T(app, "ISldWorks")
    pref = app_t.GetUserPreferenceToggle(get_constant("swInputDimValOnCreate"))
    app_t.SetUserPreferenceToggle(get_constant("swInputDimValOnCreate"), False)
    try:
        fn = {"auto": md.AddDimension2, "horizontal": md.AddHorizontalDimension2,
              "vertical": md.AddVerticalDimension2, "diameter": md.AddDiameterDimension2,
              "radius": md.AddRadialDimension2, "angle": md.AddDimension2}.get(kind)
        if fn is None:
            raise ExtError("kind must be auto / horizontal / vertical / diameter / radius / angle")
        dd = fn(p[0], p[1], p[2])
    finally:
        app_t.SetUserPreferenceToggle(get_constant("swInputDimValOnCreate"), pref)
    md.ClearSelection2(True)
    if dd is None:
        raise ExtError(f"Dimension not created ({kind} on {args['entities']}). Check entity types "
                       f"(diameter: one circle; horizontal/vertical: two points or a line)")
    dim = T(T(dd, "IDisplayDimension").GetDimension2(0), "IDimension")
    name = f"{dim.Name}@{sk_name}"
    val = args.get("value")
    if val is not None and str(val) != "":
        num = float(val)
        dim.SetSystemValue3(math.radians(num) if dim.GetType() == 1 else num / 1000,
                            get_constant("swAllConfiguration"), None)
    link = args.get("link")
    if link:
        em = T(md.GetEquationMgr(), "IEquationMgr")
        if _eq_index(em, link) < 0:
            raise ExtError(f"Global variable '{link}' not found -- add_parameter first")
        if _eq_add(em, f'"{name}" = "{link}"') < 0:
            raise ExtError(f"Could not link {name} to {link}")
        em.EvaluateAll()
    # equation evaluation rebuilds and drops SW out of sketch edit mode --
    # put the user back in the sketch they were dimensioning
    if md.SketchManager.ActiveSketch is None:
        md.ClearSelection2(True)
        find_feature(md, sk_name).Select2(False, 0)
        md.EditSketch()
        md.ClearSelection2(True)
        sk = _active_sketch(md)
    cur = dim.SystemValue
    shown = f"{fnum(math.degrees(cur), 3)}°" if dim.GetType() == 1 else fnum(cur * 1000, 4)
    return _ok(f"{name} = {shown}" + (f' ("{link}")' if link else "") +
               f" -> sketch {_STATUS.get(sketch_status(sk), '?')}")


def t_transaction(app, args):
    md = model(app)
    action = (args.get("action") or "begin").lower()
    key = md.GetTitle()
    if action == "begin":
        if key in _TX:
            raise ExtError(f"Transaction '{_TX[key]['name']}' already open on {key}")
        md.Extension.StartRecordingUndoObject()
        _TX[key] = {"name": args.get("name") or "MCP batch",
                    "feats": [f.Name for f in user_features(md)], "t": time.time()}
        return _ok(f"begin '{_TX[key]['name']}' ({len(_TX[key]['feats'])} features)")
    tx = _TX.pop(key, None)
    if tx is None:
        raise ExtError("No open transaction on this document")
    md.Extension.FinishRecordingUndoObject2(tx["name"], False)
    if action == "commit":
        new = [f.Name for f in user_features(md) if f.Name not in tx["feats"]]
        return _ok(f"commit '{tx['name']}': +{len(new)} features {new}")
    if action != "abort":
        raise ExtError("action must be begin / commit / abort")
    md.EditUndo2(1)
    extra = [f.Name for f in user_features(md) if f.Name not in tx["feats"]]
    if extra:
        # undo did not cover everything (e.g. an op outside the undo record) --
        # fall back to deleting what is left, newest first
        for n in reversed(extra):
            md.ClearSelection2(True)
            find_feature(md, n).Select2(False, 0)
            md.Extension.DeleteSelection2(get_constant("swDelete_Absorbed"))
        md.ClearSelection2(True)
    left = [f.Name for f in user_features(md) if f.Name not in tx["feats"]]
    return _ok(f"abort '{tx['name']}': rolled back" + (f", STILL PRESENT: {left}" if left else ""))


HANDLERS = {
    "inspect": t_inspect,
    "list_faces": t_list_faces,
    "list_edges": t_list_edges,
    "find_face": t_find_face,
    "find_edge": t_find_edge,
    "select_entities": t_select,
    "draw_profile": t_draw_profile,
    "hole": t_hole,
    "circular_pattern": t_circular_pattern,
    "linear_pattern": t_linear_pattern,
    "mirror_features": t_mirror,
    "get_parameters": t_get_parameters,
    "set_parameter": t_set_parameter,
    "add_parameter": t_add_parameter,
    "delete_feature": t_delete_feature,
    "suppress_feature": t_suppress_feature,
    "transaction": t_transaction,
    "sketch_entities": t_sketch_entities,
    "add_sketch_relation": t_add_sketch_relation,
    "add_sketch_dimension": t_add_sketch_dimension,
}


def dispatch(name: str, args: Dict, automation) -> Dict:
    if not automation.is_connected:
        r = automation.connect()
        if not r["success"]:
            return r
    try:
        return HANDLERS[name](automation.app, args or {})
    except ExtError as e:
        return _err(str(e))
    except KeyError as e:
        return _err(f"Missing argument: {e}", 107, "swInvalidInput")
    except Exception as e:
        logger.error(f"{name} failed: {e}\n{traceback.format_exc()}")
        return _err(f"{type(e).__name__}: {e}")


# Tools whose success message gets the volume/topology delta appended,
# and what the volume is expected to do.
REPORTED = {
    "extrude_sketch": "add",
    "cut_extrude": "remove",
    "revolve_sketch": None,  # boss or cut, decided by args
    "fillet_edges": None,
    "chamfer_edges": "remove",
}


def expect_for(name: str, args: Dict) -> Optional[str]:
    if name == "revolve_sketch":
        return "remove" if args.get("cut") else "add"
    return REPORTED.get(name)


# ============================================================================
# Tool schemas (server.py appends these to list_tools)
# ============================================================================

def _obj(props: Dict, required: Optional[List[str]] = None) -> Dict:
    return {"type": "object", "properties": props, "required": required or []}


_XYZ_MATCH = {
    "x": {"type": "number", "default": 0}, "y": {"type": "number", "default": 0},
    "z": {"type": "number", "default": 0},
    "use_x": {"type": "boolean", "default": False}, "use_y": {"type": "boolean", "default": False},
    "use_z": {"type": "boolean", "default": False},
    "tolerance": {"type": "number", "default": 0.1, "description": "mm"},
    "type": {"type": "string", "description": "Optional type filter (PLANE/CYL/CONE/TORUS... or LINE/CIRC...)"},
}

TOOL_SCHEMAS = [
    ("inspect",
     "One-call snapshot of the active part: volume, area, bbox, face-type histogram, "
     "feature tree, active sketch + its frame. Use instead of several diagnostic calls.",
     _obj({})),
    ("list_faces",
     "Compact face table: '[i] PLANE A=... p=(x,y,z) n=(...)', CYL/CONE/TORUS with radius/axis. "
     "p is a point ON the face (mm) -- usable for create_sketch_on_face / hole. "
     "Indices change after every feature.",
     _obj({"min_area": {"type": "number", "default": 0, "description": "mm²"},
           "max_count": {"type": "integer", "default": 60},
           "type": {"type": "string", "description": "Filter: PLANE, CYL, CONE, TORUS, ..."}})),
    ("list_edges",
     "Compact edge table: '[i] LINE L=... mid=(...) dir=(...)', CIRC with r/center/normal. "
     "Use indices with fillet_edges/chamfer_edges(edge_indices) or select_entities.",
     _obj({"min_length": {"type": "number", "default": 0, "description": "mm"},
           "max_count": {"type": "integer", "default": 60},
           "type": {"type": "string", "description": "Filter: LINE, CIRC, ..."}})),
    ("find_face", "Find faces whose on-face point p matches given coordinates (mm) on the use_* axes.",
     _obj(_XYZ_MATCH)),
    ("find_edge", "Find edges whose midpoint matches given coordinates (mm) on the use_* axes.",
     _obj(_XYZ_MATCH)),
    ("select_entities",
     "Select faces/edges by index (from list_faces/list_edges) with an optional selection mark, "
     "for raw FeatureManager calls via execute_python.",
     _obj({"kind": {"type": "string", "enum": ["edge", "face"]},
           "indices": {"type": "string", "description": "e.g. '1,4,7'"},
           "append": {"type": "boolean", "default": False},
           "mark": {"type": "integer", "default": 0}}, ["kind", "indices"])),
    ("draw_profile",
     "Draw a polyline in the active sketch from [[x,y],...] (mm, sketch coordinates), closed by default. "
     "Auto-relations are off, so coordinates stay exact. construction=true draws centerlines.",
     _obj({"points": {"type": "string", "description": "JSON array of [x,y] pairs in mm"},
           "closed": {"type": "boolean", "default": True},
           "construction": {"type": "boolean", "default": False}}, ["points"])),
    ("hole",
     "Native Hole Wizard hole on a planar face. Place by a point ON the face (x,y,z in mm, model space -- "
     "take p from list_faces) or several points in `points`. kind: drill (diameter), tap (size 'M6'), "
     "cbore (size 'M6', ISO 4762 seat). depth 0 = through all.",
     _obj({"kind": {"type": "string", "enum": ["drill", "tap", "cbore"], "default": "drill"},
           "x": {"type": "number"}, "y": {"type": "number"}, "z": {"type": "number"},
           "points": {"type": "string", "description": "Optional JSON [[x,y,z],...] mm, all on the same face"},
           "diameter": {"type": "number", "description": "mm (drill)"},
           "size": {"type": "string", "description": "ISO size, e.g. M6 (tap/cbore)"},
           "depth": {"type": "number", "default": 0, "description": "mm, 0 = through all (tap: thread depth)"},
           "thread": {"type": "string", "enum": ["cosmetic", "modeled"], "default": "cosmetic",
                      "description": "tap only: cosmetic thread on the tap-drill hole (standard) or modeled major Ø"}})),
    ("circular_pattern",
     "Circular pattern of features around an axis: 'X'/'Y'/'Z' (world axis through origin), "
     "'edge:N', 'face:N' (cylindrical face), or an axis feature name.",
     _obj({"features": {"type": "string", "description": "Comma-separated feature names"},
           "count": {"type": "integer", "default": 4, "description": "Total instances incl. seed"},
           "angle": {"type": "number", "default": 360, "description": "Total angle, deg (equal spacing)"},
           "axis": {"type": "string", "default": "Z"},
           "reverse": {"type": "boolean", "default": False},
           "geometry_pattern": {"type": "boolean", "default": False}}, ["features"])),
    ("linear_pattern",
     "Linear pattern of features. direction: 'X'/'Y'/'Z', 'edge:N', or a feature name. Optional 2nd direction.",
     _obj({"features": {"type": "string"},
           "direction": {"type": "string", "default": "X"},
           "count": {"type": "integer", "default": 2},
           "spacing": {"type": "number", "default": 10, "description": "mm"},
           "reverse": {"type": "boolean", "default": False},
           "direction2": {"type": "string"},
           "count2": {"type": "integer", "default": 1},
           "spacing2": {"type": "number", "default": 10},
           "reverse2": {"type": "boolean", "default": False},
           "geometry_pattern": {"type": "boolean", "default": False}}, ["features"])),
    ("mirror_features",
     "Mirror features about a plane: 'Front'/'Top'/'Right', a plane feature name, or 'face:N'.",
     _obj({"features": {"type": "string"}, "plane": {"type": "string", "default": "Right"}},
          ["features"])),
    ("get_parameters",
     "Equations/global variables and all feature dimensions (D1@Sketch1=...; mm / deg).",
     _obj({"feature": {"type": "string", "description": "Only dimensions of this feature"}})),
    ("set_parameter",
     "Set a dimension ('D1@Boss-Extrude1', number in mm/deg, or an expression -> equation) "
     "or a global variable (name without @). Rebuilds and reports the delta.",
     _obj({"name": {"type": "string"}, "expression": {"type": "string"}}, ["name", "expression"])),
    ("add_parameter",
     "Add a global variable: name + expression, e.g. ('wall', '3mm') or ('d2', '\"d1\" * 2').",
     _obj({"name": {"type": "string"}, "expression": {"type": "string"}}, ["name", "expression"])),
    ("delete_feature", "Delete features by name (with absorbed sketches).",
     _obj({"names": {"type": "string", "description": "Comma-separated"}}, ["names"])),
    ("suppress_feature", "Suppress / unsuppress features by name.",
     _obj({"names": {"type": "string"}, "suppressed": {"type": "boolean", "default": True}}, ["names"])),
    ("transaction",
     "Group several operations: 'begin' before a risky multi-step build, then 'commit', or 'abort' "
     "to roll back everything since begin (undo + cleanup of leftovers).",
     _obj({"action": {"type": "string", "enum": ["begin", "commit", "abort"], "default": "begin"},
           "name": {"type": "string", "default": "MCP batch"}})),
    ("sketch_entities",
     "Active sketch: definition status (UNDER/fully/OVER-defined), segments S1.., points P1.. "
     "(sketch coords, mm), O = origin, existing dimensions. Use before add_sketch_relation/"
     "add_sketch_dimension. Every sketch that drives a feature must end up fully defined.",
     _obj({})),
    ("add_sketch_relation",
     "Add a geometric relation between active-sketch entities (S<n>, P<n>, O from sketch_entities): "
     "horizontal/vertical (line, or two points), coincident, concentric, coradial, tangent, equal, "
     "collinear, parallel, perpendicular, midpoint, symmetric, merge, fix. Returns the sketch status.",
     _obj({"entities": {"type": "string", "description": "e.g. 'P1,O' or 'S1,S2'"},
           "relation": {"type": "string"}}, ["entities", "relation"])),
    ("add_sketch_dimension",
     "Add a driving dimension to active-sketch entities and set its value (mm / deg), optionally "
     "linked to a global variable (add_parameter first). kind: diameter/radius (one circle/arc), "
     "horizontal/vertical (two points, or point+O), auto (line length, point-line distance), angle "
     "(two lines). x,y = text position in sketch coords. Returns the sketch status.",
     _obj({"entities": {"type": "string", "description": "e.g. 'S1' or 'P2,O'"},
           "kind": {"type": "string", "enum": ["auto", "horizontal", "vertical", "diameter", "radius", "angle"],
                    "default": "auto"},
           "value": {"type": "number", "description": "mm or deg; omit to keep the drawn value"},
           "link": {"type": "string", "description": "Global variable name to drive this dimension"},
           "x": {"type": "number"}, "y": {"type": "number"}}, ["entities"])),
    ("reload_api",
     "Hot-reload the server's Python code (automation/*, ext.py) from disk after editing it. "
     "Keeps the SolidWorks connection. New/changed tool SCHEMAS still need an MCP restart.",
     _obj({})),
]
