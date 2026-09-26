"""
SolidWorks MCP -- extended tools (hot-reloadable)
-------------------------------------------------
Everything added on top of the upstream server lives here so that
`reload_api` can pick up edits without restarting the MCP process
(tool *schemas* still need a client restart; handler code does not).

Design rules (see CLAUDE.md / NOTES.md):
- Late binding only. Every COM object is a `win32com.client.dynamic` wrapper;
  `T(obj, "IInterface")` = QueryInterface to that interface + dynamic wrap.
  No makepy / gencache / EnsureDispatch anywhere (see utils/com_helpers.py
  for why and for the calling rules).
- Zero-arg members (property or method): `v(obj, "Name")`. Out-params:
  `com.out_int()` & co, or `com.call_out(obj, "M", ..., OUT)`. Null object
  arguments: `com.nothing()`, never a bare None.
- All user-facing numbers are millimetres / degrees. SW internals are
  metres / radians.
- Results are compact single-line text: this output is read by a model,
  every token counts.
"""

import hashlib
import json
import math
import os
import tempfile
import time
import logging
import traceback
from typing import Dict, List, Optional

import pythoncom
import win32com.client
from win32com.client import dynamic

from .utils import com_helpers as com
from .utils.com_helpers import v, nothing, OUT, call_out
from .utils.typelib import get_iid, get_constant, constants as typelib_constants
from .utils.sw_finder import find_template

logger = logging.getLogger(__name__)


# ============================================================================
# Late-bound COM access
# ============================================================================

def T(obj, iface: str):
    """Late-bound view of a COM object through one of its interfaces.

    QueryInterface first: SW objects implement several interfaces (a sketch
    object is also its IFeature), and names resolved on one interface's
    IDispatch can hit another interface's dispids. The IID comes from the
    registered typelib (utils/typelib.py, no codegen). Falls back to the
    object's own IDispatch if it refuses the QI."""
    if obj is None:
        return None
    ole = getattr(obj, "_oleobj_", obj)
    try:
        ole = ole.QueryInterface(get_iid(iface), pythoncom.IID_IDispatch)
    except pythoncom.com_error:
        pass
    return dynamic.Dispatch(ole, iface)


def empty_dispatch():
    return nothing()


def select_by_id(ext_obj, name, typ, x=0.0, y=0.0, z=0.0, append=False, mark=0) -> bool:
    """IModelDocExtension.SelectByID2 with an empty Callout (VT_DISPATCH NULL)."""
    try:
        return bool(ext_obj.SelectByID2(name, typ, float(x), float(y), float(z),
                                        append, mark, nothing(), 0))
    except pythoncom.com_error as e:
        raise ExtError(f"SelectByID2({name!r}, {typ}) failed: {e}")


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
    """Active document as IModelDoc2 (raises ExtError if none)."""
    if app is None:
        raise ExtError("Not connected to SolidWorks")
    doc = app.ActiveDoc
    if doc is None:
        raise ExtError("No active document")
    return T(doc, "IModelDoc2")


def save_in_place(md):
    """IModelDoc2.Save3 -> (ok, errors, warnings); Errors/Warnings are
    by-ref out params."""
    ok, err, warn = call_out(md, "Save3", 1, OUT, OUT)  # swSaveAsOptions_Silent
    return bool(ok), int(err or 0), int(warn or 0)


def is_part(md) -> bool:
    return v(md, "GetType") == 1


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
    f = v(md, "FirstFeature")
    while f is not None:
        f = T(f, "IFeature")
        yield f
        f = v(f, "GetNextFeature")


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
            "F": sum(v(b, "GetFaceCount") for b in bs),
            "E": sum(v(b, "GetEdgeCount") for b in bs),
            "feats": [f.Name for f in user_features(md)],
        }
    except Exception as e:
        logger.debug(f"snapshot failed: {e}")
        return None


_FEATURE_ERROR_NAMES: Optional[Dict[int, str]] = None


def feature_error_name(code) -> str:
    """'FilletRadiusTooBig2' for 19, built from swFeatureError* in the
    constants typelib. A bare 'code 19' tells the model nothing; the name is
    usually the whole diagnosis. Falls back to the number if the typelib is
    unavailable."""
    global _FEATURE_ERROR_NAMES
    if _FEATURE_ERROR_NAMES is None:
        table: Dict[int, str] = {}
        try:
            prefix = "swFeatureError"
            for n, val in typelib_constants().items():
                if n.startswith(prefix) and isinstance(val, int):
                    table.setdefault(val, n[len(prefix):])
        except Exception:
            pass
        _FEATURE_ERROR_NAMES = table
    name = _FEATURE_ERROR_NAMES.get(int(code))
    return f"{name} ({code})" if name else f"code {code}"


def error_code(feat):
    """IFeature.GetErrorCode2 -> (code, is_warning); IsWarning is a by-ref out."""
    code, warn = call_out(feat, "GetErrorCode2", com.Out(pythoncom.VT_BOOL))
    return int(code or 0), bool(warn)


def feature_errors(md, names: List[str]) -> List[str]:
    out = []
    for n in names:
        try:
            code, warn = error_code(find_feature(md, n))
            if code:
                out.append(f"{'warning' if warn else 'ERROR'} in {n}: "
                           f"{feature_error_name(code)}")
        except Exception:
            pass
    return out


_STATUS = {1: "unknown", 2: "UNDER-defined", 3: "fully defined", 4: "OVER-defined",
           5: "no solution", 6: "invalid solution"}


def sketch_status(sk) -> int:
    return v(T(sk, "ISketch"), "GetConstrainedStatus")


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
        sub = v(f, "GetFirstSubFeature")
        while sub is not None:
            subs.append(T(sub, "IFeature"))
            sub = v(T(sub, "IFeature"), "GetNextSubFeature")
        for sf in subs + [T(x, "IFeature") for x in (v(f, "GetParents") or ())]:
            if v(sf, "GetTypeName2") != "ProfileFeature" or sf.Name in seen:
                continue
            seen.add(sf.Name)
            st = sketch_status(v(sf, "GetSpecificFeature2"))
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
        for fo in v(b, "GetFaces") or ():
            i += 1
            f = T(fo, "IFace2")
            s = T(v(f, "GetSurface"), "ISurface")
            st = v(s, "Identity")
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
                         "area": v(f, "GetArea") * 1e6, "p": p, "extra": extra, "obj": f})
    return rows


def edge_rows(md) -> List[Dict]:
    rows, i = [], 0
    for bi, b in enumerate(bodies(md)):
        for eo in v(b, "GetEdges") or ():
            i += 1
            e = T(eo, "IEdge")
            c = T(v(e, "GetCurve"), "ICurve")
            cp = T(v(e, "GetCurveParams3"), "ICurveParamData")
            u0, u1 = cp.UMinValue, cp.UMaxValue
            length = c.GetLength3(u0, u1) * 1000
            ct = v(c, "Identity")
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
        ev = T(v(best["obj"], "GetSurface"), "ISurface").EvaluateAtPoint(*p)
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
        data = T(v(sm, "CreateSelectData"), "ISelectData")
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
    (IMathUtility.CreatePoint with a Python list gave garbage -- 1e-311
    values -- so the math is done here.)"""
    s = a[12] if len(a) > 12 and a[12] else 1.0
    return [sum(p[i] * a[3 * i + j] for i in range(3)) * s + a[9 + j] for j in range(3)]


def sketch_to_model(sk) -> List[float]:
    """ArrayData of the active sketch's sketch->model transform."""
    return list(T(v(T(sk.ModelToSketchTransform, "IMathTransform"), "Inverse"), "IMathTransform").ArrayData)


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
    rp = T(v(plane_feat, "GetSpecificFeature2"), "IRefPlane")
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
    v(md, "BlankRefGeom")
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
    title = v(md, "GetTitle")
    lines = [f"{title} | {v(md, "GetPathName") or '<unsaved>'}"]
    if is_part(md):
        bs = bodies(md)
        vol = volume_mm3(md)
        area = 0.0
        if bs:
            area = T(v(md.Extension, "CreateMassProperty"), "IMassProperty").SurfaceArea * 1e6
        lines.append(f"bodies {len(bs)} | V {fnum(vol, 1)} mm³ | A {fnum(area, 1)} mm²")
        for bi, b in enumerate(bs):
            bx = v(b, "GetBodyBox")
            size = [bx[k + 3] - bx[k] for k in range(3)]
            fr = [T(f, "IFace2") for f in v(b, "GetFaces") or ()]
            kinds = {}
            for f in fr:
                t = _SURF.get(v(T(v(f, "GetSurface"), "ISurface"), "Identity"), "?")
                kinds[t] = kinds.get(t, 0) + 1
            lines.append(
                f"body{bi + 1}: bbox X[{fnum(bx[0]*1000)}..{fnum(bx[3]*1000)}] "
                f"Y[{fnum(bx[1]*1000)}..{fnum(bx[4]*1000)}] Z[{fnum(bx[2]*1000)}..{fnum(bx[5]*1000)}] "
                f"size {'×'.join(fnum(s*1000) for s in size)} | faces {v(b, "GetFaceCount")} "
                f"({', '.join(f'{k}×{n}' for k, n in sorted(kinds.items()))}) | edges {v(b, "GetEdgeCount")}")
        try:
            mat, db = call_out(T(md, "IPartDoc"), "GetMaterialPropertyName2", "",
                               com.Out(pythoncom.VT_BSTR))
            if mat:
                lines.append(f"material: {mat}" + (f" ({db})" if db else ""))
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
        em = T(v(md, "GetEquationMgr"), "IEquationMgr")
        if v(em, "GetCount"):
            lines.append(f"equations: {v(em, "GetCount")} (get_parameters for details)")
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
            data = T(v(feat, "GetDefinition"), "IWizardHoleFeatureData2")
            data.AccessSelections(md, nothing())
            for k, val in ed.items():
                setattr(data, k, val)
            if not feat.ModifyDefinition(data, md, nothing()):
                v(data, "ReleaseSelectionAccess")
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
    sub = v(feat, "GetFirstSubFeature")
    sketch_feat = None
    while sub is not None:
        sub = T(sub, "IFeature")
        if "Profile" in v(sub, "GetTypeName2"):
            sketch_feat = sub
            break
        sub = v(sub, "GetNextSubFeature")
    if sketch_feat is None:
        raise ExtError("Hole position sketch not found")
    md.ClearSelection2(True)
    sketch_feat.Select2(False, 0)
    v(md, "EditSketch")
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
    v(md, "EditRebuild3")
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
    em = T(v(md, "GetEquationMgr"), "IEquationMgr")
    for i in range(v(em, "GetCount")):
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
            if v(dim, "GetType") == 1:
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
    for i in range(v(em, "GetCount")):
        if v(em, "Equation", i).strip().startswith(key):
            return i
    return -1


def t_set_parameter(app, args):
    md = model(app)
    name, expr = args["name"], str(args["expression"]).strip()
    em = T(v(md, "GetEquationMgr"), "IEquationMgr")
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
            val = math.radians(num) if v(dim, "GetType") == 1 else num / 1000
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
    v(em, "EvaluateAll")
    v(md, "EditRebuild3")
    return _ok(f"{name} = {expr} | " + delta(app, before, snapshot(app)))


def t_add_parameter(app, args):
    md = model(app)
    name, expr = args["name"], str(args["expression"]).strip()
    em = T(v(md, "GetEquationMgr"), "IEquationMgr")
    if _eq_index(em, name) >= 0:
        raise ExtError(f"'{name}' already exists -- use set_parameter")
    eq = f'"{name}" = {expr}'
    if _eq_add(em, eq) < 0:
        raise ExtError(f"Equation rejected: {eq}")
    v(em, "EvaluateAll")
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
    v(md, "EditRebuild3")
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
            s2 = T(v(f, "GetSpecificFeature2"), "ISketch")
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
    segs = [T(x, "ISketchSegment") for x in (v(sk, "GetSketchSegments") or ())]
    pts = [T(x, "ISketchPoint") for x in (v(sk, "GetSketchPoints2") or ())]
    return sk, segs, pts


def _pt(q) -> str:
    return f"({fnum(q.X * 1000)},{fnum(q.Y * 1000)})"


def display_dims(f):
    """Iterate IFeature display dimensions. With no dimensions SW returns an
    int 0 instead of Nothing -- anything that is not a COM object ends the walk."""
    dd = v(f, "GetFirstDisplayDimension")
    while hasattr(dd, "_oleobj_"):
        dd = T(dd, "IDisplayDimension")
        yield dd
        dd = f.GetNextDisplayDimension(dd)


def _describe_seg(sg) -> str:
    t = v(sg, "GetType")
    kind = _SEG_TYPES.get(t, str(t))
    c = " construction" if sg.ConstructionGeometry else ""
    if t == 0:
        ln = T(sg, "ISketchLine")
        return f"LINE {_pt(T(v(ln, "GetStartPoint2"), 'ISketchPoint'))}->{_pt(T(v(ln, "GetEndPoint2"), 'ISketchPoint'))}{c}"
    if t == 1:
        arc = T(sg, "ISketchArc")
        cen = T(v(arc, "GetCenterPoint2"), "ISketchPoint")
        full = v(arc, "IsCircle")
        return f"{'CIRCLE' if full else 'ARC'} c={_pt(cen)} r={fnum(v(arc, "GetRadius") * 1000, 4)}{c}"
    return kind + c


def _sketch_dims(md, sk) -> List[str]:
    out = []
    for dd in display_dims(sketch_feature(md, sk)):
        dim = T(dd.GetDimension2(0), "IDimension")
        val = dim.SystemValue
        val_s = f"{fnum(math.degrees(val), 3)}°" if v(dim, "GetType") == 1 else fnum(val * 1000, 4)
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
        data = T(v(sm, "CreateSelectData"), "ISelectData")
        data.Mark = mark
        if not pool[i].Select4(append, data):
            raise ExtError(f"Could not select {tok}")
        kinds.append(t[0] if t[0] == "P" else ("L" if v(pool[i], "GetType") == 0 else "A"))
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
        dim.SetSystemValue3(math.radians(num) if v(dim, "GetType") == 1 else num / 1000,
                            get_constant("swAllConfiguration"), None)
    link = args.get("link")
    if link:
        em = T(v(md, "GetEquationMgr"), "IEquationMgr")
        if _eq_index(em, link) < 0:
            raise ExtError(f"Global variable '{link}' not found -- add_parameter first")
        if _eq_add(em, f'"{name}" = "{link}"') < 0:
            raise ExtError(f"Could not link {name} to {link}")
        v(em, "EvaluateAll")
    # equation evaluation rebuilds and drops SW out of sketch edit mode --
    # put the user back in the sketch they were dimensioning
    if md.SketchManager.ActiveSketch is None:
        md.ClearSelection2(True)
        find_feature(md, sk_name).Select2(False, 0)
        v(md, "EditSketch")
        md.ClearSelection2(True)
        sk = _active_sketch(md)
    cur = dim.SystemValue
    shown = f"{fnum(math.degrees(cur), 3)}°" if v(dim, "GetType") == 1 else fnum(cur * 1000, 4)
    return _ok(f"{name} = {shown}" + (f' ("{link}")' if link else "") +
               f" -> sketch {_STATUS.get(sketch_status(sk), '?')}")


def t_transaction(app, args):
    md = model(app)
    action = (args.get("action") or "begin").lower()
    key = v(md, "GetTitle")
    if action == "begin":
        if key in _TX:
            raise ExtError(f"Transaction '{_TX[key]['name']}' already open on {key}")
        v(md.Extension, "StartRecordingUndoObject")
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


# ============================================================================
# Stage 2: busy probe -- refuse instead of hanging
# ============================================================================
#
# The problem this solves: every COM call is marshalled into SLDWORKS.exe from
# outside, synchronously, on the asyncio event-loop thread. If SolidWorks has a
# modal dialog open or is grinding through a rebuild, the call does not fail --
# it BLOCKS, and takes the whole MCP server with it: no other tool can run, no
# cancel is processed, and the client eventually times out with nothing useful
# to say. This is the single worst failure mode of an out-of-process server,
# and the in-process add-in architecture (SW+) does not have it at all.
#
# The fix is the trick SW+ uses from inside, which works just as well from
# outside because it needs only a window handle: ping the main window with
# SendMessageTimeout(WM_NULL, SMTO_ABORTIFHUNG). A responsive window answers in
# well under a millisecond (measured: 0.3 ms). A window that is modal or busy
# does not answer, and the API returns ERROR_TIMEOUT instead of blocking.

import ctypes
import ctypes.wintypes as _w

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_SendMessageTimeoutW = _user32.SendMessageTimeoutW
_SendMessageTimeoutW.argtypes = [_w.HWND, _w.UINT, _w.WPARAM, _w.LPARAM,
                                 _w.UINT, _w.UINT, ctypes.POINTER(ctypes.c_size_t)]
_SendMessageTimeoutW.restype = _w.LPARAM
_IsWindow = _user32.IsWindow
_IsWindow.argtypes = [_w.HWND]
_IsWindow.restype = _w.BOOL

_WM_NULL = 0x0000
_SMTO_BLOCK = 0x0001
_SMTO_ABORTIFHUNG = 0x0002
_ERROR_TIMEOUT = 1460

# Default ping budget. Generous next to the 0.3 ms an idle window takes, and
# still 400x cheaper than waiting out a client-side timeout on a hung call.
BUSY_TIMEOUT_MS = 1200

_HWND_CACHE: Dict[int, int] = {}


def sw_hwnd(app) -> Optional[int]:
    """Handle of the SolidWorks main window, cached.

    Cached deliberately: fetching it goes through Frame(), which is itself a
    COM call and would hang for exactly the reason we are trying to detect.
    The cache is validated with IsWindow, so a restarted SolidWorks is picked
    up on the next call rather than probing a dead handle forever.
    """
    key = id(app)
    hwnd = _HWND_CACHE.get(key)
    if hwnd and _IsWindow(_w.HWND(hwnd)):
        return hwnd
    try:
        hwnd = int(v(T(v(app, "Frame"), "IFrame"), "GetHWndx64"))
    except Exception:
        try:
            hwnd = int(v(T(v(app, "Frame"), "IFrame"), "GetHWnd"))
        except Exception:
            return None
    if not hwnd or not _IsWindow(_w.HWND(hwnd)):
        return None
    _HWND_CACHE[key] = hwnd
    return hwnd


def sw_busy(app, timeout_ms: Optional[int] = None) -> Optional[str]:
    """None if SolidWorks is ready to take a call, else why it is not.

    Never raises and never blocks longer than timeout_ms: if this function
    cannot decide, it returns None (assume ready) rather than blocking work on
    its own uncertainty. A probe that itself becomes a failure mode would be
    worse than no probe.
    """
    if timeout_ms is None:
        timeout_ms = BUSY_TIMEOUT_MS
    if timeout_ms <= 0:
        return None

    hwnd = sw_hwnd(app)
    if hwnd is None:
        # No window to ping: SolidWorks may be starting or headless. Not our
        # call to make -- let the real COM call produce the real error.
        return None

    out = ctypes.c_size_t(0)
    ctypes.set_last_error(0)
    rc = _SendMessageTimeoutW(_w.HWND(hwnd), _WM_NULL, 0, 0,
                              _SMTO_BLOCK | _SMTO_ABORTIFHUNG,
                              int(timeout_ms), ctypes.byref(out))
    if rc:
        return None
    err = ctypes.get_last_error()
    if err == _ERROR_TIMEOUT:
        return (f"SolidWorks is not responding ({timeout_ms} ms): a modal "
                f"dialog is open, or a long rebuild/macro is running. Close "
                f"the dialog or wait for it to finish, then retry. "
                f"(Refused instead of blocking the whole server.)")
    # Window vanished between IsWindow and the ping, or some other Win32
    # problem: say what happened, but do not pretend to know it is busy.
    _HWND_CACHE.pop(id(app), None)
    return None


# ============================================================================
# Stage 1: reference geometry, mass, rebuild errors, feature editing
# ============================================================================

# swRefPlaneReferenceConstraint_* -- verified against swconst.tlb through
# scripts/swapi.py find swRefPlaneReferenceConstraint, not from memory.
_PLANE_CONSTRAINT = {
    "distance": 8,
    "angle": 16,
    "coincident": 4,
    "parallel": 1,
    "perpendicular": 2,
    "tangent": 32,
    "midplane": 128,
}
_PLANE_FLIP = 256


def select_geom(md, spec: str, append: bool = False, mark: int = 0):
    """Select one reference for the reference-geometry tools.

    Accepts 'front'/'top'/'right' (locale-independent, resolved through the
    first three RefPlane features), 'face:3' / 'edge:7' (indices from
    list_faces / list_edges), or any feature name as shown by inspect.
    """
    spec = str(spec).strip()
    if ":" in spec:
        kind, idx = spec.split(":", 1)
        select_entities(md, kind.strip().lower(), idx, append=append, mark=mark)
        return spec
    name = resolve_plane(md, spec)
    if not find_feature(md, name).Select2(append, mark):
        raise ExtError(f"Could not select '{spec}'")
    return name


def _tree_names(md):
    return [f.Name for f in iter_features(md)]


def _new_feature(md, before_names):
    """The feature that appeared since before_names was taken, if any."""
    new = [n for n in _tree_names(md) if n not in before_names]
    if not new:
        return None
    return find_feature(md, new[-1])


def t_create_reference_plane(app, args):
    """Reference plane from one or two references.

    This is the documented way around the sketch-on-cylindrical-face dead end
    (NOTES.md): an offset plane works where InsertSketch2 on a curved face
    does not.
    """
    md = model(app)
    kind = (args.get("kind") or "distance").strip().lower()
    if kind not in _PLANE_CONSTRAINT:
        raise ExtError(f"kind must be one of {', '.join(sorted(_PLANE_CONSTRAINT))}")

    value = args.get("value")
    if kind in ("distance", "angle") and value is None:
        raise ExtError(f"kind '{kind}' needs a value (mm for distance, deg for angle)")

    constraint = _PLANE_CONSTRAINT[kind]
    if args.get("reverse"):
        constraint |= _PLANE_FLIP

    if kind == "angle":
        sysval = math.radians(float(value))
    elif value is None:
        sysval = 0.0
    else:
        sysval = float(value) / 1000.0

    before_names = _tree_names(md)
    md.ClearSelection2(True)
    refs = [select_geom(md, args["reference"], append=False, mark=0)]
    second = args.get("second_reference")
    if second:
        refs.append(select_geom(md, second, append=True, mark=0))

    fm = T(md.FeatureManager, "IFeatureManager")
    fm.InsertRefPlane(constraint, sysval, 0, 0, 0, 0)
    md.ClearSelection2(True)

    # InsertRefPlane does not reliably hand back the feature (NOTES.md,
    # typed-wrapper gotchas), so the tree is the source of truth regardless.
    made = _new_feature(md, before_names)
    if made is None:
        raise ExtError(
            f"No plane created from {refs} ({kind}={value}). "
            f"distance/angle take ONE parallel reference; coincident and "
            f"perpendicular usually need a second one (second_reference)."
        )
    if args.get("name"):
        made.Name = args["name"]
    unit = "deg" if kind == "angle" else "mm"
    txt = f"{made.Name}: {kind}"
    if value is not None:
        txt += f" {fnum(float(value), 3)} {unit}"
    txt += f" from {', '.join(refs)}"
    if args.get("reverse"):
        txt += " (flipped)"
    return _ok(txt)


def t_create_reference_axis(app, args):
    """Reference axis from two planes/points, one cylindrical face, or a line."""
    md = model(app)
    before_names = _tree_names(md)
    md.ClearSelection2(True)
    refs = []
    for n, spec in enumerate(parse_names(args["references"])):
        refs.append(select_geom(md, spec, append=n > 0, mark=0))
    if not refs:
        raise ExtError("references is empty")

    if not md.InsertAxis2(True):
        md.ClearSelection2(True)
        raise ExtError(
            f"InsertAxis2 failed for {refs}. Two intersecting planes, two "
            f"points, a cylindrical face or a line are valid; two PARALLEL "
            f"planes are not."
        )
    md.ClearSelection2(True)
    made = _new_feature(md, before_names)
    if made is None:
        raise ExtError(f"No axis appeared in the tree for {refs}")
    if args.get("name"):
        made.Name = args["name"]
    if args.get("hide"):
        md.ClearSelection2(True)
        made.Select2(False, 0)
        v(md, "BlankRefGeom")
        md.ClearSelection2(True)
    return _ok(f"{made.Name}: axis from {', '.join(refs)}")


def _mass_property(md):
    """IMassProperty2 if this install has it, else IMassProperty.

    UseSystemUnits pins the readings to SI no matter how the document is
    configured, so the conversions below start from a known base instead of
    from whatever units the user last picked.
    """
    ext_obj = md.Extension
    for factory, iface in (("CreateMassProperty2", "IMassProperty2"),
                           ("CreateMassProperty", "IMassProperty")):
        try:
            raw = v(ext_obj, factory)
        except Exception:
            continue
        if raw is not None:
            mp = T(raw, iface)
            try:
                mp.UseSystemUnits = True
            except Exception:
                pass
            return mp, iface
    raise ExtError("Could not create a mass property object for this document")


def t_mass_properties(app, args):
    md = model(app)
    bs = bodies(md)
    if not bs:
        raise ExtError("No solid bodies in the active document")
    mp, iface = _mass_property(md)
    try:
        v(mp, "Recalculate")
    except Exception:
        pass

    mass_kg = float(mp.Mass)
    vol_mm3 = float(mp.Volume) * 1e9
    area_mm2 = float(mp.SurfaceArea) * 1e6
    density = float(mp.Density)                       # kg/m3
    com_m = list(mp.CenterOfMass or (0.0, 0.0, 0.0))  # metres
    com_mm = [c * 1000.0 for c in com_m]

    grams = mass_kg * 1000.0
    mass_txt = f"{fnum(grams, 2)} g" if grams < 1000 else f"{fnum(mass_kg, 4)} kg"
    lines = [
        f"mass {mass_txt} | volume {fnum(vol_mm3, 1)} mm3 | "
        f"area {fnum(area_mm2, 1)} mm2 | density {fnum(density / 1000.0, 4)} g/cm3",
        f"center of mass (mm): {fpt(com_m)}",
    ]

    data = {
        "mass_kg": mass_kg,
        "volume_mm3": vol_mm3,
        "surface_area_mm2": area_mm2,
        "density_kg_m3": density,
        "center_of_mass_mm": com_mm,
        "bodies": len(bs),
        "interface": iface,
    }

    try:
        pm = mp.PrincipalMomentsOfInertia
        if pm:
            data["principal_moments_kg_m2"] = list(pm)
            lines.append("principal moments (kg*m2): "
                         + ", ".join(fnum(x, 6) for x in pm))
    except Exception:
        pass

    if len(bs) > 1:
        lines.append(f"({len(bs)} bodies -- values cover all of them)")
    return _ok("\n".join(lines), data)


def t_get_rebuild_errors(app, args):
    """What is broken in the active document.

    Built on IFeature.GetErrorCode2, which this codebase already relies on,
    rather than on IModelDocExtension.GetWhatsWrong: the latter returns three
    by-ref arrays and is exactly the kind of call that hands back nothing
    without saying so. GetWhatsWrong is used to enrich, never to decide.
    """
    md = model(app)
    feats = list(user_features(md))

    errors, warnings = [], []
    for f in feats:
        try:
            code, is_warn = error_code(f)
        except Exception:
            continue
        if not code:
            continue
        entry = f"{f.Name} ({v(f, 'GetTypeName2')}): {feature_error_name(code)}"
        (warnings if is_warn else errors).append(entry)

    bad_sketches = []
    for f in feats:
        if v(f, "GetTypeName2") != "ProfileFeature":
            continue
        try:
            st = sketch_status(v(f, "GetSpecificFeature2"))
        except Exception:
            continue
        if st != 3:
            bad_sketches.append(f"{f.Name} is {_STATUS.get(st, st)}")

    whats_wrong = None
    try:
        n = v(md.Extension, "GetWhatsWrongCount")
        if n:
            whats_wrong = int(n)
    except Exception:
        pass

    parts = []
    if errors:
        parts.append(f"ERRORS ({len(errors)}):\n  " + "\n  ".join(errors))
    if warnings:
        parts.append(f"warnings ({len(warnings)}):\n  " + "\n  ".join(warnings))
    if bad_sketches:
        parts.append(f"sketches not fully defined ({len(bad_sketches)}):\n  "
                     + "\n  ".join(bad_sketches))

    if not parts:
        msg = (f"clean: {len(feats)} features, no rebuild errors, "
               f"every sketch fully defined")
        if whats_wrong:
            msg += (f" -- but SW reports GetWhatsWrongCount={whats_wrong}, "
                    f"so rebuild and re-check")
        return _ok(msg)

    if whats_wrong is not None:
        parts.append(f"SW GetWhatsWrongCount = {whats_wrong}")
    return _ok("\n".join(parts), {
        "errors": errors,
        "warnings": warnings,
        "sketches_not_fully_defined": bad_sketches,
    })


def _parse_assignments(spec):
    """'D1=25, D2=10' -> [('D1', 25.0), ('D2', 10.0)]. A dict also works."""
    if isinstance(spec, dict):
        return [(str(k), float(val)) for k, val in spec.items()]
    out = []
    for chunk in str(spec).replace(";", ",").split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "=" not in chunk:
            raise ExtError(f"'{chunk}' is not name=value")
        k, val = chunk.split("=", 1)
        try:
            out.append((k.strip(), float(val.strip().replace(",", "."))))
        except ValueError:
            raise ExtError(f"'{val.strip()}' is not a number (in '{chunk}')")
    if not out:
        raise ExtError("dimensions is empty")
    return out


def t_edit_feature(app, args):
    """Change dimensions of an existing feature in place, by short name.

    set_parameter already sets a dimension given its FULL name
    ('D1@Sketch1'). This exists because the short name is what you actually
    see, and because editing several dimensions of one feature should cost
    one rebuild instead of N.
    """
    md = model(app)
    fname = args["feature"]
    feat = find_feature(md, fname)
    wanted = _parse_assignments(args["dimensions"])
    all_cfg = get_constant("swAllConfiguration")

    found = {}
    for dd in display_dims(feat):
        dim = T(dd.GetDimension2(0), "IDimension")
        found[dim.Name] = dim
        found[dim.FullName] = dim

    missing = [k for k, _ in wanted if k not in found]
    if missing:
        names = sorted({d.Name for d in found.values()})
        raise ExtError(f"{fname} has no dimension(s) {missing}. "
                       f"It has: {names or 'none'} "
                       f"(get_parameters shows their values)")

    before = snapshot(app)
    applied = []
    for key, num in wanted:
        dim = found[key]
        is_angle = v(dim, "GetType") == 1
        dim.SetSystemValue3(math.radians(num) if is_angle else num / 1000.0,
                            all_cfg, None)
        applied.append(f"{dim.Name}={fnum(num, 4)}" + ("deg" if is_angle else ""))

    v(md, "EditRebuild3")
    errs = feature_errors(md, [f.Name for f in user_features(md)])
    msg = (f"{fname}: " + ", ".join(applied) + " | "
           + delta(app, before, snapshot(app)))
    if errs:
        msg += "\nWARNING " + "; ".join(errs)
    return _ok(msg)


# ============================================================================
# Drawings
# ============================================================================

# GetUserPreferenceStringValue index for the drawing template's own default,
# same empirical constant as SolidWorksAutomation._TEMPLATE_PREF_INDEX in
# automation/base.py (part=8, assembly=9, drawing=10) -- duplicated here
# because that method hangs off `self`, not off the bare `app` this module
# gets from ext.dispatch().
_DRAWING_TEMPLATE_PREF_INDEX = 10

_STANDARD_VIEW_METHODS = {"third": "Create3rdAngleViews2", "first": "Create1stAngleViews2"}


def drawing(app):
    """Active document as IDrawingDoc (raises ExtError if it isn't one)."""
    md = model(app)
    if v(md, "GetType") != 3:
        raise ExtError(f"Active document '{v(md, "GetTitle")}' is not a drawing "
                       f"(create_drawing / open a .slddrw first)")
    return T(md, "IDrawingDoc")


def _drawing_template(app) -> Optional[str]:
    """Same resolution order as SolidWorksAutomation._get_template('drawing'):
    disk search first, then the live app's own configured default."""
    tmpl = find_template("drawing")
    if tmpl and os.path.exists(tmpl):
        return tmpl
    try:
        cand = app.GetUserPreferenceStringValue(_DRAWING_TEMPLATE_PREF_INDEX)
        if cand and os.path.exists(cand):
            return cand
    except Exception as e:
        logger.debug(f"GetUserPreferenceStringValue(drawing) failed: {e}")
    return None


def _resolve_view_name(view: str) -> str:
    """'front' / 'Front' / '*Front' -> '*Front' (the form CreateDrawViewFromModelView3
    wants). 'current' keeps the model's last-active orientation."""
    s = str(view or "current").strip()
    if s.startswith("*"):
        return s
    return "*" + s[:1].upper() + s[1:].lower()


def t_create_drawing(app, args):
    tmpl = _drawing_template(app)
    if not tmpl:
        raise ExtError(
            "No drawing template found (disk search and the running app's own "
            "default template preference both came up empty). Set a default "
            "drawing template in SolidWorks Options > Default Templates.")
    doc = app.NewDocument(tmpl, 0, 0, 0)
    if doc is None:
        raise ExtError("NewDocument returned None for the drawing template")
    md = T(doc, "IModelDoc2")
    drw = T(doc, "IDrawingDoc")
    return _ok(f"Created drawing: {v(md, "GetTitle")} | sheets {v(drw, "GetSheetCount")}")


def t_add_standard_views(app, args):
    drw = drawing(app)
    path = str(args["model_path"])
    if not os.path.exists(path):
        raise ExtError(f"Model file not found: {path}")
    projection = str(args.get("projection", "third")).strip().lower()
    method_name = _STANDARD_VIEW_METHODS.get(projection)
    if method_name is None:
        raise ExtError(f"Unknown projection '{projection}' (use 'first' or 'third')")
    before = v(drw, "GetViewCount")
    # Create1st/3rdAngleViews2 return False unless the model is open in SW
    # (confirmed 2026-09-26: closed part -> False, same part open -> 3 views).
    # CreateDrawViewFromModelView3 has no such requirement.
    opened_here = _ensure_model_open(app, path)
    try:
        ok = v(drw, method_name, path)
    finally:
        if opened_here:
            _close_model_keep_drawing(app, path, drw)
    if not ok:
        raise ExtError(f"{method_name} returned False for {path} "
                       f"-- check the path and that the model rebuilds cleanly")
    after = v(drw, "GetViewCount")
    return _ok(f"{projection}-angle standard views from {os.path.basename(path)} "
               f"| views {before}->{after}" + (" (model opened and closed for it)" if opened_here else ""))


def _ensure_model_open(app, path) -> bool:
    """Open `path` silently if it is not open yet and give the focus back to
    the active drawing. True if this call opened it."""
    if v(app, "GetOpenDocumentByName", path) is not None:
        return False
    drawing_title = v(model(app), "GetTitle")
    doc_type = 2 if path.lower().endswith(".sldasm") else 1
    doc, err, _ = call_out(app, "OpenDoc6", path, doc_type, 1, "", OUT, OUT)  # swOpenDocOptions_Silent
    if doc is None:
        raise ExtError(f"Could not open {path} for the drawing views (error {err})")
    call_out(app, "ActivateDoc3", drawing_title, False, 0, OUT)
    return True


def _close_model_keep_drawing(app, path, drw):
    """Close a model this tool opened and give the focus back to the drawing
    (CloseDoc activates whatever window SW picks next). The drawing keeps its
    views; SW renames an untitled drawing after its model once views exist, so
    the title is read again here, not remembered from before."""
    drawing_title = v(T(drw, "IModelDoc2"), "GetTitle")
    doc = v(app, "GetOpenDocumentByName", path)
    if doc is not None:
        v(app, "CloseDoc", v(doc, "GetTitle"))
    call_out(app, "ActivateDoc3", drawing_title, False, 0, OUT)


def t_add_drawing_view(app, args):
    drw = drawing(app)
    path = str(args["model_path"])
    if not os.path.exists(path):
        raise ExtError(f"Model file not found: {path}")
    view_name = _resolve_view_name(args.get("view"))
    x = float(args.get("x", 100)) / 1000.0
    y = float(args.get("y", 100)) / 1000.0
    before = v(drw, "GetViewCount")
    view = drw.CreateDrawViewFromModelView3(path, view_name, x, y, 0.0)
    if view is None:
        raise ExtError(
            f"CreateDrawViewFromModelView3 failed for view '{view_name}' of "
            f"{os.path.basename(path)} -- check the name (Front/Top/Right/Left/"
            f"Back/Bottom/Isometric/Dimetric/Trimetric/Current) and that the "
            f"model rebuilds cleanly")
    after = v(drw, "GetViewCount")
    iview = T(view, "IView")
    scale = args.get("scale")
    if scale:
        iview.ScaleDecimal = float(scale)
    return _ok(f"Added view {v(iview, "GetName2")!r} ({view_name}) at "
               f"({fnum(x*1000)},{fnum(y*1000)})mm | views {before}->{after}")


# CreateSectionViewAt5's Options bitmask (confirmed via swapi against this
# install's typelib -- swCreateSectionView_* constants).
_SECTION_VIEW_FLAGS = {
    "not_aligned": 1, "offset": 2, "change_direction": 4, "scale_with_model": 8,
    "partial": 16, "display_surface_cut": 32, "exclude_fasteners": 64,
    "cut_surface_bodies": 128,
}

# CreateDetailViewAt4's Style param (swDetView* constants).
_DETAIL_VIEW_STYLES = {"standard": 0, "broken": 1, "leader": 2, "noleader": 3, "connected": 4}


def _activate_view(drw, name: str):
    """ActivateView + the now-active IView (drw.ActiveDrawingView)."""
    if not drw.ActivateView(name):
        raise ExtError(f"View '{name}' not found on the active sheet "
                       f"(name is as reported by add_drawing_view/add_standard_views, "
                       f"e.g. 'Drawing View1')")
    view = drw.ActiveDrawingView
    if view is None:
        raise ExtError(f"ActivateView('{name}') succeeded but ActiveDrawingView is None")
    return T(view, "IView")


def _view_sketch_xform(view) -> List[float]:
    """ArrayData of SHEET (m) -> the view's own sketch space (m).

    Anything SketchManager.Create* draws while a view is active lands in that
    view's sketch, whose coordinates are NOT sheet coordinates: they are
    model-scale (sheet / view scale), with the origin at IView.Position (the
    view centre) -- confirmed live 2026-09-22. That is NOT generally the model
    origin's projection: a 1:5 Front view of an asymmetric ring had its
    Position at sheet x=150 and the model origin at x=119. The view sketch's
    ISketch.ModelToSketchTransform is exactly this sheet->sketch map (its
    'model' is the drawing sheet), scale and view rotation included, so no
    hand-rolled Position/ScaleDecimal math."""
    sk = v(view, "GetSketch")
    if sk is None:
        raise ExtError(f"View '{v(view, "GetName2")}' has no sketch")
    return list(T(T(sk, "ISketch").ModelToSketchTransform, "IMathTransform").ArrayData)


def sheet_to_view_sketch(a, x_mm: float, y_mm: float) -> List[float]:
    """Sheet mm -> view sketch metres, given _view_sketch_xform's ArrayData."""
    p = xform(a, [x_mm / 1000.0, y_mm / 1000.0, 0.0])
    return [p[0], p[1]]


def view_sketch_to_sheet(a, x_m: float, y_m: float) -> List[float]:
    """Inverse of sheet_to_view_sketch: view sketch metres -> sheet mm.
    Pure math on the ArrayData (rotation + uniform scale + shift), so
    selftest.py can check the pair offline."""
    s = a[12] if len(a) > 12 and a[12] else 1.0
    q = [(x_m - a[9]) / s, (y_m - a[10]) / s, (0.0 - a[11]) / s]
    # p' = (p·R)*s + t  =>  p = ((p'-t)/s)·R^T, i.e. p_i = sum_j R[i][j] q_j
    p = [sum(a[3 * i + j] * q[j] for j in range(3)) for i in range(3)]
    return [p[0] * 1000.0, p[1] * 1000.0]


def model_to_sheet(view, p_mm) -> List[float]:
    """Model point (mm, in the referenced part/assembly's own frame) -> sheet
    mm, via IView.ModelToViewTransform (whose 'view' space is sheet metres)."""
    a = list(T(view.ModelToViewTransform, "IMathTransform").ArrayData)
    p = xform(a, [c / 1000.0 for c in p_mm])
    return [p[0] * 1000.0, p[1] * 1000.0]


def _sketch_line_sheet(view, sm, x1, y1, x2, y2):
    """A cutting line for section views, given in SHEET mm, drawn into the
    active view's sketch (see _view_sketch_xform for why this converts).
    Returns (segment, worst round-trip error in sheet mm)."""
    a = _view_sketch_xform(view)
    s1, s2 = sheet_to_view_sketch(a, x1, y1), sheet_to_view_sketch(a, x2, y2)
    with NoInference(sm):
        line = sm.CreateLine(s1[0], s1[1], 0.0, s2[0], s2[1], 0.0)
    if line is None:
        raise ExtError("CreateLine returned None for the section cutting line")
    sl = T(line, "ISketchLine")
    err = 0.0
    for pt, want in ((v(sl, "GetStartPoint2"), (x1, y1)), (v(sl, "GetEndPoint2"), (x2, y2))):
        sp = T(pt, "ISketchPoint")
        got = view_sketch_to_sheet(a, sp.X, sp.Y)
        err = max(err, math.hypot(got[0] - want[0], got[1] - want[1]))
    return T(line, "ISketchSegment"), err


def _sketch_circle_sheet(view, sm, x, y, r):
    """A detail-view boundary circle given in SHEET mm (centre + radius),
    drawn into the active view's sketch. CreateCircle silently returns None
    without the NoInference (AddToDB/AutoInference) bracket -- see NOTES.md."""
    a = _view_sketch_xform(view)
    c = sheet_to_view_sketch(a, x, y)
    e = sheet_to_view_sketch(a, x + r, y)
    with NoInference(sm):
        circ = sm.CreateCircle(c[0], c[1], 0.0, e[0], e[1], 0.0)
    if circ is None:
        raise ExtError("CreateCircle returned None for the detail-view circle")
    return T(circ, "ISketchSegment")


def _discard_segment(drw, md, view_name, seg):
    """Best-effort cleanup of the marker line/circle when view creation fails,
    so a failed call does not litter the sheet. A view-sketch segment only
    deletes while ITS view is active -- otherwise Select4 still returns True
    but EditDelete/DeleteSelection2 silently do nothing (confirmed live
    2026-09-22). Failure here must not mask the original error, so it is
    swallowed."""
    try:
        drw.ActivateView(view_name)
        md.ClearSelection2(True)
        seg.Select4(False, nothing())
        T(md.Extension, "IModelDocExtension").DeleteSelection2(0)
    except Exception:
        pass


def _outline_mm(view) -> str:
    o = v(view, "GetOutline")
    return f"x {fnum(o[0]*1000)}..{fnum(o[2]*1000)}, y {fnum(o[1]*1000)}..{fnum(o[3]*1000)} mm"


def t_add_section_view(app, args):
    """CreateSectionViewAt5 does not build its own cutting plane from X,Y,Z as
    the name suggests -- confirmed live (2026-09-22): with nothing selected it
    returns None and adds no view. It needs a construction line, drawn ON THE
    ACTIVE VIEW (ActivateView first, same requirement as detail views) and
    selected, and X,Y,Z is only WHERE THE RESULTING SECTION VIEW LANDS on the
    sheet -- unrelated to the line's own position.

    The line itself lives in the active view's SKETCH space, not sheet space
    (fixed 2026-09-22: sheet mm used to be passed straight to CreateLine,
    which cut skewed and off-centre). The public API stays in sheet mm and
    _sketch_line_sheet converts."""
    drw = drawing(app)
    md = model(app)
    source = str(args["source_view"])
    view = _activate_view(drw, source)
    sm = T(md.SketchManager, "ISketchManager")
    x1, y1, x2, y2 = (float(args[k]) for k in ("x1", "y1", "x2", "y2"))

    flags = 0
    for name in parse_names(args.get("options", "")):
        bit = _SECTION_VIEW_FLAGS.get(name)
        if bit is None:
            raise ExtError(f"Unknown section option '{name}' (have: {sorted(_SECTION_VIEW_FLAGS)})")
        flags |= bit

    seg, err = _sketch_line_sheet(view, sm, x1, y1, x2, y2)
    md.ClearSelection2(True)
    if not seg.Select4(False, nothing()):
        _discard_segment(drw, md, source, seg)
        raise ExtError("Could not select the cutting line after drawing it")

    px, py = float(args["x"]) / 1000.0, float(args["y"]) / 1000.0
    label = str(args.get("label", ""))
    depth = float(args.get("depth", 0)) / 1000.0
    before = v(drw, "GetViewCount")
    try:
        new = drw.CreateSectionViewAt5(px, py, 0.0, label, flags, None, depth)
    except Exception:
        _discard_segment(drw, md, source, seg)
        raise
    if new is None:
        _discard_segment(drw, md, source, seg)
        raise ExtError(f"CreateSectionViewAt5 failed -- check that the line actually "
                       f"crosses '{source}' (x1,y1,x2,y2 are sheet mm, same space as "
                       f"add_drawing_view's x,y; the view spans {_outline_mm(view)})")
    after = v(drw, "GetViewCount")
    iview = T(new, "IView")
    warn = f" | ⚠ line landed {fnum(err, 3)}mm off the requested sheet points" if err > 0.01 else ""
    return _ok(f"Added section view {v(iview, "GetName2")!r} through "
               f"({fnum(x1)},{fnum(y1)})-({fnum(x2)},{fnum(y2)})mm sheet "
               f"of '{source}' | views {before}->{after}{warn}")


def t_add_detail_view(app, args):
    """CreateDetailViewAt4 needs the same ActivateView + pre-drawn/selected
    marker pattern as section views (a circle here, confirmed live 2026-09-22),
    and the same sheet -> view-sketch conversion for the circle.
    LabelIn is a STRING (the letter, '' = auto-next), not a boolean -- passing
    True/False there silently stamps the literal text 'True'/'False' as the
    label with no error, which is what a naive port of the signature's
    (bool-looking) position would do."""
    drw = drawing(app)
    md = model(app)
    source = str(args["source_view"])
    style = _DETAIL_VIEW_STYLES.get(str(args.get("style", "standard")).strip().lower())
    if style is None:
        raise ExtError(f"Unknown style '{args.get('style')}' (have: {sorted(_DETAIL_VIEW_STYLES)})")
    view = _activate_view(drw, source)
    sm = T(md.SketchManager, "ISketchManager")
    seg = _sketch_circle_sheet(view, sm, float(args["x"]), float(args["y"]), float(args["radius"]))
    md.ClearSelection2(True)
    if not seg.Select4(False, nothing()):
        _discard_segment(drw, md, source, seg)
        raise ExtError("Could not select the detail circle after drawing it")

    scale = float(args.get("scale", 2.0))
    px, py = float(args["place_x"]) / 1000.0, float(args["place_y"]) / 1000.0
    label = str(args.get("label", ""))

    before = v(drw, "GetViewCount")
    try:
        new = drw.CreateDetailViewAt4(
            px, py, 0.0, style, scale, 1.0, label, 0,
            bool(args.get("full_outline", True)),
            bool(args.get("jagged_outline", False)),
            bool(args.get("no_outline", False)),
            int(args.get("shape_intensity", 1)))
    except Exception:
        _discard_segment(drw, md, source, seg)
        raise
    if new is None:
        _discard_segment(drw, md, source, seg)
        raise ExtError(f"CreateDetailViewAt4 failed -- check that "
                       f"({fnum(args['x'])},{fnum(args['y'])})mm r={fnum(args['radius'])}mm "
                       f"actually falls on '{source}' ({_outline_mm(view)})")
    after = v(drw, "GetViewCount")
    iview = T(new, "IView")
    return _ok(f"Added detail view {v(iview, "GetName2")!r} of "
               f"({fnum(args['x'])},{fnum(args['y'])})mm r={fnum(args['radius'])}mm on "
               f"'{source}' at {fnum(scale)}:1 | views {before}->{after}")


def _half_depth_mm(view) -> float:
    """Half the part's extent along the view direction, in model mm -- the
    depth that cuts a broken-out section exactly to the middle plane. For a
    body of revolution seen from the side that is the outer radius (Ø1230 ring
    -> 615, which is what the user set by hand on kolco-flanec-1230 View9).
    Body boxes are axis-aligned, so this is exact for standard views of parts
    whose axes follow the model frame."""
    rmd = T(view.ReferencedDocument, "IModelDoc2") if view.ReferencedDocument is not None else None
    if rmd is None or not is_part(rmd):
        raise ExtError(f"View '{v(view, "GetName2")}' does not show a part -- give depth explicitly")
    a = list(T(view.ModelToViewTransform, "IMathTransform").ArrayData)
    s = a[12] if len(a) > 12 and a[12] else 1.0
    zs = []
    for b in bodies(rmd):
        bx = v(b, "GetBodyBox")
        for i in (0, 3):
            for j in (1, 4):
                for k in (2, 5):
                    zs.append(xform(a, [bx[i], bx[j], bx[k]])[2])
    if not zs:
        raise ExtError(f"Part in '{v(view, "GetName2")}' has no bodies")
    return (max(zs) - min(zs)) / s / 2.0 * 1000.0


def _broken_out_names(md, view_feature: str) -> set:
    out = set()
    f = v(md, "FirstFeature")
    while f is not None:
        ff = T(f, "IFeature")
        if v(ff, "GetTypeName2") == "DrSheet":
            sub = v(ff, "GetFirstSubFeature")
            while sub is not None:
                sf = T(sub, "IFeature")
                if sf.Name == view_feature:
                    ss = v(sf, "GetFirstSubFeature")
                    while ss is not None:
                        s3 = T(ss, "IFeature")
                        if v(s3, "GetTypeName2") == "DrBreakoutSectionLine":
                            out.add(s3.Name)
                        ss = v(s3, "GetNextSubFeature")
                sub = v(sf, "GetNextSubFeature")
        f = v(ff, "GetNextFeature")
    return out


def t_add_broken_out_section(app, args):
    """Broken-out section on an existing view: a closed contour drawn in the
    view's sketch, selected, then IDrawingDoc.CreateBreakOutSection(depth).
    Depth is a MODEL distance from the nearest point of the part. The user's
    way to section a body of revolution (2026-09-23): ONE side view + a
    broken-out section around the whole view, depth = outer radius -- a full
    hatched section with no end view, section line or 'A-A' label. So the
    defaults are: contour = the view outline + margin, depth = half the part's
    extent along the view direction (= the outer radius for such a part).
    After creation the contour is absorbed by the feature: IBrokenOutSection-
    FeatureData.SketchSegment reads back None and its sketches report 0
    segments, so the contour cannot be inspected afterwards (confirmed live
    2026-09-23) -- the depth can (Depth, metres)."""
    drw = drawing(app)
    md = model(app)
    view, fname = _find_view(drw, md, str(args["view"]))
    disp = v(view, "GetName2")

    depth = args.get("depth")
    depth_mm = float(depth) if depth not in (None, "", 0) else _half_depth_mm(view)
    if depth_mm <= 0:
        raise ExtError(f"depth must be > 0 mm (got {fnum(depth_mm)})")

    pts = args.get("points")
    if pts:
        poly = json.loads(pts) if isinstance(pts, str) else pts
        if len(poly) < 3:
            raise ExtError("points: need at least 3 [x,y] sheet-mm pairs for a closed contour")
    else:
        o = [c * 1000.0 for c in v(view, "GetOutline")]
        mg = float(args.get("margin", 2))
        poly = [[o[0] - mg, o[1] - mg], [o[2] + mg, o[1] - mg],
                [o[2] + mg, o[3] + mg], [o[0] - mg, o[3] + mg]]

    view = _activate_view(drw, disp)
    sm = T(md.SketchManager, "ISketchManager")
    a = _view_sketch_xform(view)
    sk_pts = [sheet_to_view_sketch(a, float(p[0]), float(p[1])) for p in poly]
    segs = []
    with NoInference(sm):
        for i in range(len(sk_pts)):
            p, q = sk_pts[i], sk_pts[(i + 1) % len(sk_pts)]
            ln = sm.CreateLine(p[0], p[1], 0.0, q[0], q[1], 0.0)
            if ln is None:
                for s in segs:
                    _discard_segment(drw, md, disp, s)
                raise ExtError(f"CreateLine returned None for contour edge {i + 1}")
            segs.append(T(ln, "ISketchSegment"))

    md.ClearSelection2(True)
    for i, s in enumerate(segs):
        if not s.Select4(i > 0, nothing()):
            for s2 in segs:
                _discard_segment(drw, md, disp, s2)
            raise ExtError("Could not select the contour after drawing it")

    before = _broken_out_names(md, fname)
    ok = drw.CreateBreakOutSection(depth_mm / 1000.0)
    md.ClearSelection2(True)
    new = _broken_out_names(md, fname) - before
    if not ok or not new:
        for s in segs:
            _discard_segment(drw, md, disp, s)
        raise ExtError(f"CreateBreakOutSection({fnum(depth_mm)}mm) returned {ok!r}, no new "
                       f"feature -- check the contour encloses part of '{disp}' "
                       f"({_outline_mm(view)}) and depth stays inside the part")
    how = "given" if depth not in (None, "", 0) else "half the part depth along the view"
    return _ok(f"Added {', '.join(sorted(new))} to '{disp}'"
               + (f" ({fname})" if fname != disp else "")
               + f" | depth {fnum(depth_mm)}mm ({how}) | contour {len(poly)} pts, sheet mm")


def _all_view_names(drw) -> List[str]:
    """Every real view on the sheet (GetFirstView is the sheet itself -- skipped)."""
    vw = v(drw, "GetFirstView")
    if vw is None:
        return []
    vw = v(T(vw, "IView"), "GetNextView")
    names = []
    while vw is not None:
        vt = T(vw, "IView")
        names.append(vt.Name)
        vw = v(vt, "GetNextView")
    return names


def _force_view_render(md):
    """Freshly-placed views (straight after add_standard_views/add_drawing_view)
    silently starve InsertModelDimensions until SolidWorks actually renders
    them. Confirmed live 2026-09-22 by elimination -- none of these fixed it:
    waiting up to ~7s across repeated select+call retries; ForceRebuild3;
    GraphicsRedraw2 + ViewZoomtofit2; switching ActiveDoc to a different open
    document and back. What DID fix it, reliably, every time: calling the
    capture_view tool (SaveBMP) in between. So this forces that same render
    side effect directly -- a throwaway bitmap, written and discarded, only
    for the render it forces."""
    fd, path = tempfile.mkstemp(suffix=".bmp")
    os.close(fd)
    try:
        md.SaveBMP(path, 200, 200)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def t_insert_model_dimensions(app, args):
    """InsertModelDimensions(Option) pulls dimensions only into whatever
    view(s) are SELECTED at call time (SelectByID2, kind='DRAWINGVIEW') --
    ActivateView alone inserts nothing, confirmed live 2026-09-22 (several
    Option values tried against an activated-but-unselected view, all
    silent no-ops). Each selected view only gets the dimensions native to
    geometry actually visible from it -- a feature whose defining geometry
    isn't shown edge-on in any placed view (e.g. this install's test hole,
    always seen as a centre mark, never its round profile) will not appear
    regardless of Option or how many views are selected. See
    _force_view_render for the freshly-placed-view starvation issue this
    also works around."""
    drw = drawing(app)
    md = model(app)
    ext_ = T(md.Extension, "IModelDocExtension")
    names = parse_names(args.get("views", ""))
    if not names:
        names = _all_view_names(drw)
    if not names:
        raise ExtError("No views on the active sheet (add_drawing_view/add_standard_views first)")
    option = int(args.get("option", 0))

    _force_view_render(md)
    md.ClearSelection2(True)
    for i, name in enumerate(names):
        if not ext_.SelectByID2(name, "DRAWINGVIEW", 0, 0, 0, i > 0, 0, nothing(), 0):
            raise ExtError(f"View '{name}' not found on the active sheet")
    drw.InsertModelDimensions(option)  # void return -- no success signal from the API itself
    md.ClearSelection2(True)
    return _ok(f"Inserted model dimensions into {len(names)} view(s): {', '.join(names)} "
              f"(InsertModelDimensions gives no success signal -- verify with capture_view)")


def t_add_note(app, args):
    """IDrawingDoc.NewNote / InsertNewNote2 are void-return UI-command
    wrappers that arm the 'place a note' cursor for a mouse click that never
    comes when called headlessly -- confirmed live 2026-09-22: no object, no
    new selection, nothing added to the sheet, and no error either.
    IModelDoc2.InsertNote(Text) is the one that actually returns an INote;
    position it via its IAnnotation.SetPosition (metres). A position outside
    the sheet's own paper bounds still shows on screen (zoom-to-fit expands
    to include it) but is cropped out of a PDF export -- keep x,y within the
    sheet's paper size."""
    md = model(app)
    if v(md, "GetType") != 3:
        raise ExtError(f"Active document '{v(md, "GetTitle")}' is not a drawing")
    text = str(args["text"])
    n = md.InsertNote(text)
    if n is None:
        raise ExtError("InsertNote returned None")
    note = T(n, "INote")
    ann = T(v(note, "GetAnnotation"), "IAnnotation")
    x, y = float(args.get("x", 100)) / 1000.0, float(args.get("y", 100)) / 1000.0
    ann.SetPosition(x, y, 0.0)
    height = args.get("height")
    if height:
        note.SetHeight(float(height) / 1000.0)
    return _ok(f"Added note {v(note, "GetName")!r} at ({fnum(x*1000)},{fnum(y*1000)})mm: {text[:60]!r}")


def t_export_pdf(app, args):
    md = model(app)
    if v(md, "GetType") != 3:
        raise ExtError(f"Active document '{v(md, "GetTitle")}' is not a drawing")
    path = str(args["path"])
    if not path.lower().endswith(".pdf"):
        raise ExtError("path must end in .pdf")
    out_dir = os.path.dirname(path)
    if out_dir and not os.path.isdir(out_dir):
        raise ExtError(f"Directory does not exist: {out_dir}")
    ok = md.SaveAs(path)
    if not ok or not os.path.exists(path):
        raise ExtError(f"SaveAs returned {ok!r}, file "
                       f"{'exists' if os.path.exists(path) else 'was not created'}")
    return _ok(f"Exported PDF: {path} ({os.path.getsize(path)} bytes)")


# ----------------------------------------------------------------------------
# Drawing views: lookup, move, delete
# ----------------------------------------------------------------------------

def _views(drw) -> list:
    """Every real view on the active sheet as IView (the sheet itself,
    GetFirstView, skipped)."""
    out = []
    first = v(drw, "GetFirstView")
    n = v(T(first, "IView"), "GetNextView") if first is not None else None
    while n is not None:
        vt = T(n, "IView")
        out.append(vt)
        n = v(vt, "GetNextView")
    return out


def _view_feature_names(md) -> Dict[str, str]:
    """{view display name: view FEATURE name}. They differ for derived views:
    a section view shows as 'Section View A-A' (IView.GetName2/Name) while its
    feature -- the name annotation selection strings need, 'RD1@Drawing View3'
    -- is 'Drawing View3' (confirmed live 2026-09-22). Read off the sheet
    feature's subfeatures, whose GetSpecificFeature2 is the IView."""
    out = {}
    f = v(md, "FirstFeature")
    while f is not None:
        ff = T(f, "IFeature")
        if v(ff, "GetTypeName2") == "DrSheet":
            sub = v(ff, "GetFirstSubFeature")
            while sub is not None:
                sf = T(sub, "IFeature")
                spec = v(sf, "GetSpecificFeature2")
                if spec is not None:
                    try:
                        out[v(T(spec, "IView"), "GetName2")] = sf.Name
                    except Exception:
                        pass  # not a view (Sketch1, Plane1, Detail Folder...)
                sub = v(sf, "GetNextSubFeature")
        f = v(ff, "GetNextFeature")
    return out


def _find_view(drw, md, name: str):
    """(IView, feature name) by display OR feature name."""
    fnames = _view_feature_names(md)
    for vw in _views(drw):
        disp = v(vw, "GetName2")
        if name in (disp, fnames.get(disp)):
            return vw, fnames.get(disp, disp)
    have = ", ".join(f"{d!r}" + (f" (={fnames[d]!r})" if fnames.get(d, d) != d else "")
                     for d in (v(vw, "GetName2") for vw in _views(drw)))
    raise ExtError(f"No view '{name}' on the active sheet (have: {have or 'none'})")


def t_move_drawing_view(app, args):
    """IView.Position must be a VARIANT(VT_ARRAY|VT_R8) (com.double_array): a
    plain tuple is accepted without error and sets garbage -- (0.31,0.2) came
    back as (0, 310) mm, confirmed live 2026-09-22."""
    drw = drawing(app)
    md = model(app)
    view, _ = _find_view(drw, md, str(args["view"]))
    x, y = float(args["x"]), float(args["y"])
    before = [p * 1000 for p in view.Position]
    view.Position = com.double_array([x / 1000.0, y / 1000.0])
    after = [p * 1000 for p in view.Position]
    warn = ""
    if abs(after[0] - x) > 0.01 or abs(after[1] - y) > 0.01:
        warn = " | ⚠ position did not take (aligned views are locked to their parent's axis)"
    return _ok(f"Moved '{v(view, "GetName2")}' ({fnum(before[0])},{fnum(before[1])}) -> "
               f"({fnum(after[0])},{fnum(after[1])})mm{warn}")


def _segment_names(view) -> set:
    sk = v(view, "GetSketch")
    segs = v(T(sk, "ISketch"), "GetSketchSegments") if sk is not None else None
    return {v(T(s, "ISketchSegment"), "GetName") for s in (segs or ())}


def t_delete_drawing_view(app, args):
    """Deleting a section view leaves its section line in the parent view as
    an orphan (GetSectionLineCount2 still counts it, IDrSection.GetSectionView
    -> None). Deleting that line (SelectByID2(name, 'SECTIONLINE')) in turn
    hands the original cutting line back to the parent view's sketch as a
    loose segment -- all confirmed live 2026-09-22. This removes all three."""
    drw = drawing(app)
    md = model(app)
    ext_ = T(md.Extension, "IModelDocExtension")
    view, fname = _find_view(drw, md, str(args["view"]))
    disp = v(view, "GetName2")

    # section lines that belong to this view, per parent view, BEFORE deleting
    owned = []  # (parent IView, parent display name, section line name)
    for pv in _views(drw):
        for s in v(pv, "GetSectionLines") or ():
            ds = T(s, "IDrSection")
            sv = v(ds, "GetSectionView")
            if sv is not None and v(T(sv, "IView"), "GetName2") == disp:
                owned.append((pv, v(pv, "GetName2"), v(ds, "GetName")))

    before = v(drw, "GetViewCount")
    md.ClearSelection2(True)
    if not ext_.SelectByID2(fname, "DRAWINGVIEW", 0, 0, 0, False, 0, nothing(), 0):
        raise ExtError(f"Could not select view '{disp}' ({fname})")
    if not ext_.DeleteSelection2(0):
        raise ExtError(f"DeleteSelection2 refused view '{disp}'")
    after = v(drw, "GetViewCount")

    notes = []
    for pv, pname, lname in owned:
        segs_before = _segment_names(pv)
        md.ClearSelection2(True)
        if not (ext_.SelectByID2(lname, "SECTIONLINE", 0, 0, 0, False, 0, nothing(), 0)
                and ext_.DeleteSelection2(0)):
            notes.append(f"⚠ section line {lname!r} left in '{pname}'")
            continue
        loose = _segment_names(pv) - segs_before
        if loose:
            drw.ActivateView(pname)  # view-sketch segments only delete while their view is active
            md.ClearSelection2(True)
            sk = T(v(pv, "GetSketch"), "ISketch")
            for s in v(sk, "GetSketchSegments") or ():
                seg = T(s, "ISketchSegment")
                if v(seg, "GetName") in loose:
                    seg.Select4(True, nothing())
            ext_.DeleteSelection2(0)
            left = _segment_names(pv) & loose
            if left:
                notes.append(f"⚠ cutting line {sorted(left)} left in '{pname}' sketch")
        notes.append(f"removed section line {lname!r} + cutting line from '{pname}'")
    md.ClearSelection2(True)
    return _ok(f"Deleted view '{disp}'" + (f" ({fname})" if fname != disp else "")
               + f" | views {before}->{after}" + "".join(f" | {n}" for n in notes))


# ----------------------------------------------------------------------------
# Drawing annotations: dimensions, gtol, datums, surface finish
# ----------------------------------------------------------------------------

def _ref_part(view):
    ref = view.ReferencedDocument
    if ref is None:
        raise ExtError(f"View '{v(view, "GetName2")}' has no referenced model")
    rmd = T(ref, "IModelDoc2")
    if not is_part(rmd):
        raise ExtError(f"View '{v(view, "GetName2")}' shows an assembly -- edge references "
                       f"are only supported for part views")
    return rmd


def _resolve_edge(view, spec):
    """An edge of the view's part, from either
      - an index: '12' -- same numbering as list_edges on that part, or
      - a model point: '0,155,615' (mm, part frame) -- nearest edge to it.
    Returns (IEdge, point ON the edge in model mm)."""
    rmd = _ref_part(view)
    s = str(spec).strip()
    if "," not in s:
        idx = int(s)
        rows = edge_rows(rmd)
        if not 1 <= idx <= len(rows):
            raise ExtError(f"edge index {idx} out of range 1..{len(rows)} "
                           f"(list_edges on {v(rmd, "GetTitle")})")
        r = rows[idx - 1]
        return r["obj"], [c * 1000 for c in r["mid"]]
    p = [float(c) / 1000.0 for c in s.split(",")]
    if len(p) != 3:
        raise ExtError(f"edge point must be 'x,y,z' in model mm, got {spec!r}")
    best = None
    for b in bodies(rmd):
        for eo in v(b, "GetEdges") or ():
            e = T(eo, "IEdge")
            q = e.GetClosestPointOn(*p)[:3]
            d = math.dist(p, q)
            if best is None or d < best[0]:
                best = (d, e, q)
    if best is None:
        raise ExtError(f"{v(rmd, "GetTitle")} has no edges")
    d, e, q = best
    if d > 0.001:
        logger.info(f"edge point {spec} is {d*1000:.2f}mm from the nearest edge")
    return e, [c * 1000 for c in q]


def _select_edges(md, view, specs) -> List[List[float]]:
    """Select edges of a view by IView.SelectEntity(model edge) -- the model's
    own IEdge, no sheet-coordinate picking. Picking with SelectByID2('EDGE') at
    the edge's mapped sheet point misses short edges at zoom-to-fit (reported
    2026-09-22); a live retest the same day got the VIEW back (type 12) even
    after ViewZoomTo2 around the point, so picking is not used at all. The one
    cost: SolidWorks, not the caller, picks where on the edge a leader lands.
    Returns the model points (mm) on each edge."""
    md.ClearSelection2(True)
    pts = []
    for i, spec in enumerate(specs):
        e, p = _resolve_edge(view, spec)
        if not view.SelectEntity(e, i > 0):
            raise ExtError(f"IView.SelectEntity failed for edge {spec!r} in '{v(view, "GetName2")}' "
                           f"(is it visible in this view?)")
        pts.append(p)
    return pts


def _place(ann, x, y):
    if x is None or y is None:
        return
    if not T(ann, "IAnnotation").SetPosition2(float(x) / 1000.0, float(y) / 1000.0, 0.0):
        raise ExtError("IAnnotation.SetPosition2 failed")


def _ann_pos(ann) -> str:
    p = v(T(ann, "IAnnotation"), "GetPosition")
    return f"({fnum(p[0]*1000)},{fnum(p[1]*1000)})mm"


def _default_text_pos(view, pts, off=10.0):
    """Sheet mm point `off` mm above the midpoint of the picked model points."""
    s = [model_to_sheet(view, p) for p in pts]
    return (sum(q[0] for q in s) / len(s), max(q[1] for q in s) + off)


def t_add_drawing_dimension(app, args):
    """Driven dimension between one or two model edges in a drawing view.
    One circular edge -> diameter/radius; two -> distance. Created via
    IModelDoc2.AddDimension2 with the edges selected through
    IView.SelectEntity; x,y (sheet mm) is the text position and also decides
    horizontal/vertical/aligned, exactly like placing it by hand.

    Tolerance, confirmed live 2026-09-22: IDimensionTolerance.SetValues2 returns
    False in a drawing (both WhichConfigurations values); the older
    SetValues(min, max) works. Fit: Type=swTolFIT(7) + SetFitValues(hole, shaft).
    Drawing-made dimensions are driven (reference) and GOST wraps them in
    parentheses -- parenthesis=False (default) clears ShowParenthesis."""
    drw = drawing(app)
    md = model(app)
    view, _ = _find_view(drw, md, str(args["view"]))
    specs = [s for s in (args.get("edge1"), args.get("edge2")) if s not in (None, "")]
    if not specs:
        raise ExtError("edge1 is required (edge index from list_edges on the part, or 'x,y,z' model mm)")
    pts = _select_edges(md, view, specs)
    x, y = args.get("x"), args.get("y")
    if x is None or y is None:
        x, y = _default_text_pos(view, pts)
    dd = md.AddDimension2(float(x) / 1000.0, float(y) / 1000.0, 0.0)
    md.ClearSelection2(True)
    if dd is None:
        raise ExtError(f"AddDimension2 returned None for edges {specs} in '{v(view, "GetName2")}'")
    D = T(dd, "IDisplayDimension")
    dim = T(D.GetDimension2(0), "IDimension")

    parts = []
    prefix, suffix = args.get("prefix"), args.get("suffix")
    if prefix:
        D.SetText(1, str(prefix))  # swDimensionTextPrefix
    if suffix:
        D.SetText(2, str(suffix))  # swDimensionTextSuffix
    D.ShowParenthesis = bool(args.get("parenthesis", False))

    tol = T(dim.Tolerance, "IDimensionTolerance")
    up, lo, fit = args.get("tol_upper"), args.get("tol_lower"), args.get("fit")
    if fit:
        f = str(fit).strip()
        if "/" in f:
            hole, shaft = f.split("/", 1)       # 'H7/g6'
        elif f[:1].isupper():
            hole, shaft = f, ""                 # 'H7' -- hole
        else:
            hole, shaft = "", f                 # 'h7' -- shaft
        tol.Type = 8 if (up is not None or lo is not None) else 7  # swTolFITWITHTOL / swTolFIT
        if not tol.SetFitValues(hole, shaft):
            parts.append(f"⚠ SetFitValues({hole!r},{shaft!r}) failed")
        else:
            parts.append(f"fit {f}")
    if up is not None or lo is not None:
        u, l = float(up or 0.0), float(lo or 0.0)
        if fit is None:
            tol.Type = 4 if abs(u + l) < 1e-9 and u > 0 else 2  # swTolSYMMETRIC / swTolBILAT
        if not tol.SetValues(l / 1000.0, u / 1000.0):
            parts.append("⚠ IDimensionTolerance.SetValues failed")
        else:
            parts.append(f"tol {'+' if u >= 0 else ''}{fnum(u, 3)}/{fnum(l, 3)}")
    v(md, "GraphicsRedraw2")

    val = dim.SystemValue * 1000.0
    warn = " | ⚠ value is 0 -- edges probably coincide in this view" if abs(val) < 1e-6 else ""
    ann = v(D, "GetAnnotation")
    return _ok(f"Added dimension {v(T(ann, 'IAnnotation'), "GetName")!r} = {fnum(val, 3)}mm "
               f"({dim.FullName}) in '{v(view, "GetName2")}' at {_ann_pos(ann)}"
               + (f" | {prefix!r} prefix" if prefix else "")
               + "".join(f" | {p}" for p in parts) + warn)


# ISO geometric-tolerance symbols, names as in <SW data>/lang/english/gtol.sym
# (#IGTOL section; #GGTOL is the GOST set with the same names plus AXIS/LONG).
_GTOL_ALIASES = {
    "runout": "SRUN", "circular_runout": "SRUN", "total_runout": "TRUN",
    "cylindricity": "CYL", "parallelism": "PARA", "perpendicularity": "PERP",
    "flatness": "FLAT", "position": "POSI", "concentricity": "CONC",
    "coaxiality": "CONC", "circularity": "CIRC", "roundness": "CIRC",
    "straightness": "STRAIGHT", "symmetry": "SYMMETRY", "angularity": "ANGULAR",
    "profile_line": "LPROF", "profile_surface": "SPROF",
}


def _gtol_symbol(sym: str, library: str) -> str:
    s = str(sym).strip()
    if s.startswith("<"):
        return s
    s = _GTOL_ALIASES.get(s.lower(), s.upper())
    return f"<{library.upper()}-{s}>"


def _gtols(view) -> list:
    out, g = [], v(view, "GetFirstGTOL")
    while g is not None:
        gt = T(g, "IGtol")
        out.append(gt)
        g = v(gt, "GetNextGTOL")
    return out


def t_add_gtol(app, args):
    """Feature control frame, optionally leadered to a model edge.

    Confirmed live 2026-09-22 (SW 2026):
    - SetFrameSymbols2's GCS is a STRING (typelib VT_BSTR), the symbol's name
      from gtol.sym ('<IGTOL-SRUN>'), not a swGcs* int: 25 renders literally.
    - InsertGtol gives a legacy-format frame (GetFormat()==1 == GTOL_SW2021).
      Set symbols and values on THAT, then ConvertFormat() -> GTOL_SW2022 with
      everything kept. Converting FIRST is what breaks: the tolerance value is
      lost (GetFrameValues -> None) or SetFrameValues2 throws 'server threw
      an exception', because the handle is stale after conversion -- re-fetch
      it (view.GetFirstGTOL chain, by annotation name) before touching it."""
    drw = drawing(app)
    md = model(app)
    view, _ = _find_view(drw, md, str(args["view"]))
    edge = args.get("edge")
    if edge not in (None, ""):
        _select_edges(md, view, [edge])
    else:
        md.ClearSelection2(True)
    g = v(md, "InsertGtol")
    md.ClearSelection2(True)
    if g is None:
        raise ExtError("InsertGtol returned None")
    G = T(g, "IGtol")
    name = v(T(v(G, "GetAnnotation"), "IAnnotation"), "GetName")

    sym = _gtol_symbol(args["symbol"], str(args.get("library", "IGTOL")))
    G.SetFrameSymbols2(1, sym, bool(args.get("diameter", False)), "", False, "", "", "", "")
    tolv = str(args.get("tolerance", ""))
    datums = parse_names(args.get("datum", ""))[:3]
    datums += [""] * (3 - len(datums))
    if not G.SetFrameValues2(1, tolv, "", *datums):
        raise ExtError("SetFrameValues2 returned False")
    if bool(args.get("convert", True)) and v(G, "GetFormat") == 1:
        v(G, "ConvertFormat")
        fresh = [x for x in _gtols(view) if v(T(v(x, "GetAnnotation"), "IAnnotation"), "GetName") == name]
        G = fresh[0] if fresh else G

    _place(v(G, "GetAnnotation"), args.get("x"), args.get("y"))
    got_vals = G.GetFrameValues(1) or ()
    got_sym = (G.GetFrameSymbols3(1) or ("",))[0]
    warn = ""
    if (got_vals[:1] or ("",))[0] != tolv or got_sym != sym:
        warn = f" | ⚠ read back symbol={got_sym!r} values={got_vals!r}"
    return _ok(f"Added gtol {name!r} {sym} {tolv} {' '.join(d for d in datums if d)} "
               f"in '{v(view, "GetName2")}' (format {v(G, "GetFormat")}) at {_ann_pos(v(G, "GetAnnotation"))}"
               + (f" on edge {edge}" if edge not in (None, "") else "") + warn)


def t_add_datum(app, args):
    """Datum feature symbol on a model edge. InsertDatumTag2 attaches to the
    selection (IView.SelectEntity), SetLabel sets the letter."""
    drw = drawing(app)
    md = model(app)
    view, _ = _find_view(drw, md, str(args["view"]))
    _select_edges(md, view, [args["edge"]])
    dt = v(md, "InsertDatumTag2")
    md.ClearSelection2(True)
    if dt is None:
        raise ExtError("InsertDatumTag2 returned None")
    DT = T(dt, "IDatumTag")
    label = str(args.get("label", "")).strip()
    if label and not DT.SetLabel(label):
        raise ExtError(f"SetLabel({label!r}) failed")
    ann = v(DT, "GetAnnotation")
    _place(ann, args.get("x"), args.get("y"))
    return _ok(f"Added datum {v(DT, "GetLabel")!r} ({v(T(ann, 'IAnnotation'), "GetName")}) on edge "
               f"{args['edge']} in '{v(view, "GetName2")}' at {_ann_pos(ann)}")


_SF_SYMBOLS = {"basic": 0, "machined": 1, "no_machining": 2}  # swSFBasic / swSFMachining_Req / swSFDont_Machine


def t_add_surface_finish(app, args):
    """Surface finish symbol, optionally leadered to a model edge.

    Which ISFSymbol text slot is DRAWN depends on the document's surface
    finish standard (swDetailingSFSymbolStandard, pref 629), not on the
    drafting standard -- confirmed live 2026-09-22 on two GOST drawings:
      0 ISO 1302:1992     -> slots 1-7 render; MaxRoughness lands in slot 5 -- OK
      1 ISO 1302:2002     -> slots 1,2,8,9,10 render
      2 ISO 21920-1       -> slots 2,8,9,10 render
    so the roughness goes to slot 5 under 1992 and slot 8
    (swSFSymbolRoughnessValue1) otherwise. Text in a non-rendering slot is
    stored and read back fine -- only the picture tells.
    InsertSurfaceFinishSymbol3's LocX/Y are ignored for a symbol with no
    leader (it lands at 0,0), hence SetPosition2 afterwards."""
    drw = drawing(app)
    md = model(app)
    ext_ = T(md.Extension, "IModelDocExtension")
    view, _ = _find_view(drw, md, str(args["view"]))
    edge = args.get("edge")
    attached = edge not in (None, "")
    if attached:
        pts = _select_edges(md, view, [edge])
    else:
        md.ClearSelection2(True)
    kind = str(args.get("symbol", "machined")).strip().lower()
    if kind not in _SF_SYMBOLS:
        raise ExtError(f"Unknown symbol '{kind}' (have: {sorted(_SF_SYMBOLS)})")
    x, y = args.get("x"), args.get("y")
    if (x is None or y is None) and attached:
        x, y = _default_text_pos(view, pts, 5.0)
    if x is None or y is None:
        raise ExtError("x,y are required for a surface finish symbol without an edge")
    sf = ext_.InsertSurfaceFinishSymbol3(_SF_SYMBOLS[kind], 1 if attached else 0,
                                         float(x) / 1000.0, float(y) / 1000.0, 0.0,
                                         0, 0, "", "", "", "", "", "", "")
    md.ClearSelection2(True)
    if sf is None:
        raise ExtError("InsertSurfaceFinishSymbol3 returned None")
    S = T(sf, "ISFSymbol")
    std = ext_.GetUserPreferenceInteger(629, 0)  # swDetailingSFSymbolStandard
    slot = 5 if std == 0 else 8
    value = str(args.get("value", ""))
    if value and not S.SetText(slot, value):
        raise ExtError(f"ISFSymbol.SetText({slot}, {value!r}) failed")
    _place(v(S, "GetAnnotation"), x, y)
    std_name = {0: "ISO 1302:1992", 1: "ISO 1302:2002", 2: "ISO 21920-1"}.get(std, f"#{std}")
    return _ok(f"Added surface finish {v(T(v(S, "GetAnnotation"), 'IAnnotation'), "GetName")!r} "
               f"{value!r} (slot {slot}, SF standard {std_name}) in '{v(view, "GetName2")}' "
               f"at {_ann_pos(v(S, "GetAnnotation"))}" + (f" on edge {edge}" if attached else ""))


_ANNOTATION_KINDS = {"dimension": "DIMENSION", "datum": "DATUMTAG", "gtol": "GTOL",
                     "surface_finish": "SFSYMBOL", "note": "NOTE"}


def t_delete_annotation(app, args):
    """IAnnotation.Select3 returns False for dims/datum tags in a drawing view
    (confirmed live 2026-09-22); SelectByID2('<name>@<view FEATURE name>',
    TYPE) + DeleteSelection2 works. The view part must be the feature name
    ('Drawing View3'), not a section view's display name -- resolved here."""
    drw = drawing(app)
    md = model(app)
    ext_ = T(md.Extension, "IModelDocExtension")
    view, fname = _find_view(drw, md, str(args["view"]))
    kind = str(args.get("kind", "dimension")).strip().lower()
    typ = _ANNOTATION_KINDS.get(kind)
    if typ is None:
        raise ExtError(f"Unknown kind '{kind}' (have: {sorted(_ANNOTATION_KINDS)})")
    names = parse_names(args["names"])
    done, missing = [], []
    for nm in names:
        md.ClearSelection2(True)
        if ext_.SelectByID2(f"{nm}@{fname}", typ, 0, 0, 0, False, 0, nothing(), 0) and ext_.DeleteSelection2(0):
            done.append(nm)
        else:
            missing.append(nm)
    md.ClearSelection2(True)
    msg = f"Deleted {len(done)} {kind}(s) from '{v(view, "GetName2")}': {', '.join(done) or '-'}"
    if missing:
        msg += f" | ⚠ not found as {typ} in {fname}: {', '.join(missing)}"
    return _ok(msg)


# ============================================================================
# Assembly tree
# ============================================================================

# swComponentSuppressionState_e
_COMP_STATE = {0: "suppressed", 1: "lightweight", 2: "resolved", 3: "resolved",
               4: "lightweight", 5: "lightweight"}


class _Reader:
    """Reads members of many objects of ONE interface with dispids looked up
    once (one Invoke per value instead of GetIDsOfNames + Invoke). Falls back
    to v() if an object answers differently."""

    def __init__(self):
        self._ids: Dict[str, int] = {}

    def __call__(self, obj, name, *args):
        ole = obj._oleobj_
        dispid = self._ids.get(name)
        if dispid is None:
            dispid = self._ids[name] = ole.GetIDsOfNames(0, name)
        try:
            ret = ole.Invoke(dispid, 0, com._GET_OR_CALL, True, *args)
        except pythoncom.com_error:
            return v(obj, name, *args)
        return com._good(ret, name)


def _comp_file_code(path: str) -> str:
    return os.path.basename(path or "") or "?"


class _RefResolver:
    """Where a component's file really is. A suppressed component is never
    loaded, so GetPathName returns the path stored at the author's last save
    (C:\\Users\\Administrator\\Desktop\\... on a model from elsewhere) --
    os.path.exists on it says "missing" for files that sit right next to the
    assembly (8320A-ZZ-01, 2026-09-26: 3 false alarms). SolidWorks itself
    looks in the referencing document's folder and the top assembly's folder
    (also under the stored path's trailing sub-folders) before the stored
    path, and so does this:
      ok        -- the stored path exists
      relocated -- stale stored path, the file is found where SW will look
      missing   -- found nowhere SW looks; `found` then holds a same-name
                   file elsewhere under the assembly folder, if any (a hint
                   only: SW won't pick that one up by itself)."""

    _WALK_LIMIT = 20000  # files; a model folder is hundreds

    def __init__(self, asm_path: str):
        self.asm_dir = os.path.dirname(asm_path or "")
        self._cache: Dict[tuple, tuple] = {}
        self._index: Optional[Dict[str, str]] = None

    def _by_name(self, base: str) -> str:
        if self._index is None:
            self._index, n = {}, 0
            if self.asm_dir:
                for root, _dirs, files in os.walk(self.asm_dir):
                    for f in files:
                        self._index.setdefault(f.lower(), os.path.join(root, f))
                    n += len(files)
                    if n > self._WALK_LIMIT:
                        break
        return self._index.get(base.lower(), "")

    def __call__(self, stored: str, parent_path: str = "") -> tuple:
        """-> (status, found_path)"""
        if not stored:
            return "missing", ""
        pdir = os.path.dirname(parent_path or "")
        key = (stored.lower(), pdir.lower())
        hit = self._cache.get(key)
        if hit:
            return hit
        if os.path.exists(stored):
            hit = ("ok", stored)
        else:
            parts = [p for p in stored.replace("/", "\\").split("\\") if p]
            dirs = list(dict.fromkeys(d for d in (pdir, self.asm_dir) if d))
            found = next((c for d in dirs for k in range(1, min(len(parts), 3) + 1)
                          for c in [os.path.join(d, *parts[-k:])] if os.path.exists(c)), "")
            hit = ("relocated", found) if found else ("missing", self._by_name(parts[-1] if parts else ""))
        self._cache[key] = hit
        return hit


def t_get_assembly_tree(app, args):
    """Component tree of the active assembly from one flat GetComponents call
    (hierarchy from the 'parent-1/child-1' instance paths -- a recursive
    GetChildren walk timed out on an 844-component assembly)."""
    md = model(app)
    if v(md, "GetType") != 2:
        raise ExtError("Active document is not an assembly (get_assembly_tree needs a .sldasm)")
    t0 = time.time()
    mode = str(args.get("mode", "tree")).lower()
    max_depth = int(args.get("max_depth") or 0)
    max_lines = int(args.get("max_lines") or 150)
    details = bool(args.get("details", False))
    out_path = str(args.get("output_path") or "").strip()

    budget = float(args.get("time_budget_s") or 40)
    asm = T(md, "IAssemblyDoc")
    comps = v(asm, "GetComponents", bool(args.get("top_level_only", False))) or ()
    t_list = time.time() - t0
    rd = _Reader()
    # Every member read is a cross-process Invoke (~4 ms each on SW 2026 with
    # an 844-component assembly), so read Name2 first -- it alone gives level
    # and parent -- and the rest only for components inside max_depth, within
    # a time budget (the MCP client gives up after ~60 s).
    named = [(c, rd(c, "Name2")) for c in comps]
    rows, partial = [], False
    for c, name in named:
        level = name.count("/") + 1
        if max_depth and level > max_depth:
            continue
        if time.time() - t0 > budget:
            partial = True
            break
        path = rd(c, "GetPathName") or ""
        row = {"name": name, "level": level,
               "parent": name.rsplit("/", 1)[0] if "/" in name else "",
               "file": _comp_file_code(path), "path": path,
               "asm": path.lower().endswith(".sldasm"),
               "state": _COMP_STATE.get(rd(c, "GetSuppression2"), "?")}
        if details:
            row["config"] = rd(c, "ReferencedConfiguration")
            row["fixed"] = bool(rd(c, "IsFixed"))
            row["hidden"] = not bool(rd(c, "Visible"))
            row["virtual"] = bool(rd(c, "IsVirtual"))
            row["exclude_bom"] = bool(rd(c, "ExcludeFromBOM"))
        rows.append(row)
    all_levels = [n.count("/") + 1 for _, n in named]
    t_read = time.time() - t0

    by_name = {r["name"]: r for r in rows}
    # Parents before children: a child inherits suppression from any
    # ancestor, and its file is searched next to the parent's real file.
    resolve = _RefResolver(v(md, "GetPathName"))
    for r in sorted(rows, key=lambda r: r["level"]):
        p = by_name.get(r["parent"])
        r["suppressed"] = r["state"] == "suppressed" or bool(p and p["suppressed"])
        r["path_status"], found = resolve(r["path"], (p.get("found_path") or p["path"]) if p else "")
        if r["path_status"] != "ok" and found:
            r["found_path"] = found
    files = {r["file"]: r["asm"] for r in rows}
    active_files = {r["file"] for r in rows if not r["suppressed"]}
    missing = sorted({r["file"] for r in rows if r["path_status"] == "missing"})
    relocated = sorted({r["file"] for r in rows if r["path_status"] == "relocated"} - set(missing))
    elsewhere = sorted({r["file"] for r in rows if r["path_status"] == "missing" and r.get("found_path")})
    n_supp = sum(r["suppressed"] for r in rows)
    depth = max(all_levels, default=0)
    states = {}
    for r in rows:
        states[r["state"]] = states.get(r["state"], 0) + 1

    total = len(named)
    head = (f"{v(md, 'GetTitle')}: {total} instances"
            + (f" ({len(rows)} within depth {max_depth})" if max_depth and len(rows) != total else "")
            + f", {len(files)} unique files "
            f"({sum(1 for a in files.values() if not a)} parts, {sum(1 for a in files.values() if a)} "
            f"assemblies"
            + (f"; {len(files) - len(active_files)} only in suppressed instances" if len(active_files) != len(files) else "")
            + f"), depth {depth} | " + ", ".join(f"{k} {n}" for k, n in sorted(states.items())))
    if n_supp != states.get("suppressed", 0):
        head += f" (+{n_supp - states.get('suppressed', 0)} inside suppressed sub-assemblies)"
    if details:
        head += (f" | fixed {sum(r['fixed'] for r in rows)}, hidden {sum(r['hidden'] for r in rows)}, "
                 f"virtual {sum(r['virtual'] for r in rows)}, excluded from BOM "
                 f"{sum(r['exclude_bom'] for r in rows)}")

    def _names(fs):
        return f"{len(fs)} ({', '.join(fs[:5])}{'...' if len(fs) > 5 else ''})"
    if missing:
        head += f" | ⚠ missing on disk: {_names(missing)}"
        if elsewhere:
            head += (f", same name elsewhere under the assembly folder (SW won't find it itself): "
                     f"{_names(elsewhere)}")
    if relocated:
        head += (f" | stale stored path (author's machine), file found next to the assembly: "
                 f"{_names(relocated)}")
    head += f" | read {t_read:.1f}s (list {t_list:.1f}s)"
    if partial:
        head += (f" | ⚠ time budget {budget:.0f}s reached: {len(rows)} of {total} components read -- "
                 f"use max_depth, top_level_only, or raise time_budget_s")

    lines = []
    if mode == "flat":
        # unique files: active qty (+ suppressed apart), levels, parents -- a
        # BOM skeleton. Suppressed = own state or inside a suppressed
        # sub-assembly; counting them in ×N inflated the BOM (8200-16-10
        # pulley showed ×4 with all 4 suppressed).
        agg: Dict[str, Dict] = {}
        for r in rows:
            a = agg.setdefault(r["file"], {"qty": 0, "supp": 0, "levels": set(), "parents": {},
                                           "asm": r["asm"]})
            a["supp" if r["suppressed"] else "qty"] += 1
            a["levels"].add(r["level"])
            pf = by_name[r["parent"]]["file"] if r["parent"] in by_name else "<top>"
            a["parents"][pf] = a["parents"].get(pf, 0) + 1
        lines.append("(×N = active instances; suppressed ones counted apart, parents count both)")
        for f, a in sorted(agg.items(), key=lambda kv: (not kv[1]["asm"], kv[0].lower())):
            if max_depth and min(a["levels"]) > max_depth:
                continue
            lines.append(f"{'A' if a['asm'] else 'P'} {f} ×{a['qty']}"
                         + (f" (+{a['supp']} suppressed)" if a["supp"] else "")
                         + f" L{','.join(map(str, sorted(a['levels'])))} "
                         f"in {', '.join(f'{p}×{n}' for p, n in a['parents'].items())}")
    else:
        # indented tree; identical sibling files collapsed to "×N"
        children: Dict[str, List[dict]] = {}
        for r in rows:
            children.setdefault(r["parent"], []).append(r)

        def walk(parent, indent):
            groups: Dict[str, List[dict]] = {}
            for r in children.get(parent, []):
                groups.setdefault(r["file"], []).append(r)
            for f, rs in groups.items():
                # show the structure of an active instance if there is one
                # (a suppressed sub-assembly has no children to walk)
                r = next((x for x in rs if not x["suppressed"]), rs[0])
                if max_depth and r["level"] > max_depth:
                    continue
                st: Dict[str, int] = {}
                for x in rs:
                    st[x["state"]] = st.get(x["state"], 0) + 1
                if len(st) == 1:
                    flag = "" if r["state"] == "resolved" else f" [{r['state']}]"
                else:  # mixed group: say how many of the ×N are what
                    flag = " [" + ", ".join(f"{n} {s}" for s, n in sorted(st.items()) if s != "resolved") + "]"
                if details:
                    flag += "".join(f" [{k}]" for k in ("fixed", "hidden", "virtual") if any(x[k] for x in rs))
                lines.append(f"{'  ' * indent}{f}{' ×' + str(len(rs)) if len(rs) > 1 else ''}{flag}")
                if r["asm"]:
                    walk(r["name"], indent + 1)  # siblings of one file share structure

        walk("", 0)

    shown = lines[:max_lines]
    text = head + "\n" + "\n".join(shown)
    if len(lines) > max_lines:
        text += f"\n... {len(lines) - max_lines} more lines (raise max_lines, lower max_depth, or use output_path)"
    if out_path:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as fh:
            json.dump({"assembly": v(md, "GetPathName"), "instances": rows}, fh,
                      ensure_ascii=False, indent=0)
        text += f"\nfull instance list ({len(rows)}) -> {out_path}"
    return _ok(text)


# ============================================================================
# Component properties (custom properties, material, mass) per unique file
# ============================================================================

# Custom property names that hold the drawing / part number and the material
# (compared case-insensitively, surrounding spaces ignored).
_DRAWING_NO_PROPS = ("图号", "partno", "part no", "part number", "partnumber", "number",
                     "零件号", "代号", "обозначение")
_MATERIAL_PROPS = ("material", "材料", "材质", "материал")
_SW_LOADED = (2, 3)  # swComponentResolved / swComponentFullyResolved

# Worklists of recent calls, so a cursor continuation doesn't re-read every
# component (~10 ms per COM call on a loaded assembly). Lost on reload_api,
# then rebuilt and checked against the cursor's fingerprint.
_CP_WORK: Dict[str, dict] = {}


class _Raw:
    """IDispatch::Invoke on raw PyIDispatch with dispids cached per
    (role, member). v() pays GetIDsOfNames on every call and wraps every
    returned object in dynamic.Dispatch (more cross-process typeinfo calls):
    ~190 ms per file vs ~90 ms this way on 8320A-ZZ-01. Roles keep dispids of
    different COM classes apart (a part and an assembly document)."""

    def __init__(self):
        self._ids: Dict[tuple, int] = {}

    def __call__(self, ole, role: str, name: str, *args):
        ole = getattr(ole, "_oleobj_", ole)
        key = (role, name)
        d = self._ids.get(key)
        if d is None:
            d = self._ids[key] = ole.GetIDsOfNames(0, name)
        return ole.Invoke(d, 0, com._GET_OR_CALL, True, *args)

    def put(self, ole, role: str, name: str, value):
        ole = getattr(ole, "_oleobj_", ole)
        key = (role, name)
        d = self._ids.get(key)
        if d is None:
            d = self._ids[key] = ole.GetIDsOfNames(0, name)
        ole.Invoke(d, 0, pythoncom.DISPATCH_PROPERTYPUT, False, value)


def _cp_type_names() -> Dict[int, str]:
    try:
        from .utils.typelib import enum_values
        return {x: n.replace("swCustomInfo", "").lower()
                for n, x in enum_values("swCustomInfoType_e").items()}
    except Exception:
        return {}


def _cp_read(rd: _Raw, x, cfg: str, tnames: Dict[int, str]) -> Dict[str, dict]:
    """All custom properties of one configuration ('' = file level).

    GetAll3 returns names, types, RESOLVED values, a swCustomInfoGetResult_e
    per property (not the text!) and link flags -- the unevaluated expression
    ('"SW-Mass@Part1.SLDPRT"') is not in it. The legacy GetAll returns the
    raw values in the same order, so value + resolved cost two calls, not
    one Get6 per property. Checked live 2026-09-26 on a scratch part."""
    cp = rd(x, "ext", "CustomPropertyManager", cfg)
    if cp is None:
        return {}
    o3 = [com.out_variant() for _ in range(5)]
    n = rd(cp, "cpm", "GetAll3", *o3)
    if not n:
        return {}
    names, types, resolved, _res, links = (tuple(o.value or ()) for o in o3)
    o1 = [com.out_variant() for _ in range(3)]
    rd(cp, "cpm", "GetAll", *o1)
    raw = dict(zip(tuple(o1[0].value or ()), tuple(o1[2].value or ())))
    out = {}
    for i, name in enumerate(names):
        p = {"value": raw.get(name, resolved[i] if i < len(resolved) else ""),
             "resolved": resolved[i] if i < len(resolved) else "",
             "type": tnames.get(types[i], types[i]) if i < len(types) else "?"}
        if i < len(links) and links[i]:
            p["linked"] = True
        out[name] = p
    return out


def _cp_pick(props_list, keys) -> tuple:
    """(name, resolved value) of the first property matching `keys`, searching
    the dicts in order (config-specific before file level)."""
    for props in props_list:
        for name, p in props.items():
            if name.strip().lower() in keys and str(p.get("resolved", "")).strip():
                return name, str(p["resolved"]).strip()
    return "", ""


def _cp_norm(s: str) -> str:
    return "".join(str(s).split()).casefold()


def _cp_worklist(app, md, include_suppressed: bool) -> dict:
    """Unique component files of the active assembly with active/suppressed
    quantities, parents and their instances. One GetComponents call; per
    instance Name2 / GetPathName / GetSuppression2 (as get_assembly_tree)."""
    asm = T(md, "IAssemblyDoc")
    comps = v(asm, "GetComponents", False) or ()
    rd = _Reader()
    inst = []
    for c in comps:
        inst.append({"c": c, "name": rd(c, "Name2"), "path": rd(c, "GetPathName") or "",
                     "state": rd(c, "GetSuppression2")})
    by_name = {i["name"]: i for i in inst}
    # Group by the file SW would really use: a suppressed instance keeps the
    # author's stale path (982S-02 on 8320A-ZZ-01: one active instance at the
    # real path, one suppressed at C:\Users\Administrator\... -- same file,
    # which grouping by stored path counted as two).
    resolve = _RefResolver(v(md, "GetPathName"))
    for i in sorted(inst, key=lambda i: i["name"].count("/")):
        p = by_name.get(i["name"].rsplit("/", 1)[0]) if "/" in i["name"] else None
        i["supp"] = i["state"] == 0 or bool(p and p["supp"])
        i["parent"] = _comp_file_code(p["path"]) if p else "<top>"
        status, found = resolve(i["path"], p["real"] if p else "")
        i["real"] = found if status == "relocated" else i["path"]
    files: Dict[str, dict] = {}
    for i in inst:
        f = files.setdefault(i["real"].lower(), {
            "file": _comp_file_code(i["real"]), "path": i["real"],
            "type": "assembly" if i["real"].lower().endswith(".sldasm") else "part",
            "qty_active": 0, "qty_suppressed": 0, "parents": {}, "inst": []})
        f["qty_suppressed" if i["supp"] else "qty_active"] += 1
        f["parents"][i["parent"]] = f["parents"].get(i["parent"], 0) + 1
        f["inst"].append(i)
        if i["path"] != i["real"]:
            f.setdefault("stored_paths", [])
            if i["path"] not in f["stored_paths"]:
                f["stored_paths"].append(i["path"])
    keys = sorted(files)
    work = [files[k] for k in keys if include_suppressed or files[k]["qty_active"]]
    return {"asm": v(md, "GetPathName"), "work": work,
            "skipped": len(keys) - len(work), "n_inst": len(inst)}


def _cp_fingerprint(wl: dict, settings: dict) -> str:
    h = hashlib.sha1(json.dumps([wl["asm"].lower(), settings,
                                 [(w["path"].lower(), w["qty_active"], w["qty_suppressed"])
                                  for w in wl["work"]]], ensure_ascii=False).encode("utf-8"))
    return h.hexdigest()[:10]


def _cp_file(rd: _Raw, w: dict, idx: int, settings: dict, tnames, resolve) -> dict:
    """One JSONL row. Only reads: GetModelDoc2 of an already loaded
    instance; a lightweight / suppressed file is reported, never resolved."""
    row = {"kind": "file", "i": idx, "file": w["file"], "path": w["path"], "type": w["type"],
           "qty_active": w["qty_active"], "qty_suppressed": w["qty_suppressed"],
           "parents": w["parents"]}
    if w.get("stored_paths"):
        row["stored_paths"] = w["stored_paths"]
    live = [i for i in w["inst"] if not i["supp"] and i["state"] in _SW_LOADED]
    if not live and settings["include_suppressed"]:
        live = [i for i in w["inst"] if i["state"] in _SW_LOADED]
    doc = rd(live[0]["c"], "comp", "GetModelDoc2") if live else None
    if doc is None:
        states = sorted({_COMP_STATE.get(i["state"], "?") for i in w["inst"]})
        status, _found = resolve(w["inst"][0]["path"])
        row.update(status="not_loaded",
                   reason=("all instances suppressed" if all(i["supp"] for i in w["inst"])
                           else "no resolved instance") + f" (states: {', '.join(states)})",
                   path_status=status)
        if _found and status != "ok":
            row["found_path"] = _found
        return row

    role = "asm" if w["type"] == "assembly" else "part"
    x = rd(doc, role, "Extension")
    cfg_mode = settings["config"]
    if cfg_mode == "active":
        seen = {}
        for i in live:
            c = rd(i["c"], "comp", "ReferencedConfiguration") or ""
            seen[c] = seen.get(c, 0) + 1
        configs = sorted(seen, key=lambda c: -seen[c])   # most used first
    elif cfg_mode == "all":
        configs = list(rd(doc, role, "GetConfigurationNames") or ())
    elif cfg_mode == "none":
        configs = []
    else:
        configs = [cfg_mode]
    row["configs"] = configs
    file_props = _cp_read(rd, x, "", tnames)
    cfg_props = {c: _cp_read(rd, x, c, tnames) for c in configs}
    row["props"] = {"file": file_props, "config": cfg_props}
    primary = configs[0] if configs else ""
    search = [cfg_props.get(primary, {}), file_props]
    row["drawing_no"] = _cp_pick(search, _DRAWING_NO_PROPS)[1]
    pname, pval = _cp_pick(search, _MATERIAL_PROPS)

    if w["type"] == "part":
        mats = {}
        for c in (configs or [""]):
            mo = com.out_str()
            mats[c] = (rd(doc, role, "GetMaterialPropertyName2", c, mo) or "", mo.value or "")
        sw_mat, db = mats[primary]
        row["material"] = {"sw": sw_mat, "db": db, "config": primary,
                           "property": pval, "property_name": pname}
        other = {c: m[0] for c, m in mats.items() if c != primary and m[0] != sw_mat}
        if other:
            row["material"]["other_configs"] = other
        row["no_material"] = not sw_mat
        row["material_mismatch"] = bool(sw_mat and pval and _cp_norm(sw_mat) != _cp_norm(pval))
        if settings["include_mass"]:
            mp = rd(x, "ext", "CreateMassProperty")
            if mp is not None:
                try:
                    rd.put(mp, "mass", "UseSystemUnits", True)
                except pythoncom.com_error:
                    pass
                row["mass_kg"] = float(rd(mp, "mass", "Mass"))
                row["volume_mm3"] = float(rd(mp, "mass", "Volume")) * 1e9
    elif pval:
        row["material"] = {"property": pval, "property_name": pname}
    row["status"] = "ok"
    return row


def _cp_summary(rows: List[dict], skipped: int, include_mass: bool) -> dict:
    ok = [r for r in rows if r.get("status") == "ok"]
    parts = [r for r in ok if r["type"] == "part"]
    nums: Dict[str, List[str]] = {}
    for r in ok:
        n = (r.get("drawing_no") or "").strip()
        if n:
            nums.setdefault(n, []).append(r["file"])
    dups = {n: fs for n, fs in sorted(nums.items()) if len(fs) > 1}
    s = {"kind": "summary", "files": len(rows),
         "parts": sum(r["type"] == "part" for r in rows),
         "assemblies": sum(r["type"] == "assembly" for r in rows),
         "ok": len(ok),
         "not_loaded": sum(r.get("status") == "not_loaded" for r in rows),
         "errors": sum(r.get("status") == "error" for r in rows),
         "no_material": sum(bool(r.get("no_material")) for r in parts),
         "material_mismatch": sum(bool(r.get("material_mismatch")) for r in parts),
         "mismatch_files": [f"{r['file']}: SW {r['material']['sw']!r} vs "
                            f"{r['material']['property_name']} {r['material']['property']!r}"
                            for r in parts if r.get("material_mismatch")],
         "drawing_no_duplicates": dups,
         "skipped_suppressed_only": skipped}
    if include_mass:
        s["mass_active_kg"] = round(sum(r.get("mass_kg", 0.0) * r["qty_active"] for r in parts), 4)
    return s


def t_get_component_properties(app, args):
    """Custom properties, material and (optionally) mass of every unique
    component file of the active assembly, streamed to JSONL within a time
    budget and resumable by cursor. Read-only: documents come from
    Component2.GetModelDoc2 of instances already loaded in the assembly."""
    t0 = time.time()
    md = model(app)
    if v(md, "GetType") != 2:
        raise ExtError("Active document is not an assembly (get_component_properties needs a .sldasm)")
    out_path = str(args.get("output_path") or "").strip()
    if not out_path:
        raise ExtError("output_path (.jsonl) is required")
    budget = float(args.get("budget_s") or 45)
    cursor = str(args.get("cursor") or "").strip()
    cfg = str(args.get("config") if args.get("config") is not None else "active").strip() or "none"
    settings = {"config": cfg,
                "include_mass": bool(args.get("include_mass", False)),
                "include_suppressed": bool(args.get("include_suppressed", False))}

    wl = None
    start = 0
    if cursor:
        try:
            start_s, _total, fp = cursor.split("/")
            start = int(start_s)
        except ValueError:
            raise ExtError(f"Bad cursor {cursor!r} (expected 'next/total/fingerprint' from the previous call)")
        wl = _CP_WORK.get(fp)
    if wl is None:
        wl = _cp_worklist(app, md, settings["include_suppressed"])
    fp_now = _cp_fingerprint(wl, settings)
    if cursor and fp_now != fp:
        raise ExtError("cursor does not match the assembly / settings / component list any more "
                       "(different config/include_* arguments, another active assembly, or the "
                       "structure changed) -- start over without cursor")
    _CP_WORK.clear()
    _CP_WORK[fp_now] = wl
    work = wl["work"]
    t_list = time.time() - t0

    header = {"kind": "header", "assembly": wl["asm"], **settings, "files": len(work),
              "instances": wl["n_inst"], "skipped_suppressed_only": wl["skipped"],
              "fingerprint": fp_now}
    done_keys, rows_before = set(), []
    if cursor:
        if not os.path.exists(out_path):
            raise ExtError(f"cursor given but {out_path} does not exist -- start over without cursor")
        with open(out_path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                r = json.loads(line)
                if r.get("kind") == "header" and r.get("fingerprint") != fp_now:
                    raise ExtError(f"{out_path} belongs to another run (fingerprint "
                                   f"{r.get('fingerprint')} != {fp_now}) -- start over without cursor")
                if r.get("kind") == "file":
                    done_keys.add(r["path"].lower())
                    rows_before.append(r)
                elif r.get("kind") == "summary":
                    raise ExtError(f"{out_path} is already complete (has a summary line)")
        mode = "a"
    else:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
        mode = "w"

    rd = _Raw()
    tnames = _cp_type_names()
    resolve = _RefResolver(wl["asm"])
    new_rows, i = [], start
    with open(out_path, mode, encoding="utf-8", newline="\n") as fh:
        if mode == "w":
            fh.write(json.dumps(header, ensure_ascii=False) + "\n")
        while i < len(work):
            # at least one file per call, so a tiny budget still progresses
            if new_rows and time.time() - t0 > budget:
                break
            w = work[i]
            if w["path"].lower() not in done_keys:
                try:
                    row = _cp_file(rd, w, i, settings, tnames, resolve)
                except Exception as e:
                    row = {"kind": "file", "i": i, "file": w["file"], "path": w["path"],
                           "type": w["type"], "qty_active": w["qty_active"],
                           "qty_suppressed": w["qty_suppressed"], "parents": w["parents"],
                           "status": "error", "error": f"{type(e).__name__}: {e}"}
                fh.write(json.dumps(row, ensure_ascii=False, default=str) + "\n")
                fh.flush()
                new_rows.append(row)
                done_keys.add(w["path"].lower())
            i += 1
        done = i >= len(work)
        summary = None
        if done:
            summary = _cp_summary(rows_before + new_rows, wl["skipped"], settings["include_mass"])
            fh.write(json.dumps(summary, ensure_ascii=False) + "\n")

    el = time.time() - t0
    if not done:
        nxt = f"{i}/{len(work)}/{fp_now}"
        return _ok(f"done:false | {len(new_rows)} file(s) this call, {i}/{len(work)} total, "
                   f"{len(work) - i} remaining | {el:.1f}s (component list {t_list:.1f}s) -> {out_path}\n"
                   f"call again with cursor={nxt!r} (same output_path and arguments)",
                   {"done": False, "cursor": nxt, "processed": i, "remaining": len(work) - i})
    s = summary
    dups = s["drawing_no_duplicates"]
    msg = (f"done:true | {s['files']} files ({s['parts']} parts, {s['assemblies']} assemblies), "
           f"ok {s['ok']}, not_loaded {s['not_loaded']}, errors {s['errors']}"
           + (f", skipped {s['skipped_suppressed_only']} suppressed-only" if s["skipped_suppressed_only"] else "")
           + f" | no SW material {s['no_material']} parts, material mismatch {s['material_mismatch']}"
           + f" | drawing no. duplicates {len(dups)}"
           + (f" ({'; '.join(f'{n}: {len(fs)} files' for n, fs in list(dups.items())[:5])})" if dups else "")
           + (f" | mass of active parts {s['mass_active_kg']} kg" if "mass_active_kg" in s else "")
           + f" | this call {len(new_rows)} file(s), {el:.1f}s -> {out_path}")
    return _ok(msg, {"done": True, "summary": {k: v_ for k, v_ in s.items()
                                                if k not in ("kind", "mismatch_files")}})


HANDLERS = {
    "get_assembly_tree": t_get_assembly_tree,
    "get_component_properties": t_get_component_properties,
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
    "create_reference_plane": t_create_reference_plane,
    "create_reference_axis": t_create_reference_axis,
    "mass_properties": t_mass_properties,
    "get_rebuild_errors": t_get_rebuild_errors,
    "edit_feature": t_edit_feature,
    "sketch_entities": t_sketch_entities,
    "add_sketch_relation": t_add_sketch_relation,
    "add_sketch_dimension": t_add_sketch_dimension,
    "create_drawing": t_create_drawing,
    "add_standard_views": t_add_standard_views,
    "add_drawing_view": t_add_drawing_view,
    "add_section_view": t_add_section_view,
    "add_detail_view": t_add_detail_view,
    "add_broken_out_section": t_add_broken_out_section,
    "insert_model_dimensions": t_insert_model_dimensions,
    "add_note": t_add_note,
    "export_pdf": t_export_pdf,
    "move_drawing_view": t_move_drawing_view,
    "delete_drawing_view": t_delete_drawing_view,
    "add_drawing_dimension": t_add_drawing_dimension,
    "add_gtol": t_add_gtol,
    "add_datum": t_add_datum,
    "add_surface_finish": t_add_surface_finish,
    "delete_annotation": t_delete_annotation,
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

_EDGE_DESC = ("Edge of the part shown in the view: an index from list_edges on that part ('12'), "
              "or a model point 'x,y,z' in part mm -- the nearest edge to it is used")

TOOL_SCHEMAS = [
    ("get_assembly_tree",
     "Component tree of the active assembly: header with instance / unique-file / part / sub-assembly "
     "counts, depth, resolved-lightweight-suppressed states, files really missing on disk vs stale stored "
     "paths whose file sits next to the assembly, then either an indented tree (identical siblings "
     "collapsed to xN, quantities per parent) or a flat list of unique files with active quantity, "
     "suppressed quantity apart, levels and parents (a BOM skeleton). output_path writes every instance "
     "(with suppressed / path_status / found_path) to JSON for further processing.",
     _obj({"mode": {"type": "string", "enum": ["tree", "flat"], "default": "tree"},
           "max_depth": {"type": "integer", "default": 0, "description": "0 = all levels"},
           "max_lines": {"type": "integer", "default": 150},
           "details": {"type": "boolean", "default": False,
                       "description": "Also read config, fixed/float, hidden, virtual, excluded-from-BOM"},
           "top_level_only": {"type": "boolean", "default": False},
           "time_budget_s": {"type": "number", "default": 40,
                             "description": "Stop reading and report partial after this many seconds "
                                            "(~4 ms per component property on SW 2026)"},
           "output_path": {"type": "string", "description": "Optional .json path for the full instance list"}})),
    ("get_component_properties",
     "Custom properties (file-level and configuration, value + resolved), SW material vs the "
     "Material/材料 property, optional mass -- per UNIQUE component file of the active assembly, "
     "one JSON line per file in output_path. Read-only: uses documents already loaded in the "
     "assembly; lightweight/suppressed files get status not_loaded. Stops at budget_s and returns "
     "done:false + cursor; call again with the same output_path/arguments and that cursor to continue "
     "(appends, no duplicates). done:true returns a summary: files, not_loaded, no material, material "
     "mismatches, duplicate drawing numbers (图号/PartNo/Number).",
     _obj({"output_path": {"type": "string", "description": "JSONL file (overwritten on a call without cursor)"},
           "budget_s": {"type": "number", "default": 45,
                        "description": "Stop after this many seconds (at least one file per call)"},
           "cursor": {"type": "string", "description": "From a previous done:false answer"},
           "include_mass": {"type": "boolean", "default": False,
                            "description": "Also mass/volume per part (IMassProperty)"},
           "include_suppressed": {"type": "boolean", "default": False,
                                  "description": "Also list files that occur only in suppressed "
                                                 "instances (as not_loaded)"},
           "config": {"type": "string", "default": "active",
                      "description": "Configuration properties to read: 'active' = the configuration(s) "
                                     "the assembly references, 'all', 'none' (file-level only) or a name"}},
          ["output_path"])),
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
    ("create_reference_plane",
     "Create a reference plane. kind=distance (default) offsets a parallel plane by value mm -- "
     "this is the working way around the sketch-on-cylindrical-face dead end. Other kinds: angle "
     "(value in deg, needs an axis/edge as second_reference), coincident, parallel, perpendicular, "
     "tangent, midplane. Reference: 'front'/'top'/'right', a feature name, or 'face:N'/'edge:N' "
     "from list_faces/list_edges.",
     _obj({"reference": {"type": "string", "description": "'front'/'top'/'right', feature name, or 'face:N'/'edge:N'"},
           "kind": {"type": "string",
                    "enum": ["distance", "angle", "coincident", "parallel",
                             "perpendicular", "tangent", "midplane"],
                    "default": "distance"},
           "value": {"type": "number", "description": "mm for distance, deg for angle"},
           "second_reference": {"type": "string", "description": "Second reference, when the kind needs one"},
           "reverse": {"type": "boolean", "default": False, "description": "Flip to the other side"},
           "name": {"type": "string", "description": "Rename the new plane"}},
          ["reference"])),
    ("create_reference_axis",
     "Create a reference axis from two intersecting planes, two points, one cylindrical face, or a "
     "line. Two PARALLEL planes do not define an axis and are rejected. References are comma-"
     "separated: 'front,right' or 'face:2'.",
     _obj({"references": {"type": "string", "description": "e.g. 'front,right' or 'face:2'"},
           "name": {"type": "string", "description": "Rename the new axis"},
           "hide": {"type": "boolean", "default": False, "description": "Hide it after creation"}},
          ["references"])),
    ("mass_properties",
     "Mass, volume, surface area, density, centre of mass and principal moments of the active part. "
     "Read in SI and reported in g/kg, mm3, mm2, g/cm3 and mm regardless of the document's own units.",
     _obj({})),
    ("get_rebuild_errors",
     "What is broken in the active document: features with rebuild errors or warnings, and sketches "
     "that are not fully defined. Run after a failed build, or before trusting a model.",
     _obj({})),
    ("edit_feature",
     "Change dimensions of an existing feature in place, by their SHORT names (D1, D2 -- as shown by "
     "get_parameters), without rebuilding the model from scratch. Several dimensions cost one "
     "rebuild. Reports the volume/topology delta and any rebuild errors afterwards.",
     _obj({"feature": {"type": "string", "description": "Feature name, e.g. 'Boss-Extrude1'"},
           "dimensions": {"type": "string", "description": "e.g. 'D1=25, D2=10' (mm, or deg for angles)"}},
          ["feature", "dimensions"])),
    ("reload_api",
     "Hot-reload the server's Python code (automation/*, ext.py) from disk after editing it. "
     "Keeps the SolidWorks connection. New/changed tool SCHEMAS still need an MCP restart.",
     _obj({})),
    ("create_drawing",
     "Create a new drawing document from SolidWorks' default drawing template (disk search, "
     "then the running app's own configured default). Becomes the active document.",
     _obj({})),
    ("add_standard_views",
     "Lay out the standard 3 views (+ isometric) of a saved part/assembly on the active drawing "
     "sheet in one call, per first- or third-angle projection. model_path must be a real file on "
     "disk (does not need to be open). Active document must be a drawing (create_drawing first).",
     _obj({"model_path": {"type": "string", "description": "Full path to a saved .sldprt/.sldasm"},
           "projection": {"type": "string", "enum": ["first", "third"], "default": "third",
                           "description": "first = ISO/Europe, third = ANSI/US"}},
          ["model_path"])),
    ("add_drawing_view",
     "Add a single named view of a saved part/assembly to the active drawing sheet. "
     "view: Front/Back/Left/Right/Top/Bottom/Isometric/Dimetric/Trimetric/Current. "
     "x,y = sheet position in mm from the sheet origin (bottom-left). Active document must be "
     "a drawing (create_drawing first).",
     _obj({"model_path": {"type": "string", "description": "Full path to a saved .sldprt/.sldasm"},
           "view": {"type": "string", "default": "current"},
           "x": {"type": "number", "default": 100, "description": "mm"},
           "y": {"type": "number", "default": 100, "description": "mm"},
           "scale": {"type": "number", "description": "e.g. 0.5 for 1:2; omit to keep sheet scale"}},
          ["model_path"])),
    ("add_section_view",
     "Cut a section view through an existing drawing view along a straight line. source_view must "
     "already be on the sheet (add_drawing_view/add_standard_views). x1,y1,x2,y2 = cutting line "
     "endpoints, x,y = where the new section view lands -- both in sheet mm, same coordinate space "
     "as add_drawing_view's x,y (NOT relative to source_view; the tool converts to the view's own "
     "sketch space). Extend the line past the view outline. depth=0 cuts fully through. "
     "Remove with delete_drawing_view (also clears the section line it leaves in the parent).",
     _obj({"source_view": {"type": "string", "description": "e.g. 'Drawing View1'"},
           "x1": {"type": "number", "description": "mm"}, "y1": {"type": "number", "description": "mm"},
           "x2": {"type": "number", "description": "mm"}, "y2": {"type": "number", "description": "mm"},
           "x": {"type": "number", "description": "mm, new view placement"},
           "y": {"type": "number", "description": "mm, new view placement"},
           "label": {"type": "string", "description": "e.g. 'A'; omit for auto next letter"},
           "depth": {"type": "number", "default": 0, "description": "mm, 0 = full depth"},
           "options": {"type": "string",
                       "description": "Comma-separated: not_aligned, offset, change_direction, "
                                      "scale_with_model, partial, display_surface_cut, "
                                      "exclude_fasteners, cut_surface_bodies"}},
          ["source_view", "x1", "y1", "x2", "y2", "x", "y"])),
    ("add_broken_out_section",
     "Broken-out section on an existing drawing view. Default = the way to section a body of "
     "revolution: ONE side view + a contour around the whole view + depth = half the part's "
     "extent along the view direction (= outer radius), giving a full hatched section with no end "
     "view, section line or label. points = custom closed contour, sheet mm. depth = model mm from "
     "the nearest point of the part.",
     _obj({"view": {"type": "string", "description": "View display or feature name, e.g. 'Drawing View1'"},
           "depth": {"type": "number", "description": "mm (model); omit = half the part depth along the view"},
           "points": {"type": "string", "description": "JSON [[x,y],...] closed contour in sheet mm; omit = view outline + margin"},
           "margin": {"type": "number", "default": 2, "description": "mm around the view outline for the default contour"}},
          ["view"])),
    ("add_detail_view",
     "Magnify a circular area of an existing drawing view into its own detail view. source_view "
     "must already be on the sheet. x,y,radius = the circle to magnify (sheet mm, on source_view); "
     "place_x,place_y = where the new detail view lands. scale 2.0 = '2:1'.",
     _obj({"source_view": {"type": "string", "description": "e.g. 'Drawing View1'"},
           "x": {"type": "number", "description": "mm, circle centre"},
           "y": {"type": "number", "description": "mm, circle centre"},
           "radius": {"type": "number", "description": "mm"},
           "place_x": {"type": "number", "description": "mm, new view placement"},
           "place_y": {"type": "number", "description": "mm, new view placement"},
           "scale": {"type": "number", "default": 2.0},
           "label": {"type": "string", "description": "e.g. 'A'; omit for auto next letter"},
           "style": {"type": "string", "enum": ["standard", "broken", "leader", "noleader", "connected"],
                     "default": "standard"},
           "full_outline": {"type": "boolean", "default": True},
           "jagged_outline": {"type": "boolean", "default": False},
           "no_outline": {"type": "boolean", "default": False},
           "shape_intensity": {"type": "integer", "default": 1}},
          ["source_view", "x", "y", "radius", "place_x", "place_y"])),
    ("insert_model_dimensions",
     "Pull driving dimensions from the model into the given drawing view(s). Only dimensions "
     "native to geometry actually visible from a view land in it -- e.g. a hole's diameter only "
     "appears in a view where its round profile is shown edge-on, not just a centre mark. "
     "views omitted = every view on the active sheet.",
     _obj({"views": {"type": "string", "description": "Comma-separated view names; omit for all"},
           "option": {"type": "integer", "default": 0}})),
    ("add_note",
     "Add a text note to the active drawing sheet at a fixed position (not attached to any "
     "geometry). x,y = sheet mm from the sheet origin; stay within the sheet's own paper size or "
     "the note renders on screen but gets cropped out of a PDF export.",
     _obj({"text": {"type": "string"},
           "x": {"type": "number", "default": 100, "description": "mm"},
           "y": {"type": "number", "default": 100, "description": "mm"},
           "height": {"type": "number", "description": "mm text height; omit for the sheet default"}},
          ["text"])),
    ("move_drawing_view",
     "Move a drawing view so its centre sits at x,y (sheet mm). Aligned (projected/section) views "
     "can only slide along their alignment axis.",
     _obj({"view": {"type": "string", "description": "Display name ('Section View A-A') or feature name"},
           "x": {"type": "number", "description": "mm"}, "y": {"type": "number", "description": "mm"}},
          ["view", "x", "y"])),
    ("delete_drawing_view",
     "Delete a drawing view. For a section view this also removes the section line it leaves "
     "behind in the parent view and the cutting line that section line hands back to the parent's sketch.",
     _obj({"view": {"type": "string", "description": "Display name or feature name"}}, ["view"])),
    ("add_drawing_dimension",
     "Driven dimension in a drawing view from model edges: one circular edge -> diameter, two edges -> "
     "distance. x,y = text position in sheet mm (also decides horizontal/vertical/aligned; omit to put "
     "it 10 mm above the edges). Optional prefix (e.g. '<MOD-DIAM>'), tol_upper/tol_lower in mm "
     "(bilateral; +a/-a gives symmetric), fit ('H7', 'h7' or 'H7/g6'). Parentheses off unless parenthesis=true.",
     _obj({"view": {"type": "string"},
           "edge1": {"type": "string", "description": _EDGE_DESC},
           "edge2": {"type": "string", "description": "Second edge, same format; omit for a diameter/radius"},
           "x": {"type": "number", "description": "mm, text position"},
           "y": {"type": "number", "description": "mm, text position"},
           "prefix": {"type": "string"}, "suffix": {"type": "string"},
           "tol_upper": {"type": "number", "description": "mm, e.g. 0.2"},
           "tol_lower": {"type": "number", "description": "mm, e.g. -0.1"},
           "fit": {"type": "string", "description": "'H7' (hole), 'h7' (shaft) or 'H7/g6'"},
           "parenthesis": {"type": "boolean", "default": False}},
          ["view", "edge1"])),
    ("add_gtol",
     "Geometric tolerance frame in a drawing view, leadered to a model edge (or free without edge). "
     "symbol: runout, total_runout, cylindricity, parallelism, perpendicularity, flatness, position, "
     "concentricity, circularity, straightness, symmetry, angularity, profile_line, profile_surface -- "
     "or a gtol.sym name ('SRUN', '<IGTOL-SRUN>'). Decimal comma as the drawing expects ('0,04').",
     _obj({"view": {"type": "string"},
           "edge": {"type": "string", "description": _EDGE_DESC + "; omit for a free frame"},
           "symbol": {"type": "string"},
           "tolerance": {"type": "string", "description": "e.g. '0,04'"},
           "datum": {"type": "string", "description": "Up to 3, comma-separated: 'A' or 'A,B'"},
           "diameter": {"type": "boolean", "default": False, "description": "Ø before the tolerance"},
           "x": {"type": "number", "description": "mm, frame position"},
           "y": {"type": "number", "description": "mm, frame position"},
           "library": {"type": "string", "default": "IGTOL", "description": "IGTOL (ISO) or GGTOL (GOST)"},
           "convert": {"type": "boolean", "default": True,
                       "description": "Upgrade to the SW2022 gtol format after filling it in"}},
          ["view", "symbol", "tolerance"])),
    ("add_datum",
     "Datum feature symbol attached to a model edge in a drawing view.",
     _obj({"view": {"type": "string"},
           "edge": {"type": "string", "description": _EDGE_DESC},
           "label": {"type": "string", "description": "e.g. 'A'; omit for the next free letter"},
           "x": {"type": "number", "description": "mm, symbol position"},
           "y": {"type": "number", "description": "mm, symbol position"}},
          ["view", "edge"])),
    ("add_surface_finish",
     "Surface roughness symbol in a drawing view, leadered to a model edge or free (then x,y required). "
     "value e.g. 'Ra 3,2'; the text slot that renders depends on the document's surface-finish "
     "standard and is chosen automatically.",
     _obj({"view": {"type": "string"},
           "edge": {"type": "string", "description": _EDGE_DESC + "; omit for a free symbol"},
           "value": {"type": "string", "description": "e.g. 'Ra 3,2'"},
           "symbol": {"type": "string", "enum": ["machined", "basic", "no_machining"], "default": "machined"},
           "x": {"type": "number", "description": "mm"}, "y": {"type": "number", "description": "mm"}},
          ["view", "value"])),
    ("delete_annotation",
     "Delete annotations from a drawing view by name (as returned by the add_* tools: 'RD1', "
     "'DetailItem12'). kind: dimension, datum, gtol, surface_finish, note.",
     _obj({"view": {"type": "string"},
           "names": {"type": "string", "description": "Comma-separated"},
           "kind": {"type": "string", "enum": ["dimension", "datum", "gtol", "surface_finish", "note"],
                    "default": "dimension"}},
          ["view", "names"])),
    ("export_pdf",
     "Export the active drawing to PDF. path must end in .pdf and its directory must already exist.",
     _obj({"path": {"type": "string"}}, ["path"])),
]
