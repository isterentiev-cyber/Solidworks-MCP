"""
SolidWorks Feature Operations
-----------------------------
Create 3D features: extrude, cut, fillet, chamfer, etc.

Version: 4.0.0 (Fixed for SolidWorks 2025 - v33)
Author: Samsaam Ali Baig

Fixes v4.0.0:
- FeatureExtrusion2 now uses correct 23 parameters for SW 2025
- Proper sketch close + select before extrude (fixes multi-profile sketches)
- Property vs method fixes (FirstFeature, GetNextFeature, GetTypeName2)
- Added _find_last_sketch helper for reliable sketch selection
- Added _get_sketch_info helper for better error diagnostics
- Better error messages with sketch profile count
"""

import logging
import traceback
from typing import Optional, Dict

import win32com.client
import pythoncom

from ..constants import SwErrors, SwEndConditions
from ..utils.com_helpers import v

logger = logging.getLogger(__name__)


class FeatureOperations:
    """
    Mixin class for feature operations

    Requires parent class to have:
    - get_active_doc(): Document access method
    - _result(): Result factory method
    - _units: UnitConverter instance
    """

    # ========================================================================
    # Helper Methods
    # ========================================================================

    def _find_last_sketch(self, doc) -> Optional[str]:
        """Name of the last sketch in the feature tree (walk via ext)."""
        from .. import ext
        last = None
        try:
            for f in ext.user_features(ext.T(doc, "IModelDoc2")):
                if ext.v(f, "GetTypeName2") == "ProfileFeature":
                    last = f.Name
        except Exception as e:
            logger.debug(f"_find_last_sketch error: {e}")
        return last

    def _get_sketch_info(self, doc) -> Dict:
        """Sketch diagnostics for error messages."""
        from .. import ext
        info = {"sketch_count": 0, "sketch_names": [], "has_active_sketch": False,
                "feature_count": 0}
        try:
            md = ext.T(doc, "IModelDoc2")
            info["has_active_sketch"] = md.SketchManager.ActiveSketch is not None
            for f in ext.user_features(md):
                info["feature_count"] += 1
                if ext.v(f, "GetTypeName2") == "ProfileFeature":
                    info["sketch_count"] += 1
                    info["sketch_names"].append(f.Name)
        except Exception as e:
            logger.debug(f"_get_sketch_info error: {e}")
        return info

    def _close_and_select_sketch(self, doc) -> tuple:
        """
        Close active sketch if open, find and select the last sketch.

        Returns:
            Tuple of (success: bool, sketch_name: str, error_msg: str)
        """
        try:
            # Step 1: Close active sketch if one is open
            try:
                active_sketch = doc.SketchManager.ActiveSketch
                if active_sketch is not None:
                    doc.SketchManager.InsertSketch(True)
                    logger.debug("Closed active sketch")
            except:
                # Try closing anyway
                try:
                    doc.InsertSketch2(True)
                except:
                    pass

            # Step 2: Clear selection
            doc.ClearSelection2(True)

            # Step 3: Find the last sketch
            sketch_name = self._find_last_sketch(doc)
            if not sketch_name:
                return False, "", "No sketch found in feature tree"

            # Step 4: Refuse a sketch that already drives a feature. Upstream
            # silently re-used it when the new sketch failed to open, and
            # produced a zero-volume Boss-Extrude on top of the old one.
            from .. import ext
            feat = ext.find_feature(ext.T(doc, "IModelDoc2"), sketch_name)
            children = v(feat, "GetChildren") or ()
            if children:
                used_by = [ext.T(c, "IFeature").Name for c in children]
                return False, sketch_name, (
                    f"Last sketch '{sketch_name}' is already used by {used_by} -- "
                    f"no new sketch to consume (did create_sketch fail?)")

            # Step 5: Select the sketch (feature-level select, no SelectByID2
            # callout quirks)
            if not feat.Select2(False, 0):
                return False, sketch_name, f"Could not select sketch '{sketch_name}'"

            return True, sketch_name, ""

        except Exception as e:
            return False, "", f"Error in sketch selection: {e}"

    # ========================================================================
    # Extrude
    # ========================================================================

    def extrude_sketch(self, depth: float = 10, both_directions: bool = False,
                       unit: str = None, reverse: bool = False) -> Dict:
        """
        Extrude the active sketch (Boss-Extrude)
        FIXED v4.0: Properly closes sketch, selects it, uses 23-param FeatureExtrusion2

        Args:
            depth: Extrusion depth
            both_directions: Extrude in both directions (mid-plane)
            unit: Unit for depth

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            # Convert depth to meters
            depth_m = self._units.to_meters(depth, unit)
            unit_str = unit or self._units.default_unit.value

            # Step 1: Close sketch and select it
            success, sketch_name, error_msg = self._close_and_select_sketch(doc)
            if not success:
                sketch_info = self._get_sketch_info(doc)
                return self._result(False,
                    f"Extrusion failed: {error_msg}. "
                    f"Sketches found: {sketch_info['sketch_count']} {sketch_info['sketch_names']}. "
                    f"Active sketch: {sketch_info['has_active_sketch']}. "
                    f"Try: create_sketch → draw geometry → extrude_sketch",
                    SwErrors.swFeatureError,
                    {"diagnostics": sketch_info})

            # Step 2: Determine end condition
            end_cond = 6 if both_directions else 0  # 6=MidPlane, 0=Blind

            # Step 3: Try extrusion methods
            feat = None
            method_used = ""

            # Method 1: FeatureExtrusion2 with 23 params (SW 2025 / v33)
            try:
                feat = doc.FeatureManager.FeatureExtrusion2(
                    True,           # Sd - single direction
                    False,          # Flip
                    reverse,        # Dir - True = opposite to sketch normal
                    end_cond,       # T1 - end condition (0=Blind, 6=MidPlane)
                    0,              # T2 - end condition 2
                    depth_m,        # D1 - depth
                    depth_m,        # D2 - depth 2
                    False,          # Dchk1 - draft on/off
                    False,          # Dchk2 - draft on/off 2
                    False,          # Ddir1 - draft outward
                    False,          # Ddir2 - draft outward 2
                    0.0,            # Dang1 - draft angle (radians)
                    0.0,            # Dang2 - draft angle 2
                    False,          # OffsetReverse1
                    False,          # OffsetReverse2
                    False,          # TranslateSurface1
                    False,          # TranslateSurface2
                    True,           # Merge - merge result
                    True,           # UseFeatScope
                    True,           # UseAutoSelect
                    0,              # T0 - start condition
                    0.0,            # StartOffset
                    False           # FlipStartOffset
                )
                if feat:
                    method_used = "FeatureExtrusion2_23p"
            except Exception as e:
                logger.debug(f"FeatureExtrusion2 (23p) failed: {e}")

            # Method 2: FeatureExtrusion2 with 20 params (older SW versions)
            if feat is None:
                try:
                    feat = doc.FeatureManager.FeatureExtrusion2(
                        True, False, False, end_cond, 0,
                        depth_m, depth_m,
                        False, False, False, False,
                        0.0, 0.0,
                        False, False, False,
                        True, True, True, True
                    )
                    if feat:
                        method_used = "FeatureExtrusion2_20p"
                except Exception as e:
                    logger.debug(f"FeatureExtrusion2 (20p) failed: {e}")

            # Method 3: FeatureExtrusion3 (some SW versions)
            if feat is None:
                try:
                    feat = doc.FeatureManager.FeatureExtrusion3(
                        True, False, False, end_cond, 0,
                        depth_m, 0,
                        False, False, False, False,
                        0.0, 0.0,
                        False, False, False,
                        True, True, True,
                        0, 0.0, False
                    )
                    if feat:
                        method_used = "FeatureExtrusion3"
                except Exception as e:
                    logger.debug(f"FeatureExtrusion3 failed: {e}")

            if feat is None:
                sketch_info = self._get_sketch_info(doc)
                return self._result(False,
                    f"Extrusion failed on sketch '{sketch_name}'. "
                    f"Ensure sketch has a closed profile (circle, rectangle, etc). "
                    f"Sketches in model: {sketch_info['sketch_names']}",
                    SwErrors.swFeatureError,
                    {"sketch_name": sketch_name, "diagnostics": sketch_info})

            direction = "both directions (mid-plane)" if both_directions else "one direction"

            return self._result(True,
                f"Extruded {depth}{unit_str} ({direction}) [{method_used}]",
                SwErrors.swSuccess,
                {"depth": depth, "unit": unit_str,
                 "both_directions": both_directions,
                 "sketch_name": sketch_name,
                 "api_method": method_used})

        except Exception as e:
            logger.error(f"Extrude error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)

    # ========================================================================
    # Revolve
    # ========================================================================

    def revolve_sketch(self, angle: float = 360, cut: bool = False,
                       reverse: bool = False, axis: str = None) -> Dict:
        """
        Revolve the active sketch around its centerline (Boss/Cut-Revolve).
        The sketch must contain exactly one centerline -- FeatureRevolve2
        picks it up automatically as the revolve axis without an explicit
        selection.

        FeatureRevolve2's real signature is 20 positional params (confirmed
        against the SolidWorks typelib (lookup_api_signature) -- see CLAUDE.md
        section). Earlier attempts at 18 and 21 params both failed
        (`Parameter not optional` / `Invalid number of parameters`); the
        two commonly-missed ones are OffsetDistance1/OffsetDistance2,
        which sit between OffsetReverse2 and ThinType.

        Args:
            angle: Revolve angle in degrees (360 = full revolve)
            cut: True for a cut-revolve (remove material) instead of boss

        Returns:
            Result dictionary
        """
        try:
            import math
            doc, err = self.get_active_doc()
            if err:
                return err

            if axis and str(axis).strip().upper() in ("X", "Y", "Z"):
                # create the world axis BEFORE selecting the sketch: making it
                # clears the selection (and must not happen inside the sketch)
                from .. import ext
                if doc.SketchManager.ActiveSketch is not None:
                    doc.SketchManager.InsertSketch(True)
                ext.world_axis(ext.model(self._sw_app), str(axis).strip())

            success, sketch_name, error_msg = self._close_and_select_sketch(doc)
            if not success:
                sketch_info = self._get_sketch_info(doc)
                return self._result(False,
                    f"Revolve failed: {error_msg}. "
                    f"Sketches found: {sketch_info['sketch_count']} {sketch_info['sketch_names']}. "
                    f"Sketch needs exactly one centerline as the revolve axis.",
                    SwErrors.swFeatureError,
                    {"diagnostics": sketch_info})

            if axis:
                # axis outside the sketch (world X/Y/Z, a reference axis, an
                # edge): selected with Mark 4 -- no centerline to define
                from .. import ext
                ext.select_direction(ext.model(self._sw_app), axis, mark=4, append=True)

            angle_rad = math.radians(angle)

            feat = None
            try:
                feat = doc.FeatureManager.FeatureRevolve2(
                    True,           # SingleDir
                    True,           # IsSolid (False = surface revolve, even for cuts)
                    False,          # IsThin
                    cut,            # IsCut
                    reverse,        # ReverseDir
                    False,          # BothDirectionUpToSameEntity
                    0, 0,           # Dir1Type, Dir2Type (0 = blind/angle-driven)
                    angle_rad, 0.0, # Dir1Angle, Dir2Angle
                    False, False,   # OffsetReverse1, OffsetReverse2
                    0.0, 0.0,       # OffsetDistance1, OffsetDistance2
                    0,              # ThinType
                    0.0, 0.0,       # ThinThickness1, ThinThickness2
                    True, True, True  # Merge, UseFeatScope, UseAutoSelect
                )
            except Exception as e:
                logger.debug(f"FeatureRevolve2 failed: {e}")

            if feat is None:
                sketch_info = self._get_sketch_info(doc)
                return self._result(False,
                    f"Revolve failed on sketch '{sketch_name}'. "
                    f"Needs a closed profile plus exactly one centerline "
                    f"(construction line) as the axis.",
                    SwErrors.swFeatureError,
                    {"sketch_name": sketch_name, "diagnostics": sketch_info})

            kind = "Cut-Revolve" if cut else "Revolve"
            return self._result(True, f"{kind}: {angle}°",
                              SwErrors.swSuccess,
                              {"angle": angle, "cut": cut, "sketch_name": sketch_name})

        except Exception as e:
            logger.error(f"Revolve error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)

    # ========================================================================
    # Cut Extrude
    # ========================================================================

    def cut_extrude(self, depth: float = 10, through_all: bool = False,
                    both_directions: bool = False, unit: str = None,
                    reverse: bool = False) -> Dict:
        """
        Cut extrude (remove material).

        Direction is verified, not assumed: a sketch on a face wants Dir=False,
        a sketch on a reference plane in this install only cuts with Dir=True
        (NOTES.md "FeatureCut3/4 на эскизе, лежащем на референс-плоскости").
        So: try the requested direction, check that volume actually dropped,
        otherwise delete the empty feature and retry the other way.

        Args:
            depth: Cut depth (ignored if through_all=True)
            through_all: Cut through entire model
            both_directions: Cut in both directions
            unit: Unit for depth
            reverse: Start with the opposite direction
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            from .. import ext

            depth_m = self._units.to_meters(depth, unit)
            unit_str = unit or self._units.default_unit.value

            success, sketch_name, error_msg = self._close_and_select_sketch(doc)
            if not success:
                sketch_info = self._get_sketch_info(doc)
                return self._result(False,
                    f"Cut failed: {error_msg}. "
                    f"Sketches: {sketch_info['sketch_names']}.",
                    SwErrors.swFeatureError,
                    {"diagnostics": sketch_info})

            if through_all:
                end_cond = 2 if both_directions else 1  # ThroughAllBoth / ThroughAll
                cut_depth = 0.0
            else:
                end_cond = 6 if both_directions else 0  # MidPlane / Blind
                cut_depth = depth_m

            md = ext.model(self._sw_app)
            v0 = ext.volume_mm3(md)
            tried = []
            feat = None
            for dir_flag in (reverse, not reverse):
                if tried:
                    # re-select the sketch for the second attempt
                    doc.ClearSelection2(True)
                    ext.find_feature(md, sketch_name).Select2(False, 0)
                tried.append(dir_flag)
                feat = doc.FeatureManager.FeatureCut3(
                    True, False, dir_flag,      # Sd, Flip, Dir
                    end_cond, 0,                # T1, T2
                    cut_depth, 0.0,             # D1, D2
                    False, False, False, False, # Dchk1, Dchk2, Ddir1, Ddir2
                    0.0, 0.0,                   # Dang1, Dang2
                    False, False,               # OffsetReverse1/2
                    False, False,               # TranslateSurface1/2
                    False,                      # NormalCut
                    True, True,                 # UseFeatScope, UseAutoSelect
                    True, True, False,          # AssemblyFeatureScope, AutoSelectComponents, Propagate
                    0, 0.0, False)              # T0, StartOffset, FlipStartOffset
                if feat is None:
                    continue
                if ext.volume_mm3(md) < v0 - 1e-6:
                    break
                # feature created but removed nothing -> wrong side, drop it
                name = ext.T(feat, "IFeature").Name
                doc.ClearSelection2(True)
                ext.find_feature(md, name).Select2(False, 0)
                md.Extension.DeleteSelection2(0)
                feat = None

            if feat is None:
                return self._result(False,
                    f"Cut on '{sketch_name}' removed nothing in either direction "
                    f"(tried Dir={tried}). Closed profile? Does it overlap the body?",
                    SwErrors.swFeatureError,
                    {"sketch_name": sketch_name})

            cut_type = "through all" if through_all else f"{depth}{unit_str}"
            return self._result(True, f"Cut extrude: {cut_type} [Dir={tried[-1]}]",
                              SwErrors.swSuccess)

        except Exception as e:
            logger.error(f"Cut extrude error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)

    # ========================================================================
    # Fillet
    # ========================================================================

    def fillet_edges(self, radius: float = 2, unit: str = None) -> Dict:
        """
        Add fillet to selected edges

        Args:
            radius: Fillet radius
            unit: Unit for radius

        Returns:
            Result dictionary

        Note: Select edges first using execute_python or manual selection
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            radius_m = self._units.to_meters(radius, unit)

            feat = None
            method_used = ""

            # Method 1: FeatureFillet3
            try:
                # 14 params (typelib): Options, R1, R2, Rho, Ftyp,
                # OverflowType, ConicRhoType, Radii, Dist2Arr, RhoArr,
                # SetBackDistances, PointRadiusArray, PointDist2Array,
                # PointRhoArray. The upstream 15-arg call always failed.
                feat = doc.FeatureManager.FeatureFillet3(
                    195, radius_m, 0.0, 0.0, 0, 0, 0,
                    None, None, None, None, None, None, None
                )
                if feat:
                    method_used = "FeatureFillet3"
            except Exception as e:
                logger.debug(f"FeatureFillet3 failed: {e}")

            # Method 2: SimpleFillet
            if feat is None:
                try:
                    feat = doc.FeatureManager.SimpleFillet(radius_m, True, True, True)
                    if feat:
                        method_used = "SimpleFillet"
                except Exception as e:
                    logger.debug(f"SimpleFillet failed: {e}")

            if feat is None:
                return self._result(False,
                    "Fillet failed - select edges first (use execute_python to select edges programmatically)",
                    SwErrors.swFeatureError)

            unit_str = unit or self._units.default_unit.value

            return self._result(True, f"Fillet: r={radius}{unit_str} [{method_used}]",
                              SwErrors.swSuccess,
                              {"radius": radius, "unit": unit_str})

        except Exception as e:
            logger.error(f"Fillet error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)

    # ========================================================================
    # Chamfer
    # ========================================================================

    def chamfer_edges(self, distance: float = 2, angle: float = 45,
                      unit: str = None) -> Dict:
        """
        Add chamfer to selected edges
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            import math
            dist_m = self._units.to_meters(distance, unit)
            angle_rad = math.radians(angle)

            feat = None

            try:
                # InsertFeatureChamfer takes 8 params, confirmed against the
                # real typelib (lookup_api_signature) --
                # (Options, ChamferType, Width, Angle, OtherDist,
                # VertexChamDist1, VertexChamDist2, VertexChamDist3).
                # ChamferType=0 is angle-distance (Width + Angle); the trailing
                # three VertexChamDist args only apply to vertex chamfers.
                # ChamferType 1 = swChamferAngleDistance (0 is not a valid
                # swChamferType_e value)
                feat = doc.FeatureManager.InsertFeatureChamfer(
                    0, 1, dist_m, angle_rad, 0.0, 0.0, 0.0, 0.0
                )
            except Exception as e:
                logger.debug(f"InsertFeatureChamfer failed: {e}")

            if feat is None:
                return self._result(False,
                    "Chamfer failed - select edges first",
                    SwErrors.swFeatureError)

            unit_str = unit or self._units.default_unit.value

            return self._result(True, f"Chamfer: {distance}{unit_str} x {angle}\u00b0",
                              SwErrors.swSuccess,
                              {"distance": distance, "angle": angle, "unit": unit_str})

        except Exception as e:
            logger.error(f"Chamfer error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFeatureError)

    # ========================================================================
    # List Features (FIXED: properties not methods)
    # ========================================================================

    def list_features(self) -> Dict:
        """List user features (after Origin) of the active document."""
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            from .. import ext
            feats = [{"name": f.Name, "type": ext.v(f, "GetTypeName2")}
                     for f in ext.user_features(ext.T(doc, "IModelDoc2"))]
            return self._result(True, f"{len(feats)} features found",
                              SwErrors.swSuccess, {"features": feats, "count": len(feats)})
        except Exception as e:
            logger.error(f"List features error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swUnknownError)

    # ========================================================================
    # Edge Selection Helper
    # ========================================================================

    def select_edge(self, edge_index: int = 1) -> Dict:
        """
        Select an edge by index
        Note: Use execute_python for precise edge selection
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            doc.ClearSelection2(True)

            return self._result(True, "Use execute_python for edge selection",
                              SwErrors.swSuccess)

        except Exception as e:
            return self._result(False, f"Error: {e}", SwErrors.swSelectionError)
