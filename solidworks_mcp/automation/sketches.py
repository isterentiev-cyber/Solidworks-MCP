"""
SolidWorks Sketch Operations
----------------------------
Create sketches and draw 2D geometry.
"""

import math
import logging
import traceback
from typing import Optional, Dict, List, Tuple

import win32com.client
import pythoncom

from ..constants import SwErrors, SwPlanes

logger = logging.getLogger(__name__)


class SketchOperations:
    """
    Mixin class for sketch operations

    Requires parent class to have:
    - get_active_doc(): Document access method
    - _result(): Result factory method
    - _units: UnitConverter instance
    """

    def create_sketch(self, plane: str = "Front", offset: float = 0,
                      unit: str = None) -> Dict:
        """
        Create a new sketch on a default plane, optionally on a parallel
        reference plane at `offset` from it.

        Default planes are resolved by feature-tree order (first three
        RefPlanes = Front/Top/Right), so this works on localized SW installs
        where they are not called "Front Plane" etc.
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            from .. import ext
            md = ext.model(self._sw_app)
            plane_name = ext.resolve_plane(md, plane)

            doc.ClearSelection2(True)
            if not ext.select_by_id(doc.Extension, plane_name, "PLANE"):
                return self._result(False, f"Could not select {plane_name}",
                                  SwErrors.swSelectionError)

            if offset:
                off_m = self._units.to_meters(offset, unit)
                flags = 8 | (256 if off_m < 0 else 0)  # Distance | OptionFlip
                ref = ext.T(doc.FeatureManager, "IFeatureManager").InsertRefPlane(
                    flags, float(abs(off_m)), 0, 0.0, 0, 0.0)
                if ref is None:
                    return self._result(False, f"InsertRefPlane failed ({offset} from {plane_name})",
                                      SwErrors.swFeatureError)
                # InsertRefPlane returns an IRefPlane, not the IFeature --
                # the new feature is the last one in the tree.
                ref = list(ext.user_features(md))[-1]
                tag = ext.fnum(offset).replace(".", "_").replace("-", "m")
                name = f"MCP_{plane_name.split()[0]}_{tag}"
                if name not in [f.Name for f in ext.user_features(md)]:
                    ref.Name = name
                doc.ClearSelection2(True)
                ref.Select2(False, 0)
                plane_name = ref.Name

            doc.InsertSketch2(True)
            frame = ext.sketch_frame(self._sw_app)
            return self._result(True, f"Sketch on {plane_name}. {frame}",
                              SwErrors.swSuccess, {"plane": plane_name})

        except Exception as e:
            logger.error(f"Create sketch error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)

    def create_sketch_on_face(self, x: float = 0, y: float = 0, z: float = 0,
                              unit: str = None, face_index: int = 0) -> Dict:
        """
        Create a new sketch on a face selected by coordinate.
        ADDED v4.1: Enables cut-extrude on existing body faces (not just ref planes).

        Args:
            x, y, z: Coordinates (in user units) of a point ON the target face
            unit: Unit for coordinates (uses default if None)

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            # Convert to meters
            x_m = self._units.to_meters(x, unit)
            y_m = self._units.to_meters(y, unit)
            z_m = self._units.to_meters(z, unit)
            unit_str = unit or self._units.default_unit.value

            # Clear any selection
            doc.ClearSelection2(True)

            from .. import ext
            if face_index:
                ext.select_entities(ext.model(self._sw_app), "face", [face_index])
                doc.InsertSketch2(True)
                return self._result(True,
                    f"Sketch on face [{face_index}]. {ext.sketch_frame(self._sw_app)}",
                    SwErrors.swSuccess, {"face_index": face_index})

            # Select face at the given coordinates
            try:
                ext.select_face_at(ext.model(self._sw_app), [x_m * 1000, y_m * 1000, z_m * 1000])
                selected = True
            except ext.ExtError:
                selected = False

            if not selected:
                return self._result(False,
                    f"No face found at ({x}{unit_str}, {y}{unit_str}, {z}{unit_str}). "
                    f"Coordinates must be on an existing body face.",
                    SwErrors.swSelectionError)

            # Insert sketch on the selected face
            doc.InsertSketch2(True)

            return self._result(True,
                f"Sketch on face at ({x}, {y}, {z}) {unit_str}. {ext.sketch_frame(self._sw_app)}",
                SwErrors.swSuccess,
                {"face_point": {"x": x, "y": y, "z": z}, "unit": unit_str})

        except Exception as e:
            logger.error(f"Create sketch on face error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)

    def exit_sketch(self) -> Dict:
        """
        Exit the current sketch

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            doc.InsertSketch2(True)

            return self._result(True, "Exited sketch", SwErrors.swSuccess)

        except Exception as e:
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)

    def draw_line(self, x1: float = 0, y1: float = 0,
                  x2: float = 100, y2: float = 0, unit: str = None,
                  construction: bool = False) -> Dict:
        """
        Draw a line in the active sketch

        Args:
            x1, y1: Start point coordinates
            x2, y2: End point coordinates
            unit: Unit for coordinates (uses default if None)

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            # Convert to meters
            x1_m = self._units.to_meters(x1, unit)
            y1_m = self._units.to_meters(y1, unit)
            x2_m = self._units.to_meters(x2, unit)
            y2_m = self._units.to_meters(y2, unit)

            # Draw line (centerline if construction -- revolve axis)
            from ..ext import NoInference
            sm = doc.SketchManager
            with NoInference(sm):
                fn = sm.CreateCenterLine if construction else sm.CreateLine
                line = fn(x1_m, y1_m, 0, x2_m, y2_m, 0)

            if line is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)

            # Calculate length
            length_m = math.sqrt((x2_m - x1_m)**2 + (y2_m - y1_m)**2)
            length_display = self._units.from_meters(length_m, unit)
            unit_str = unit or self._units.default_unit.value

            return self._result(True, f"Line: {length_display:.2f}{unit_str}",
                              SwErrors.swSuccess,
                              {"length": length_display, "unit": unit_str})

        except Exception as e:
            logger.error(f"Draw line error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)

    def draw_circle(self, x: float = 0, y: float = 0,
                    radius: float = 25, unit: str = None) -> Dict:
        """
        Draw a circle in the active sketch

        Args:
            x, y: Center point coordinates
            radius: Circle radius
            unit: Unit for dimensions (uses default if None)

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            # Convert to meters
            x_m = self._units.to_meters(x, unit)
            y_m = self._units.to_meters(y, unit)
            radius_m = self._units.to_meters(radius, unit)

            # Draw circle (center point and edge point). Without
            # AddToDB/AutoInference-off CreateCircle silently returns None
            # in this install (NOTES.md).
            from ..ext import NoInference
            with NoInference(doc.SketchManager) as sm:
                circle = sm.CreateCircle(
                    x_m, y_m, 0,              # Center
                    x_m + radius_m, y_m, 0    # Edge point
                )

            if circle is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)

            unit_str = unit or self._units.default_unit.value

            return self._result(True, f"Circle: r={radius}{unit_str}",
                              SwErrors.swSuccess,
                              {"radius": radius, "unit": unit_str,
                               "center_x": x, "center_y": y})

        except Exception as e:
            logger.error(f"Draw circle error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)

    def draw_rectangle(self, x1: float = -50, y1: float = -25,
                       x2: float = 50, y2: float = 25, unit: str = None) -> Dict:
        """
        Draw a rectangle in the active sketch

        Args:
            x1, y1: First corner coordinates
            x2, y2: Opposite corner coordinates
            unit: Unit for dimensions

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            # Convert to meters
            x1_m = self._units.to_meters(x1, unit)
            y1_m = self._units.to_meters(y1, unit)
            x2_m = self._units.to_meters(x2, unit)
            y2_m = self._units.to_meters(y2, unit)

            # Four lines instead of CreateCornerRectangle, which silently
            # returns None in this install (NOTES.md).
            from ..ext import NoInference
            corners = [(x1_m, y1_m), (x2_m, y1_m), (x2_m, y2_m), (x1_m, y2_m)]
            with NoInference(doc.SketchManager) as sm:
                made = [sm.CreateLine(a[0], a[1], 0, b[0], b[1], 0)
                        for a, b in zip(corners, corners[1:] + corners[:1])]
            rect = None if any(m is None for m in made) else made

            if rect is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)

            width = abs(x2 - x1)
            height = abs(y2 - y1)
            unit_str = unit or self._units.default_unit.value

            return self._result(True, f"Rectangle: {width}x{height}{unit_str}",
                              SwErrors.swSuccess,
                              {"width": width, "height": height, "unit": unit_str})

        except Exception as e:
            logger.error(f"Draw rectangle error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)

    def draw_arc_center(self, cx: float = 0, cy: float = 0, radius: float = 25,
                        start_angle: float = 0, end_angle: float = 90,
                        unit: str = None) -> Dict:
        """
        Draw an arc by center point and angles

        Args:
            cx, cy: Center point
            radius: Arc radius
            start_angle: Start angle in degrees
            end_angle: End angle in degrees
            unit: Unit for radius

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            # Convert to meters
            cx_m = self._units.to_meters(cx, unit)
            cy_m = self._units.to_meters(cy, unit)
            radius_m = self._units.to_meters(radius, unit)

            # Convert angles to radians
            start_rad = math.radians(start_angle)
            end_rad = math.radians(end_angle)

            # Calculate start and end points
            x1 = cx_m + radius_m * math.cos(start_rad)
            y1 = cy_m + radius_m * math.sin(start_rad)
            x2 = cx_m + radius_m * math.cos(end_rad)
            y2 = cy_m + radius_m * math.sin(end_rad)

            # Draw arc (center, start, end, direction)
            from ..ext import NoInference
            with NoInference(doc.SketchManager) as sm:
                arc = sm.CreateArc(
                    cx_m, cy_m, 0,   # Center
                    x1, y1, 0,       # Start
                    x2, y2, 0,       # End
                    1                # Direction (1 = counter-clockwise)
                )

            if arc is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)

            unit_str = unit or self._units.default_unit.value
            arc_angle = abs(end_angle - start_angle)

            return self._result(True, f"Arc: r={radius}{unit_str}, {arc_angle}°",
                              SwErrors.swSuccess,
                              {"radius": radius, "start_angle": start_angle,
                               "end_angle": end_angle, "unit": unit_str})

        except Exception as e:
            logger.error(f"Draw arc error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)

    def draw_polygon(self, cx: float = 0, cy: float = 0, radius: float = 25,
                     sides: int = 6, unit: str = None) -> Dict:
        """
        Draw a regular polygon

        Args:
            cx, cy: Center point
            radius: Circumscribed radius
            sides: Number of sides (3-100)
            unit: Unit for radius

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            if sides < 3 or sides > 100:
                return self._result(False, "Sides must be between 3 and 100",
                                  SwErrors.swInvalidInput)

            # Convert to meters
            cx_m = self._units.to_meters(cx, unit)
            cy_m = self._units.to_meters(cy, unit)
            radius_m = self._units.to_meters(radius, unit)

            # Draw polygon
            from ..ext import NoInference
            with NoInference(doc.SketchManager) as sm:
                polygon = sm.CreatePolygon(
                    cx_m, cy_m, 0,               # Center
                    cx_m + radius_m, cy_m, 0,    # Vertex
                    sides,                        # Number of sides
                    False                         # Inscribed (False = circumscribed)
                )

            if polygon is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)

            unit_str = unit or self._units.default_unit.value

            return self._result(True, f"{sides}-sided polygon: r={radius}{unit_str}",
                              SwErrors.swSuccess,
                              {"sides": sides, "radius": radius, "unit": unit_str})

        except Exception as e:
            logger.error(f"Draw polygon error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)

    def draw_centerline(self, x1: float = 0, y1: float = -50,
                        x2: float = 0, y2: float = 50, unit: str = None) -> Dict:
        """
        Draw a centerline (for revolve axis)

        Args:
            x1, y1: Start point
            x2, y2: End point
            unit: Unit for coordinates

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            # Convert to meters
            x1_m = self._units.to_meters(x1, unit)
            y1_m = self._units.to_meters(y1, unit)
            x2_m = self._units.to_meters(x2, unit)
            y2_m = self._units.to_meters(y2, unit)

            # Draw centerline
            line = doc.SketchManager.CreateCenterLine(x1_m, y1_m, 0, x2_m, y2_m, 0)

            if line is None:
                return self._result(False, "Failed - ensure sketch is active",
                                  SwErrors.swSketchError)

            return self._result(True, "Centerline created",
                              SwErrors.swSuccess)

        except Exception as e:
            logger.error(f"Draw centerline error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swSketchError)
