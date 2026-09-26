"""
SolidWorks View Capture
------------------------
Screenshot the active document's current view for downstream
vision/OCR-based indexing (see CLAUDE.md: screenshot -> guess -> refine
with the geometry tools in features.py).
"""

import os
import logging
import traceback
import datetime
from typing import Optional, Dict

from ..constants import SwErrors
from ..utils.com_helpers import v

logger = logging.getLogger(__name__)

# SaveBMP is the one screenshot API that has worked unchanged across
# SolidWorks versions in this codebase's testing; PNG conversion happens
# in Python afterwards since SW's own PNG export methods vary by version.
_NAMED_VIEWS = {
    "isometric": 7, "front": 1, "back": 2, "left": 3, "right": 4,
    "top": 5, "bottom": 6, "trimetric": 8, "dimetric": 9,
}


class ViewCaptureOperations:
    """
    Mixin class for screenshotting the active view.

    Requires parent class to have:
    - get_active_doc(): Document access method
    - _result(): Result factory method
    - _get_doc_title(): Document title helper
    - self._config: SolidWorksConfig instance
    """

    def capture_view(self, output_path: Optional[str] = None,
                      view: Optional[str] = None,
                      width: int = None, height: int = None,
                      zoom_to_fit: bool = True) -> Dict:
        """
        Screenshot the active document's current view to a PNG file.

        Args:
            output_path: Where to save the .png (default: alongside the
                document, in a `_screenshots` subfolder, named after the
                document title + timestamp)
            view: Optional named view to switch to first
                (isometric/front/back/left/right/top/bottom/trimetric/dimetric)
            width, height: Capture resolution in pixels (defaults from config)
            zoom_to_fit: Zoom to fit the model in view before capturing

        Returns:
            Result dictionary with {"path": <png path>, "width", "height"}
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            width = width or self._config.capture_width
            height = height or self._config.capture_height

            if view:
                view_key = view.strip().lower()
                if view_key not in _NAMED_VIEWS:
                    return self._result(False,
                        f"Unknown view '{view}'. Options: {', '.join(_NAMED_VIEWS)}",
                        SwErrors.swInvalidInput)
                try:
                    doc.ShowNamedView2("", _NAMED_VIEWS[view_key])
                except Exception as e:
                    logger.debug(f"ShowNamedView2 failed: {e}")

            if zoom_to_fit:
                try:
                    v(doc, "ViewZoomtofit2")
                except Exception as e:
                    logger.debug(f"ViewZoomtofit2 failed: {e}")

            if output_path is None:
                title = self._get_doc_title(doc)
                safe_title = "".join(c if c.isalnum() or c in "-_ " else "_"
                                      for c in title).strip() or "part"
                doc_path = self._get_doc_path(doc)
                if doc_path:
                    base_dir = os.path.join(os.path.dirname(doc_path), "_screenshots")
                else:
                    # Unsaved document: os.getcwd() is whatever directory the
                    # SolidWorks process happened to start in (often a
                    # non-writable system dir like C:\Windows\System32, not
                    # this script's cwd) -- fall back to somewhere always
                    # writable instead of guessing.
                    base_dir = os.path.join(os.path.expanduser("~"),
                                             "SolidWorks_Screenshots")
                os.makedirs(base_dir, exist_ok=True)
                stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = os.path.join(base_dir, f"{safe_title}_{stamp}.png")

            output_path = os.path.abspath(output_path)
            bmp_path = os.path.splitext(output_path)[0] + ".bmp"

            ok = doc.SaveBMP(bmp_path, width, height)
            if not ok:
                return self._result(False,
                    "SaveBMP failed - is a document actually visible/active?",
                    SwErrors.swFeatureError)

            try:
                from PIL import Image
                with Image.open(bmp_path) as im:
                    im.save(output_path, "PNG")
                os.remove(bmp_path)
                final_path = output_path
            except ImportError:
                logger.warning("Pillow not installed - leaving screenshot as BMP")
                final_path = bmp_path
            except Exception as e:
                logger.error(f"BMP->PNG conversion failed: {e}")
                final_path = bmp_path

            return self._result(True, f"Captured view to {final_path}",
                              SwErrors.swSuccess,
                              {"path": final_path, "width": width, "height": height,
                               "view": view or "current"})

        except Exception as e:
            logger.error(f"Capture view error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swUnknownError)
