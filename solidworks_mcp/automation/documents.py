"""
SolidWorks Document Operations
------------------------------
Create, open, save, and manage SolidWorks documents.
"""

import os
import logging
import traceback
from typing import Optional, Dict

import pythoncom

from ..constants import SwErrors, SwDocumentTypes, SwFileTypes
from ..utils import com_get
from ..utils.com_helpers import v, nothing, call_out, OUT
from ..utils.typelib import enum_flags

logger = logging.getLogger(__name__)


class DocumentOperations:
    """
    Mixin class for document operations
    
    Requires parent class to have:
    - self._sw_app: SolidWorks application object
    - self.is_connected: Connection status property
    - self.connect(): Connection method
    - self._result(): Result factory method
    - self._units: UnitConverter instance
    """
    
    def create_new_part(self) -> Dict:
        """
        Create a new part document
        
        Returns:
            Result dictionary with document info
        """
        try:
            if not self.is_connected:
                r = self.connect()
                if not r["success"]:
                    return r
            
            # Find part template. NewDocument("", ...) is NOT "use SW's
            # default" -- it fails with swFileLoadError, so a real path is
            # required (see _get_template in base.py).
            template = self._get_template("part")
            if not template:
                return self._result(False,
                    "No part template found (disk search and the running "
                    "app's own default template preference both came up "
                    "empty). Set exe_path/part_template in config, or set "
                    "a default part template in SW's own "
                    "Options > Default Templates.",
                    SwErrors.swTemplateNotFound)
            logger.info(f"Using template: {template}")

            # Create document
            doc = self._sw_app.NewDocument(template, 0, 0, 0)

            if doc is None:
                return self._result(False, "Failed to create part document",
                                  SwErrors.swFileLoadError)
            
            # Set view
            try:
                doc.ShowNamedView2("*Isometric", 7)
                v(doc, "ViewZoomtofit2")
            except:
                pass
            
            title = self._get_doc_title(doc)
            
            return self._result(True, f"Created part: {title}",
                              SwErrors.swSuccess,
                              {"name": title, "type": "Part"})
            
        except Exception as e:
            logger.error(f"Create part error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFileLoadError)
    
    def create_new_assembly(self) -> Dict:
        """
        Create a new assembly document
        
        Returns:
            Result dictionary with document info
        """
        try:
            if not self.is_connected:
                r = self.connect()
                if not r["success"]:
                    return r
            
            template = self._get_template("assembly")
            if not template:
                return self._result(False,
                    "No assembly template found (disk search and the "
                    "running app's own default template preference both "
                    "came up empty). Set exe_path/assembly_template in "
                    "config, or set a default assembly template in SW's "
                    "own Options > Default Templates.",
                    SwErrors.swTemplateNotFound)

            doc = self._sw_app.NewDocument(template, 0, 0, 0)

            if doc is None:
                return self._result(False, "Failed to create assembly",
                                  SwErrors.swFileLoadError)
            
            try:
                doc.ShowNamedView2("*Isometric", 7)
                v(doc, "ViewZoomtofit2")
            except:
                pass
            
            title = self._get_doc_title(doc)
            
            return self._result(True, f"Created assembly: {title}",
                              SwErrors.swSuccess,
                              {"name": title, "type": "Assembly"})
            
        except Exception as e:
            logger.error(f"Create assembly error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFileLoadError)
    
    def create_new_drawing(self, paper_size: str = "A4") -> Dict:
        """
        Create a new drawing document
        
        Args:
            paper_size: Paper size (A4, A3, A2, A1, Letter)
        
        Returns:
            Result dictionary with document info
        """
        try:
            if not self.is_connected:
                r = self.connect()
                if not r["success"]:
                    return r
            
            template = self._get_template("drawing")
            if not template:
                return self._result(False,
                    "No drawing template found (disk search and the "
                    "running app's own default template preference both "
                    "came up empty -- common if SW was never pointed at a "
                    "drawing template/sheet format). Set "
                    "exe_path/drawing_template in config, or set a "
                    "default drawing template in SW's own "
                    "Options > Default Templates.",
                    SwErrors.swTemplateNotFound)

            doc = self._sw_app.NewDocument(template, 0, 0, 0)

            if doc is None:
                return self._result(False, "Failed to create drawing",
                                  SwErrors.swFileLoadError)
            
            title = self._get_doc_title(doc)
            
            return self._result(True, f"Created drawing: {title}",
                              SwErrors.swSuccess,
                              {"name": title, "type": "Drawing", "paper_size": paper_size})
            
        except Exception as e:
            logger.error(f"Create drawing error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFileLoadError)
    
    def open_document(self, filepath: str, resolve_lightweight: bool = True,
                      read_only: bool = False) -> Dict:
        """
        Open an existing document (late-bound OpenDoc6, silent).

        Args:
            filepath: Path to SolidWorks file
            resolve_lightweight: for assemblies, resolve lightweight components
                right after opening (Large Assembly Mode loads everything
                lightweight and GetModelDoc2 then returns None for every part)
            read_only: open read-only

        Returns:
            Result dictionary; data has errors/warnings (decoded) and, for an
            assembly, component counts
        """
        try:
            if not self.is_connected:
                r = self.connect()
                if not r["success"]:
                    return r

            if not os.path.exists(filepath):
                return self._result(False, f"File not found: {filepath}",
                                  SwErrors.swFileNotFoundError)

            ext = os.path.splitext(filepath)[1].lower()
            type_map = {
                ".sldprt": SwDocumentTypes.swDocPART,
                ".sldasm": SwDocumentTypes.swDocASSEMBLY,
                ".slddrw": SwDocumentTypes.swDocDRAWING,
            }
            doc_type = type_map.get(ext, SwDocumentTypes.swDocPART)

            # swOpenDocOptions_Silent (1) | swOpenDocOptions_ReadOnly (2)
            options = 1 | (2 if read_only else 0)
            doc, errors, warnings = call_out(self._sw_app, "OpenDoc6", filepath,
                                             int(doc_type), options, "", OUT, OUT)
            err_names = enum_flags("swFileLoadError_e", errors)
            warn_names = enum_flags("swFileLoadWarning_e", warnings)

            if doc is None:
                return self._result(False,
                    f"Failed to open (errors {errors}: {', '.join(err_names) or '?'}; "
                    f"warnings {warnings}: {', '.join(warn_names) or '-'})",
                    SwErrors.swFileLoadError,
                    {"errors": errors, "error_names": err_names,
                     "warnings": warnings, "warning_names": warn_names})

            title = self._get_doc_title(doc)
            data = {"name": title, "path": filepath,
                    "errors": errors, "error_names": err_names,
                    "warnings": warnings, "warning_names": warn_names}
            msg = f"Opened: {title}"
            if errors or warnings:
                msg += f" | errors {errors} {err_names} warnings {warnings} {warn_names}"

            if int(doc_type) == SwDocumentTypes.swDocASSEMBLY:
                data.update(self._assembly_load_state(doc, resolve_lightweight))
                msg += (f" | components {data['components']} "
                        f"(top level {data['top_level']}), "
                        f"lightweight {data['lightweight_before']}")
                if data.get("resolved"):
                    msg += f" -> resolved in {data['resolve_s']}s"
                if data.get("large_assembly_mode"):
                    msg += ", Large Assembly Mode"

            return self._result(True, msg, SwErrors.swSuccess, data)

        except Exception as e:
            logger.error(f"Open document error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFileLoadError)

    def _assembly_load_state(self, doc, resolve: bool) -> Dict:
        """Component counts and lightweight state of a freshly opened assembly;
        optionally resolves lightweight components."""
        import time
        from .. import ext
        asm = ext.T(doc, "IAssemblyDoc")
        comps = v(asm, "GetComponents", False) or ()
        top = v(asm, "GetComponents", True) or ()
        # swComponentSuppressionState_e: 1 = lightweight, 4 = fully lightweight
        light = sum(1 for c in comps if v(c, "GetSuppression2") in (1, 4))
        info = {"components": len(comps), "top_level": len(top),
                "lightweight_before": light, "resolved": False}
        try:
            info["large_assembly_mode"] = bool(v(ext.T(doc, "IModelDoc2"), "LargeAssemblyMode"))
        except Exception:
            pass
        if resolve and light:
            t0 = time.time()
            v(asm, "ResolveAllLightWeightComponents", False)
            info["resolved"] = True
            info["resolve_s"] = round(time.time() - t0, 1)
        return info

    def save_document(self, filepath: str = None) -> Dict:
        """
        Save the active document.

        Args:
            filepath: Save As path (None = save in place). Save As switches the
                open document to the new file, like SolidWorks' own Save As.

        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err

            from .. import ext
            md = ext.T(doc, "IModelDoc2")

            if filepath:
                filepath = os.path.abspath(filepath)
                dir_path = os.path.dirname(filepath)
                if dir_path and not os.path.exists(dir_path):
                    os.makedirs(dir_path)

                # IModelDocExtension.SaveAs3(Name, Version, Options, ExportData,
                # AdvancedSaveAsOptions, [out] Errors, [out] Warnings)
                # swSaveAsCurrentVersion = 0, swSaveAsOptions_Silent = 1
                ok, errs, warns = call_out(v(md, "Extension"), "SaveAs3", filepath, 0, 1,
                                           nothing(), nothing(), OUT, OUT)
                if not ok or errs:
                    names = enum_flags("swFileSaveError_e", errs)
                    return self._result(False,
                        f"Save As failed (errors {errs}: {', '.join(names) or '?'}, "
                        f"warnings {warns})", SwErrors.swFileSaveError,
                        {"errors": errs, "error_names": names, "warnings": warns})
                return self._result(True, f"Saved: {filepath}", SwErrors.swSuccess,
                                    {"path": filepath, "warnings": warns})

            ok, errs, warns = ext.save_in_place(md)
            if not ok:
                names = enum_flags("swFileSaveError_e", errs)
                return self._result(False,
                    f"Save failed (errors {errs}: {', '.join(names) or '?'}, warnings {warns})",
                    SwErrors.swFileSaveError)
            path = self._get_doc_path(doc)
            return self._result(True, f"Saved: {path}", SwErrors.swSuccess, {"path": path})

        except Exception as e:
            logger.error(f"Save error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swFileSaveError)

    def close_document(self, save: bool = False) -> Dict:
        """
        Close the active document
        
        Args:
            save: Save before closing
        
        Returns:
            Result dictionary
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return self._result(True, "No document to close")
            
            title = self._get_doc_title(doc)
            
            if save:
                from .. import ext
                ok, errs, _ = ext.save_in_place(ext.T(doc, "IModelDoc2"))
                if not ok:
                    return self._result(False, f"Save before close failed (errors {errs}); "
                                        f"document left open", SwErrors.swFileSaveError)

            v(self._sw_app, "CloseDoc", title)
            
            return self._result(True, f"Closed: {title}",
                              SwErrors.swSuccess, {"document": title})
            
        except Exception as e:
            logger.error(f"Close error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swUnknownError)
    
    def get_document_info(self) -> Dict:
        """
        Get information about the active document
        
        Returns:
            Result dictionary with document details
        """
        try:
            doc, err = self.get_active_doc()
            if err:
                return err
            
            type_names = {
                0: "None",
                1: "Part",
                2: "Assembly",
                3: "Drawing"
            }
            
            doc_type = com_get(doc, "GetType")
            title = self._get_doc_title(doc)
            path = self._get_doc_path(doc)
            
            info = {
                "title": title,
                "path": path if path else "Not saved",
                "type": type_names.get(doc_type, "Unknown"),
                "type_code": doc_type,
            }
            
            return self._result(True, f"{title} ({info['type']})",
                              SwErrors.swSuccess, info)
            
        except Exception as e:
            logger.error(f"Get info error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swUnknownError)
    
    def list_open_documents(self) -> Dict:
        """
        List all open documents
        
        Returns:
            Result dictionary with document list
        """
        try:
            if not self.is_connected:
                r = self.connect()
                if not r["success"]:
                    return r
            
            docs = []
            doc = com_get(self._sw_app, "GetFirstDocument")

            while doc:
                try:
                    title = com_get(doc, "GetTitle")
                    doc_type = com_get(doc, "GetType")

                    type_names = {1: "Part", 2: "Assembly", 3: "Drawing"}

                    docs.append({
                        "title": title,
                        "type": type_names.get(doc_type, "Unknown")
                    })
                except:
                    pass

                doc = com_get(doc, "GetNext")
            
            return self._result(True, f"{len(docs)} document(s) open",
                              SwErrors.swSuccess, {"documents": docs})
            
        except Exception as e:
            logger.error(f"List documents error: {e}\n{traceback.format_exc()}")
            return self._result(False, f"Error: {e}", SwErrors.swUnknownError)
