# Copyright (c) Meta Platforms, Inc. and affiliates.
# All rights reserved.
#
# This source code is licensed under the BSD-style license found in the
# LICENSE file in the root directory of this source tree.

"""
LibreOffice Spreadsheet Environment Implementation.

This environment provides a spreadsheet interface for data manipulation,
formulas, and analysis using LibreOffice Calc through Python-UNO bridge.
"""

import os
import tempfile
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from core.env_server.interfaces import Environment
from core.env_server.types import State

from ..models import LibreOfficeAction, LibreOfficeObservation


class LibreOfficeEnvironment(Environment):
    """
    A LibreOffice spreadsheet environment for data manipulation and analysis.

    This environment provides a comprehensive interface to LibreOffice Calc
    through the Python-UNO bridge, allowing for spreadsheet operations like
    setting/getting cell values, formulas, formatting, charts, and more.

    Example:
        >>> env = LibreOfficeEnvironment()
        >>> obs = env.reset()
        >>> print(obs.result)  # "LibreOffice environment ready!"
        >>>
        >>> # Set a cell value
        >>> action = LibreOfficeAction(
        ...     command="set_cell",
        ...     parameters={"cell": "A1", "value": "Hello World"}
        ... )
        >>> obs = env.step(action)
        >>> print(obs.success)  # True
        >>>
        >>> # Get the cell value
        >>> action = LibreOfficeAction(
        ...     command="get_cell",
        ...     parameters={"cell": "A1"}
        ... )
        >>> obs = env.step(action)
        >>> print(obs.data)  # "Hello World"
    """

    def __init__(self, base_excel: Optional[str] = None, goal_excel: Optional[str] = None):
        """Initialize LibreOffice environment with optional base and goal Excel files."""
        self._state = State(episode_id=str(uuid.uuid4()), step_count=0)
        self._reset_count = 0
        self._libreoffice_context = None
        self._document = None
        self._current_sheet = None
        self._file_path = None
        self._temp_dir = None
        self._initialized = False

        # New
        self._base_excel = base_excel
        self._goal_excel = goal_excel


    def _initialize_libreoffice(self) -> bool:
        """Initialize LibreOffice connection (robust socket-based)."""
        print("🔧 Initializing LibreOffice environment...")
        import time
        import uno
        from com.sun.star.connection import NoConnectException
        from com.sun.star.uno import Exception as UnoException

        try:
            # Connection details
            connection_url = "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
            print(f"🔌 Attempting to connect to LibreOffice UNO service at {connection_url}")

            # Get local component context 
            print("⏳ Waiting for LibreOffice to be ready...")
            local_ctx = uno.getComponentContext()
            print("✅ Obtained local UNO component context.")
            resolver = local_ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_ctx
            )

            # Try connecting multiple times (LibreOffice may take a few seconds to start)
            for attempt in range(10):
                try:
                    self._libreoffice_context = resolver.resolve(connection_url)
                    print("✅ Connected to LibreOffice UNO service.")
                    break
                except NoConnectException:
                    print(f"⚠️  LibreOffice not ready (try {attempt + 1}/10)...")
                    time.sleep(1.5)
                except UnoException as e:
                    print(f"⚠️  UNO Exception: {e}")
                    time.sleep(1.5)
            else:
                raise RuntimeError("Could not connect to LibreOffice UNO socket.")

            # Get the service manager
            service_manager = self._libreoffice_context.getServiceManager()

            # Create a desktop service
            desktop = service_manager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self._libreoffice_context
            )

            # Create a new spreadsheet document
            document_url = "private:factory/scalc"
            self._document = desktop.loadComponentFromURL(document_url, "_blank", 0, ())

            if self._document is None:
                print("❌ Failed to create spreadsheet document.")
                return False

            # Get the first sheet
            sheets = self._document.getSheets()
            self._current_sheet = sheets.getByIndex(0)

            # Create temporary directory for file operations
            self._temp_dir = tempfile.mkdtemp(prefix="libreoffice_env_")

            self._initialized = True
            print("✅ LibreOffice environment initialized successfully.")
            if self._base_excel and os.path.exists(self._base_excel):
                print(f"📂 Loading base Excel file: {self._base_excel}")
                open_obs = self._open_file({"file_path": self._base_excel})
                if not open_obs.success:
                    print(f"⚠️ Warning: Failed to open base Excel file: {open_obs.error_message}")
            else:
                print("ℹ️ No base Excel file provided, starting with a new sheet.")

            return True

        except Exception as e:
            print(f"❌ Failed to initialize LibreOffice: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _cleanup_libreoffice(self):
        """Clean up LibreOffice resources."""
        try:
            if self._document:
                self._document.close(True)
                self._document = None
            
            if self._temp_dir and os.path.exists(self._temp_dir):
                import shutil
                shutil.rmtree(self._temp_dir)
                self._temp_dir = None
                
        except Exception as e:
            print(f"Error during cleanup: {e}")

    def _get_sheet(self, sheet_name: Optional[str] = None):
        """Get a sheet by name or return current sheet."""
        if sheet_name is None:
            return self._current_sheet
        
        sheets = self._document.getSheets()
        try:
            return sheets.getByName(sheet_name)
        except Exception:
            return self._current_sheet

    def _execute_command(self, action: LibreOfficeAction) -> LibreOfficeObservation:
        """Execute a LibreOffice command."""
        if not self._initialized:
            return LibreOfficeObservation(
                result="LibreOffice not initialized",
                success=False,
                error_message="LibreOffice connection failed"
            )

        try:
            command = action.command
            params = action.parameters or {}

            if command == "create_sheet":
                return self._create_sheet(params)
            elif command == "open_file":
                return self._open_file(params)
            elif command == "save_file":
                return self._save_file(params)
            elif command == "set_cell":
                return self._set_cell(params)
            elif command == "get_cell":
                return self._get_cell(params)
            elif command == "set_range":
                return self._set_range(params)
            elif command == "get_range":
                return self._get_range(params)
            elif command == "set_formula":
                return self._set_formula(params)
            elif command == "get_formula":
                return self._get_formula(params)
            elif command == "add_sheet":
                return self._add_sheet(params)
            elif command == "delete_sheet":
                return self._delete_sheet(params)
            elif command == "rename_sheet":
                return self._rename_sheet(params)
            elif command == "format_cell":
                return self._format_cell(params)
            elif command == "export_pdf":
                return self._export_pdf(params)
            elif command == "export_csv":
                return self._export_csv(params)
            else:
                return LibreOfficeObservation(
                    result=f"Unknown command: {command}",
                    success=False,
                    error_message=f"Command '{command}' not supported"
                )

        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error executing command: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _create_sheet(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Create a new spreadsheet (already done in initialization)."""
        return LibreOfficeObservation(
            result="New spreadsheet created",
            success=True,
            current_sheet=self._current_sheet.getName(),
            sheet_names=[self._current_sheet.getName()]
        )

    def _open_file(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Open an existing spreadsheet file."""
        file_path = params.get("file_path")
        if not file_path:
            return LibreOfficeObservation(
                result="No file path provided",
                success=False,
                error_message="file_path parameter is required"
            )

        try:
            from unohelper import systemPathToFileUrl
            
            # Convert to file URL
            file_url = systemPathToFileUrl(os.path.abspath(file_path))
            
            # Close current document
            if self._document:
                self._document.close(True)
            
            # Open new document
            desktop = self._libreoffice_context.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self._libreoffice_context
            )
            self._document = desktop.loadComponentFromURL(file_url, "_blank", 0, ())
            
            if self._document is None:
                return LibreOfficeObservation(
                    result="Failed to open file",
                    success=False,
                    error_message="Could not open the specified file"
                )
            
            # Get the first sheet
            sheets = self._document.getSheets()
            self._current_sheet = sheets.getByIndex(0)
            self._file_path = file_path
            
            sheet_names = [sheets.getByIndex(i).getName() for i in range(sheets.getCount())]
            
            return LibreOfficeObservation(
                result=f"File opened successfully: {file_path}",
                success=True,
                current_sheet=self._current_sheet.getName(),
                sheet_names=sheet_names,
                file_path=file_path
            )
            
        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error opening file: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _save_file(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Save the current spreadsheet."""
        file_path = params.get("file_path")
        if not file_path:
            return LibreOfficeObservation(
                result="No file path provided",
                success=False,
                error_message="file_path parameter is required"
            )

        try:
            from unohelper import systemPathToFileUrl
            
            # Convert to file URL
            file_url = systemPathToFileUrl(os.path.abspath(file_path))
            
            # Save the document
            self._document.storeAsURL(file_url, ())
            self._file_path = file_path
            
            return LibreOfficeObservation(
                result=f"File saved successfully: {file_path}",
                success=True,
                file_path=file_path
            )
            
        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error saving file: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _set_cell(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Set a cell value."""
        cell = params.get("cell")
        value = params.get("value")
        sheet_name = params.get("sheet")

        if not cell:
            return LibreOfficeObservation(
                result="No cell specified",
                success=False,
                error_message="cell parameter is required"
            )

        try:
            sheet = self._get_sheet(sheet_name)
            cell_range = sheet.getCellRangeByName(cell)

            # Use correct setter depending on type
            if isinstance(value, (int, float)):
                cell_range.setValue(float(value))
            else:
                cell_range.setString(str(value))

            return LibreOfficeObservation(
                result=f"Cell {cell} set to {value}",
                success=True,
                current_sheet=sheet.getName()
            )

        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error setting cell: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _get_cell(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Get a cell value."""
        cell = params.get("cell")
        sheet_name = params.get("sheet")

        if not cell:
            return LibreOfficeObservation(
                result="No cell specified",
                success=False,
                error_message="cell parameter is required"
            )

        try:
            sheet = self._get_sheet(sheet_name)
            cell_range = sheet.getCellRangeByName(cell)

            # Try to read both string and numeric content
            text = cell_range.getString()
            if text.strip() != "":
                value = text
            else:
                value = cell_range.getValue()

            return LibreOfficeObservation(
                result=f"Retrieved value from cell {cell}",
                success=True,
                data=value,
                current_sheet=sheet.getName()
            )

        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error getting cell: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _set_range(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Set values for a range of cells (supports text and numbers)."""
        range_name = params.get("range")
        values = params.get("values")
        sheet_name = params.get("sheet")

        if not range_name or not values:
            return LibreOfficeObservation(
                result="Range and values required",
                success=False,
                error_message="range and values parameters are required"
            )

        try:
            sheet = self._get_sheet(sheet_name)
            base_range = sheet.getCellRangeByName(range_name)
            start_col = base_range.RangeAddress.StartColumn
            start_row = base_range.RangeAddress.StartRow

            # Set each cell properly
            for i, row in enumerate(values):
                for j, value in enumerate(row):
                    cell = sheet.getCellByPosition(start_col + j, start_row + i)
                    if isinstance(value, (int, float)):
                        cell.setValue(float(value))
                    else:
                        cell.setString(str(value))

            return LibreOfficeObservation(
                result=f"Range {range_name} set successfully.",
                success=True,
                current_sheet=sheet.getName()
            )

        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error setting range: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _get_range(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Get values from a range of cells."""
        range_name = params.get("range")
        sheet_name = params.get("sheet")
        
        if not range_name:
            return LibreOfficeObservation(
                result="No range specified",
                success=False,
                error_message="range parameter is required"
            )

        try:
            sheet = self._get_sheet(sheet_name)
            cell_range = sheet.getCellRangeByName(range_name)
            
            # Get the data array
            data_array = cell_range.getDataArray()
            
            return LibreOfficeObservation(
                result=f"Retrieved data from range {range_name}",
                success=True,
                data=data_array,
                current_sheet=sheet.getName()
            )
            
        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error getting range: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _set_formula(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Set a formula in a cell."""
        cell = params.get("cell")
        formula = params.get("formula")
        sheet_name = params.get("sheet")
        
        if not cell or not formula:
            return LibreOfficeObservation(
                result="Cell and formula required",
                success=False,
                error_message="cell and formula parameters are required"
            )

        try:
            sheet = self._get_sheet(sheet_name)
            cell_range = sheet.getCellRangeByName(cell)
            cell_range.setFormula(formula)
            
            return LibreOfficeObservation(
                result=f"Formula set in cell {cell}: {formula}",
                success=True,
                current_sheet=sheet.getName()
            )
            
        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error setting formula: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _get_formula(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Get a formula from a cell."""
        cell = params.get("cell")
        sheet_name = params.get("sheet")
        
        if not cell:
            return LibreOfficeObservation(
                result="No cell specified",
                success=False,
                error_message="cell parameter is required"
            )

        try:
            sheet = self._get_sheet(sheet_name)
            cell_range = sheet.getCellRangeByName(cell)
            formula = cell_range.getFormula()
            
            return LibreOfficeObservation(
                result=f"Retrieved formula from cell {cell}",
                success=True,
                data=formula,
                current_sheet=sheet.getName()
            )
            
        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error getting formula: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _add_sheet(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Add a new sheet."""
        sheet_name = params.get("name", f"Sheet{self._document.getSheets().getCount() + 1}")
        
        try:
            sheets = self._document.getSheets()
            new_sheet = sheets.insertNewByName(sheet_name, sheets.getCount())
            
            sheet_names = [sheets.getByIndex(i).getName() for i in range(sheets.getCount())]
            
            return LibreOfficeObservation(
                result=f"Sheet '{sheet_name}' added successfully",
                success=True,
                current_sheet=sheet_name,
                sheet_names=sheet_names
            )
            
        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error adding sheet: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _delete_sheet(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Delete a sheet."""
        sheet_name = params.get("name")
        
        if not sheet_name:
            return LibreOfficeObservation(
                result="No sheet name specified",
                success=False,
                error_message="name parameter is required"
            )

        try:
            sheets = self._document.getSheets()
            sheet = sheets.getByName(sheet_name)
            sheets.remove(sheet)
            
            # Update current sheet if it was deleted
            if self._current_sheet.getName() == sheet_name:
                self._current_sheet = sheets.getByIndex(0)
            
            sheet_names = [sheets.getByIndex(i).getName() for i in range(sheets.getCount())]
            
            return LibreOfficeObservation(
                result=f"Sheet '{sheet_name}' deleted successfully",
                success=True,
                current_sheet=self._current_sheet.getName(),
                sheet_names=sheet_names
            )
            
        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error deleting sheet: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _rename_sheet(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Rename a sheet."""
        old_name = params.get("old_name")
        new_name = params.get("new_name")
        
        if not old_name or not new_name:
            return LibreOfficeObservation(
                result="Old and new names required",
                success=False,
                error_message="old_name and new_name parameters are required"
            )

        try:
            sheets = self._document.getSheets()
            sheet = sheets.getByName(old_name)
            sheet.setName(new_name)
            
            # Update current sheet if it was renamed
            if self._current_sheet.getName() == old_name:
                self._current_sheet = sheet
            
            sheet_names = [sheets.getByIndex(i).getName() for i in range(sheets.getCount())]
            
            return LibreOfficeObservation(
                result=f"Sheet renamed from '{old_name}' to '{new_name}'",
                success=True,
                current_sheet=new_name,
                sheet_names=sheet_names
            )
            
        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error renaming sheet: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _format_cell(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Format a cell."""
        cell = params.get("cell")
        format_options = params.get("format_options", {})
        sheet_name = params.get("sheet")
        
        if not cell:
            return LibreOfficeObservation(
                result="No cell specified",
                success=False,
                error_message="cell parameter is required"
            )

        try:
            sheet = self._get_sheet(sheet_name)
            cell_range = sheet.getCellRangeByName(cell)
            
            # Apply formatting options
            if "bold" in format_options:
                cell_range.setPropertyValue("CharWeight", 150 if format_options["bold"] else 100)
            if "italic" in format_options:
                cell_range.setPropertyValue("CharPosture", 1 if format_options["italic"] else 0)
            if "color" in format_options:
                # Convert color to LibreOffice color format
                color = format_options["color"]
                if isinstance(color, str) and color.startswith("#"):
                    color = int(color[1:], 16)
                cell_range.setPropertyValue("CharColor", color)
            
            return LibreOfficeObservation(
                result=f"Cell {cell} formatted successfully",
                success=True,
                current_sheet=sheet.getName()
            )
            
        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error formatting cell: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _export_pdf(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Export spreadsheet as PDF (works in headless mode)."""
        file_path = params.get("file_path")
        if not file_path:
            return LibreOfficeObservation(
                result="No file path provided",
                success=False,
                error_message="file_path parameter is required"
            )

        try:
            from unohelper import systemPathToFileUrl
            from com.sun.star.beans import PropertyValue

            # Ensure directory exists and is writable
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            if not os.access(os.path.dirname(file_path), os.W_OK):
                return LibreOfficeObservation(
                    result=f"Directory not writable: {os.path.dirname(file_path)}",
                    success=False,
                    error_message="Target directory is not writable"
                )

            # Convert to UNO URL format
            file_url = systemPathToFileUrl(os.path.abspath(file_path))

            # Proper UNO property list for PDF export
            filter_props = []
            p1 = PropertyValue()
            p1.Name = "FilterName"
            p1.Value = "calc_pdf_Export"
            filter_props.append(p1)

            p2 = PropertyValue()
            p2.Name = "Overwrite"
            p2.Value = True
            filter_props.append(p2)

            # Perform export
            self._document.storeToURL(file_url, tuple(filter_props))

            # Verify export
            exists = os.path.exists(file_path)
            return LibreOfficeObservation(
                result="PDF export completed" if exists else "Export command issued but file not found",
                success=exists,
                data={"exported_file": file_path},
            )

        except Exception as e:
            import traceback
            traceback.print_exc()
            return LibreOfficeObservation(
                result=f"Error exporting PDF: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def _export_csv(self, params: Dict[str, Any]) -> LibreOfficeObservation:
        """Export spreadsheet as CSV."""
        file_path = params.get("file_path")
        sheet_name = params.get("sheet")
        
        if not file_path:
            return LibreOfficeObservation(
                result="No file path provided",
                success=False,
                error_message="file_path parameter is required"
            )

        try:
            from unohelper import systemPathToFileUrl
            
            # Convert to file URL
            file_url = systemPathToFileUrl(os.path.abspath(file_path))
            
            # Export as CSV
            filter_props = (
                ("FilterName", "Text - txt - csv (StarCalc)"),
                ("FilterOptions", "44,34,0,1,1"),
            )
            self._document.storeToURL(file_url, filter_props)
            
            return LibreOfficeObservation(
                result=f"CSV exported successfully: {file_path}",
                success=True,
                data={"exported_file": file_path}
            )
            
        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error exporting CSV: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def reset(self) -> LibreOfficeObservation:
        """
        Reset the environment.

        Returns:
            LibreOfficeObservation with initialization status
        """
        # Clean up previous session
        print("♻️  Resetting LibreOffice environment...")
        self._cleanup_libreoffice()
        
        # Reset state
        self._state = State(episode_id=str(uuid.uuid4()), step_count=0)
        self._reset_count += 1
        
        # Initialize LibreOffice
        if self._initialize_libreoffice():
            sheet_names = [self._current_sheet.getName()] if self._current_sheet else []
            return LibreOfficeObservation(
                result="LibreOffice environment ready!",
                success=True,
                current_sheet=self._current_sheet.getName() if self._current_sheet else None,
                sheet_names=sheet_names,
                done=False,
                reward=0.0,
            )
        else:
            return LibreOfficeObservation(
                result="Failed to initialize LibreOffice",
                success=False,
                error_message="Could not connect to LibreOffice",
                done=False,
                reward=0.0,
            )

    def step(self, action: LibreOfficeAction) -> LibreOfficeObservation:  # type: ignore[override]
        """
        Execute a step in the environment.

        Args:
            action: LibreOfficeAction containing the command and parameters

        Returns:
            LibreOfficeObservation with the result of the operation
        """
        self._state.step_count += 1

        # Execute the command
        observation = self._execute_command(action)
        
        # Add metadata
        observation.metadata = {
            "step": self._state.step_count,
            "command": action.command,
            "episode_id": self._state.episode_id,
        }
        
        # Simple reward: successful operations get positive reward
        observation.reward = 1.0 if observation.success else -0.1
        observation.done = False

        return observation

    @property
    def state(self) -> State:
        """
        Get the current environment state.

        Returns:
            Current State with episode_id and step_count
        """
        return self._state
    def close(self) -> LibreOfficeObservation:
        """
        Close the LibreOffice environment gracefully.
        Shuts down any open documents and frees resources.
        """
        print("🧹 Closing LibreOffice environment...")
        try:
            if self._document is not None:
                try:
                    self._document.close(True)
                    self._document = None
                    print("✅ Document closed.")
                except Exception as e:
                    print(f"⚠️ Error closing document: {e}")

            if self._libreoffice_context is not None:
                try:
                    # Attempt to terminate LibreOffice via the desktop interface
                    desktop = self._libreoffice_context.getServiceManager().createInstanceWithContext(
                        "com.sun.star.frame.Desktop", self._libreoffice_context
                    )
                    if desktop:
                        desktop.terminate()
                        print("✅ LibreOffice desktop terminated.")
                except Exception as e:
                    print(f"⚠️ Error terminating LibreOffice: {e}")
            # Before cleanup, save the document if a goal file path was provided
            if self._goal_excel:
                print(f"💾 Saving final Excel output to: {self._goal_excel}")
                try:
                    save_obs = self._save_file({"file_path": self._goal_excel})
                    if save_obs.success:
                        print("✅ Final Excel file saved successfully.")
                    else:
                        print(f"⚠️ Could not save goal Excel file: {save_obs.error_message}")
                except Exception as e:
                    print(f"⚠️ Error during goal Excel save: {e}")


            # Clean up temp files and state
            self._cleanup_libreoffice()
            self._initialized = False
            

            return LibreOfficeObservation(
                result="LibreOffice environment closed successfully.",
                success=True
            )

        except Exception as e:
            return LibreOfficeObservation(
                result=f"Error during close: {str(e)}",
                success=False,
                error_message=str(e)
            )

    def __del__(self):
        """Cleanup when the environment is destroyed."""
        self._cleanup_libreoffice()
