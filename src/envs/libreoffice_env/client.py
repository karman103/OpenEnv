# Copyright (c) Meta Platforms, Inc. and affiliates.
# All rights reserved.
#
# This source code is licensed under the BSD-style license found in the
# LICENSE file in the root directory of this source tree.

"""
LibreOffice Spreadsheet Environment HTTP Client.

This module provides the client for connecting to a LibreOffice Environment server
over HTTP.
"""

from typing import Any, Dict, List, Optional, Union

from core.client_types import StepResult
from core.env_server.types import State
from core.http_env_client import HTTPEnvClient

from .models import LibreOfficeAction, LibreOfficeObservation


class LibreOfficeEnv(HTTPEnvClient[LibreOfficeAction, LibreOfficeObservation]):
    """
    HTTP client for the LibreOffice Spreadsheet Environment.

    This client connects to a LibreOfficeEnvironment HTTP server and provides
    methods to interact with it: reset(), step(), and state access.

    Example:
        >>> # Connect to a running server
        >>> client = LibreOfficeEnv(base_url="http://localhost:8000")
        >>> result = client.reset()
        >>> print(result.observation.result)  # "LibreOffice environment ready!"
        >>>
        >>> # Set a cell value
        >>> action = LibreOfficeAction(
        ...     command="set_cell",
        ...     parameters={"cell": "A1", "value": "Hello World"}
        ... )
        >>> result = client.step(action)
        >>> print(result.observation.success)  # True
        >>>
        >>> # Get the cell value
        >>> action = LibreOfficeAction(
        ...     command="get_cell",
        ...     parameters={"cell": "A1"}
        ... )
        >>> result = client.step(action)
        >>> print(result.observation.data)  # "Hello World"

    Example with Docker:
        >>> # Automatically start container and connect
        >>> client = LibreOfficeEnv.from_docker_image("libreoffice-env:latest")
        >>> result = client.reset()
        >>> action = LibreOfficeAction(
        ...     command="set_cell",
        ...     parameters={"cell": "A1", "value": "Test"}
        ... )
        >>> result = client.step(action)
    """

    def _step_payload(self, action: LibreOfficeAction) -> Dict:
        """
        Convert LibreOfficeAction to JSON payload for step request.

        Args:
            action: LibreOfficeAction instance

        Returns:
            Dictionary representation suitable for JSON encoding
        """
        return {
            "command": action.command,
            "parameters": action.parameters or {},
        }

    def _parse_result(self, payload: Dict) -> StepResult[LibreOfficeObservation]:
        """
        Parse server response into StepResult[LibreOfficeObservation].

        Args:
            payload: JSON response from server

        Returns:
            StepResult with LibreOfficeObservation
        """
        obs_data = payload.get("observation", {})
        observation = LibreOfficeObservation(
            result=obs_data.get("result", ""),
            success=obs_data.get("success", False),
            data=obs_data.get("data"),
            current_sheet=obs_data.get("current_sheet"),
            sheet_names=obs_data.get("sheet_names"),
            error_message=obs_data.get("error_message"),
            file_path=obs_data.get("file_path"),
            done=payload.get("done", False),
            reward=payload.get("reward"),
            metadata=obs_data.get("metadata", {}),
        )

        return StepResult(
            observation=observation,
            reward=payload.get("reward"),
            done=payload.get("done", False),
        )

    def _parse_state(self, payload: Dict) -> State:
        """
        Parse server response into State object.

        Args:
            payload: JSON response from /state endpoint

        Returns:
            State object with episode_id and step_count
        """
        return State(
            episode_id=payload.get("episode_id"),
            step_count=payload.get("step_count", 0),
        )

    # Convenience methods for common operations
    def set_cell(self, cell: str, value: Any, sheet: Optional[str] = None) -> StepResult[LibreOfficeObservation]:
        """
        Set a cell value.

        Args:
            cell: Cell address (e.g., "A1")
            value: Value to set
            sheet: Sheet name (optional, uses current sheet if not provided)

        Returns:
            StepResult with the operation result
        """
        params = {"cell": cell, "value": value}
        if sheet:
            params["sheet"] = sheet
        
        action = LibreOfficeAction(command="set_cell", parameters=params)
        return self.step(action)

    def get_cell(self, cell: str, sheet: Optional[str] = None) -> StepResult[LibreOfficeObservation]:
        """
        Get a cell value.

        Args:
            cell: Cell address (e.g., "A1")
            sheet: Sheet name (optional, uses current sheet if not provided)

        Returns:
            StepResult with the cell value in observation.data
        """
        params = {"cell": cell}
        if sheet:
            params["sheet"] = sheet
        
        action = LibreOfficeAction(command="get_cell", parameters=params)
        return self.step(action)

    def set_range(self, range_name: str, values: List[List[Any]], sheet: Optional[str] = None) -> StepResult[LibreOfficeObservation]:
        """
        Set values for a range of cells.

        Args:
            range_name: Range address (e.g., "A1:C3")
            values: 2D array of values
            sheet: Sheet name (optional, uses current sheet if not provided)

        Returns:
            StepResult with the operation result
        """
        params = {"range": range_name, "values": values}
        if sheet:
            params["sheet"] = sheet
        
        action = LibreOfficeAction(command="set_range", parameters=params)
        return self.step(action)

    def get_range(self, range_name: str, sheet: Optional[str] = None) -> StepResult[LibreOfficeObservation]:
        """
        Get values from a range of cells.

        Args:
            range_name: Range address (e.g., "A1:C3")
            sheet: Sheet name (optional, uses current sheet if not provided)

        Returns:
            StepResult with the range data in observation.data
        """
        params = {"range": range_name}
        if sheet:
            params["sheet"] = sheet
        
        action = LibreOfficeAction(command="get_range", parameters=params)
        return self.step(action)

    def set_formula(self, cell: str, formula: str, sheet: Optional[str] = None) -> StepResult[LibreOfficeObservation]:
        """
        Set a formula in a cell.

        Args:
            cell: Cell address (e.g., "A1")
            formula: Formula string (e.g., "=SUM(A1:A10)")
            sheet: Sheet name (optional, uses current sheet if not provided)

        Returns:
            StepResult with the operation result
        """
        params = {"cell": cell, "formula": formula}
        if sheet:
            params["sheet"] = sheet
        
        action = LibreOfficeAction(command="set_formula", parameters=params)
        return self.step(action)

    def get_formula(self, cell: str, sheet: Optional[str] = None) -> StepResult[LibreOfficeObservation]:
        """
        Get a formula from a cell.

        Args:
            cell: Cell address (e.g., "A1")
            sheet: Sheet name (optional, uses current sheet if not provided)

        Returns:
            StepResult with the formula in observation.data
        """
        params = {"cell": cell}
        if sheet:
            params["sheet"] = sheet
        
        action = LibreOfficeAction(command="get_formula", parameters=params)
        return self.step(action)

    def open_file(self, file_path: str) -> StepResult[LibreOfficeObservation]:
        """
        Open an existing spreadsheet file.

        Args:
            file_path: Path to the file to open

        Returns:
            StepResult with the operation result
        """
        action = LibreOfficeAction(command="open_file", parameters={"file_path": file_path})
        return self.step(action)

    def save_file(self, file_path: str) -> StepResult[LibreOfficeObservation]:
        """
        Save the current spreadsheet.

        Args:
            file_path: Path where to save the file

        Returns:
            StepResult with the operation result
        """
        action = LibreOfficeAction(command="save_file", parameters={"file_path": file_path})
        return self.step(action)

    def add_sheet(self, name: Optional[str] = None) -> StepResult[LibreOfficeObservation]:
        """
        Add a new sheet.

        Args:
            name: Name for the new sheet (optional, auto-generated if not provided)

        Returns:
            StepResult with the operation result
        """
        params = {}
        if name:
            params["name"] = name
        
        action = LibreOfficeAction(command="add_sheet", parameters=params)
        return self.step(action)

    def delete_sheet(self, name: str) -> StepResult[LibreOfficeObservation]:
        """
        Delete a sheet.

        Args:
            name: Name of the sheet to delete

        Returns:
            StepResult with the operation result
        """
        action = LibreOfficeAction(command="delete_sheet", parameters={"name": name})
        return self.step(action)

    def rename_sheet(self, old_name: str, new_name: str) -> StepResult[LibreOfficeObservation]:
        """
        Rename a sheet.

        Args:
            old_name: Current name of the sheet
            new_name: New name for the sheet

        Returns:
            StepResult with the operation result
        """
        action = LibreOfficeAction(command="rename_sheet", parameters={"old_name": old_name, "new_name": new_name})
        return self.step(action)

    def format_cell(self, cell: str, format_options: Dict[str, Any], sheet: Optional[str] = None) -> StepResult[LibreOfficeObservation]:
        """
        Format a cell.

        Args:
            cell: Cell address (e.g., "A1")
            format_options: Formatting options (e.g., {"bold": True, "color": "#FF0000"})
            sheet: Sheet name (optional, uses current sheet if not provided)

        Returns:
            StepResult with the operation result
        """
        params = {"cell": cell, "format_options": format_options}
        if sheet:
            params["sheet"] = sheet
        
        action = LibreOfficeAction(command="format_cell", parameters=params)
        return self.step(action)

    def export_pdf(self, file_path: str) -> StepResult[LibreOfficeObservation]:
        """
        Export spreadsheet as PDF.

        Args:
            file_path: Path where to save the PDF

        Returns:
            StepResult with the operation result
        """
        action = LibreOfficeAction(command="export_pdf", parameters={"file_path": file_path})
        return self.step(action)

    def export_csv(self, file_path: str) -> StepResult[LibreOfficeObservation]:
        """
        Export spreadsheet as CSV.

        Args:
            file_path: Path where to save the CSV

        Returns:
            StepResult with the operation result
        """
        action = LibreOfficeAction(command="export_csv", parameters={"file_path": file_path})
        return self.step(action)
