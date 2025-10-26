# Copyright (c) Meta Platforms, Inc. and affiliates.
# All rights reserved.
#
# This source code is licensed under the BSD-style license found in the
# LICENSE file in the root directory of this source tree.

"""
Data models for the LibreOffice Spreadsheet Environment.

The LibreOffice environment provides a spreadsheet interface for data manipulation,
formulas, and analysis using LibreOffice Calc.
"""

from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Union

from core.env_server.types import Action, Observation


@dataclass(kw_only=True)
class LibreOfficeAction(Action):
    """Action for the LibreOffice Spreadsheet environment."""

    command: str
    """The command to execute. Supported commands:
    - 'create_sheet': Create a new spreadsheet
    - 'open_file': Open an existing spreadsheet file
    - 'save_file': Save the current spreadsheet
    - 'set_cell': Set a cell value
    - 'get_cell': Get a cell value
    - 'set_range': Set values for a range of cells
    - 'get_range': Get values from a range of cells
    - 'set_formula': Set a formula in a cell
    - 'get_formula': Get a formula from a cell
    - 'add_sheet': Add a new sheet
    - 'delete_sheet': Delete a sheet
    - 'rename_sheet': Rename a sheet
    - 'copy_range': Copy a range of cells
    - 'paste_range': Paste a range of cells
    - 'format_cell': Format a cell (bold, italic, color, etc.)
    - 'sort_range': Sort a range of cells
    - 'filter_range': Apply filter to a range
    - 'create_chart': Create a chart from data
    - 'export_pdf': Export spreadsheet as PDF
    - 'export_csv': Export spreadsheet as CSV
    """

    parameters: Dict[str, Any] = None
    """Parameters for the command. Structure depends on the command:
    
    For 'set_cell':
        - sheet: str (sheet name, default: current sheet)
        - cell: str (cell address like 'A1')
        - value: Any (value to set)
    
    For 'get_cell':
        - sheet: str (sheet name, default: current sheet)
        - cell: str (cell address like 'A1')
    
    For 'set_range':
        - sheet: str (sheet name, default: current sheet)
        - range: str (range like 'A1:C3')
        - values: List[List[Any]] (2D array of values)
    
    For 'get_range':
        - sheet: str (sheet name, default: current sheet)
        - range: str (range like 'A1:C3')
    
    For 'set_formula':
        - sheet: str (sheet name, default: current sheet)
        - cell: str (cell address like 'A1')
        - formula: str (formula like '=SUM(A1:A10)')
    
    For 'open_file':
        - file_path: str (path to the file)
    
    For 'save_file':
        - file_path: str (path to save the file)
    
    For 'format_cell':
        - sheet: str (sheet name, default: current sheet)
        - cell: str (cell address like 'A1')
        - format_options: Dict[str, Any] (formatting options)
    
    For 'create_chart':
        - sheet: str (sheet name, default: current sheet)
        - data_range: str (range containing data)
        - chart_type: str (type of chart)
        - position: str (where to place the chart)
    """

    def __post_init__(self):
        if self.parameters is None:
            self.parameters = {}


@dataclass(kw_only=True)
class LibreOfficeObservation(Observation):
    """Observation from the LibreOffice Spreadsheet environment."""

    result: str
    """Result message describing what happened."""

    success: bool
    """Whether the operation was successful."""

    data: Optional[Union[str, List[List[Any]], Dict[str, Any]]] = None
    """Data returned by the operation. Type depends on the command:
    - For 'get_cell': str (cell value)
    - For 'get_range': List[List[Any]] (2D array of values)
    - For 'get_formula': str (formula string)
    - For other commands: Dict[str, Any] (additional data)
    """

    current_sheet: Optional[str] = None
    """Name of the currently active sheet."""

    sheet_names: Optional[List[str]] = None
    """List of all sheet names in the workbook."""

    error_message: Optional[str] = None
    """Error message if the operation failed."""

    file_path: Optional[str] = None
    """Path to the currently open file (if any)."""
