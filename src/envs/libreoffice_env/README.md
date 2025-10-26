# LibreOffice Spreadsheet Environment

A comprehensive LibreOffice Calc environment for spreadsheet operations, data manipulation, and analysis through the OpenEnv framework.

## Overview

The LibreOffice Spreadsheet Environment provides a programmatic interface to LibreOffice Calc through the Python-UNO bridge. It enables automated spreadsheet operations including:

- Cell and range operations (set/get values, formulas)
- Sheet management (create, delete, rename sheets)
- File operations (open, save, export to PDF/CSV)
- Cell formatting (bold, italic, colors)
- Data manipulation and analysis

## Features

### Core Operations
- **Cell Operations**: Set and get individual cell values
- **Range Operations**: Work with ranges of cells (A1:C3)
- **Formula Support**: Set and retrieve formulas with full Calc formula syntax
- **Sheet Management**: Create, delete, and rename sheets
- **File Operations**: Open existing files, save workbooks, export to various formats

### Advanced Features
- **Formatting**: Apply text formatting (bold, italic, colors)
- **Export Options**: Export spreadsheets as PDF or CSV
- **Multi-sheet Support**: Work with multiple sheets in a workbook
- **Error Handling**: Comprehensive error reporting and validation

## Quick Start

### Using the Environment

```python
from envs.libreoffice_env import LibreOfficeAction, LibreOfficeEnv

# Create environment from Docker image
client = LibreOfficeEnv.from_docker_image("libreoffice-env:latest")

# Reset the environment
result = client.reset()
print(result.observation.result)  # "LibreOffice environment ready!"

# Set a cell value
result = client.set_cell("A1", "Hello World")
print(result.observation.success)  # True

# Get the cell value
result = client.get_cell("A1")
print(result.observation.data)  # "Hello World"

# Set a formula
result = client.set_formula("B1", "=A1 & \" from LibreOffice\"")
print(result.observation.success)  # True

# Get the calculated result
result = client.get_cell("B1")
print(result.observation.data)  # "Hello World from LibreOffice"

# Set a range of values
data = [
    ["Name", "Age", "City"],
    ["Alice", 25, "New York"],
    ["Bob", 30, "London"],
    ["Charlie", 35, "Tokyo"]
]
result = client.set_range("A1:C4", data)
print(result.observation.success)  # True

# Get the range data
result = client.get_range("A1:C4")
print(result.observation.data)  # [["Name", "Age", "City"], ...]

# Format a cell
result = client.format_cell("A1", {"bold": True, "color": "#FF0000"})
print(result.observation.success)  # True

# Add a new sheet
result = client.add_sheet("Data Analysis")
print(result.observation.success)  # True

# Save the workbook
result = client.save_file("/tmp/my_spreadsheet.ods")
print(result.observation.success)  # True

# Export as PDF
result = client.export_pdf("/tmp/my_spreadsheet.pdf")
print(result.observation.success)  # True

# Cleanup
client.close()
```

### Using Raw Actions

```python
from envs.libreoffice_env import LibreOfficeAction, LibreOfficeEnv

client = LibreOfficeEnv.from_docker_image("libreoffice-env:latest")
client.reset()

# Set cell using raw action
action = LibreOfficeAction(
    command="set_cell",
    parameters={"cell": "A1", "value": 42}
)
result = client.step(action)
print(result.observation.success)  # True

# Set formula using raw action
action = LibreOfficeAction(
    command="set_formula",
    parameters={"cell": "B1", "formula": "=A1*2"}
)
result = client.step(action)
print(result.observation.success)  # True

# Get cell value using raw action
action = LibreOfficeAction(
    command="get_cell",
    parameters={"cell": "B1"}
)
result = client.step(action)
print(result.observation.data)  # "84"

client.close()
```

## Supported Commands

### Cell Operations
- `set_cell`: Set a cell value
- `get_cell`: Get a cell value
- `set_formula`: Set a formula in a cell
- `get_formula`: Get a formula from a cell

### Range Operations
- `set_range`: Set values for a range of cells
- `get_range`: Get values from a range of cells

### Sheet Management
- `create_sheet`: Create a new spreadsheet (done automatically on reset)
- `add_sheet`: Add a new sheet to the workbook
- `delete_sheet`: Delete a sheet
- `rename_sheet`: Rename a sheet

### File Operations
- `open_file`: Open an existing spreadsheet file
- `save_file`: Save the current spreadsheet
- `export_pdf`: Export spreadsheet as PDF
- `export_csv`: Export spreadsheet as CSV

### Formatting
- `format_cell`: Format a cell (bold, italic, color, etc.)

## Action Parameters

### set_cell
```python
{
    "cell": "A1",           # Cell address
    "value": "Hello",       # Value to set
    "sheet": "Sheet1"       # Optional: sheet name
}
```

### get_cell
```python
{
    "cell": "A1",           # Cell address
    "sheet": "Sheet1"       # Optional: sheet name
}
```

### set_range
```python
{
    "range": "A1:C3",       # Range address
    "values": [             # 2D array of values
        ["A", "B", "C"],
        [1, 2, 3],
        [4, 5, 6]
    ],
    "sheet": "Sheet1"       # Optional: sheet name
}
```

### set_formula
```python
{
    "cell": "A1",           # Cell address
    "formula": "=SUM(B1:B10)",  # Formula string
    "sheet": "Sheet1"       # Optional: sheet name
}
```

### format_cell
```python
{
    "cell": "A1",           # Cell address
    "format_options": {     # Formatting options
        "bold": True,
        "italic": False,
        "color": "#FF0000"
    },
    "sheet": "Sheet1"       # Optional: sheet name
}
```

### open_file
```python
{
    "file_path": "/path/to/file.ods"  # Path to file
}
```

### save_file
```python
{
    "file_path": "/path/to/save.ods"  # Path to save
}
```

### add_sheet
```python
{
    "name": "New Sheet"     # Optional: sheet name
}
```

### delete_sheet
```python
{
    "name": "Sheet to Delete"  # Sheet name
}
```

### rename_sheet
```python
{
    "old_name": "Old Name",    # Current name
    "new_name": "New Name"     # New name
}
```

## Observation Data

The `LibreOfficeObservation` contains:

- `result`: String describing what happened
- `success`: Boolean indicating if the operation succeeded
- `data`: Operation-specific data (cell values, ranges, etc.)
- `current_sheet`: Name of the currently active sheet
- `sheet_names`: List of all sheet names in the workbook
- `error_message`: Error message if the operation failed
- `file_path`: Path to the currently open file (if any)

## Building and Running

### Build Docker Image

```bash
# Build the base image first
docker build -t openenv-base:latest -f src/core/containers/images/Dockerfile .

# Build the LibreOffice environment image
docker build -t libreoffice-env:latest -f src/envs/libreoffice_env/server/Dockerfile .
```

### Run Locally

```bash
# Start the server
uvicorn envs.libreoffice_env.server.app:app --host 0.0.0.0 --port 8000

# Or run directly
python -m envs.libreoffice_env.server.app
```

### Using with Docker

```python
# Automatically start container and connect
client = LibreOfficeEnv.from_docker_image("libreoffice-env:latest")
```

## Examples

### Data Analysis Workflow

```python
from envs.libreoffice_env import LibreOfficeEnv

client = LibreOfficeEnv.from_docker_image("libreoffice-env:latest")
client.reset()

# Create a dataset
data = [
    ["Product", "Sales", "Price", "Revenue"],
    ["Widget A", 100, 10.50, "=B2*C2"],
    ["Widget B", 150, 15.75, "=B3*C3"],
    ["Widget C", 200, 8.25, "=B4*C4"],
    ["Widget D", 75, 22.00, "=B5*C5"]
]

# Set the data
client.set_range("A1:D5", data)

# Add totals
client.set_cell("A6", "TOTAL")
client.set_formula("B6", "=SUM(B2:B5)")
client.set_formula("C6", "=AVERAGE(C2:C5)")
client.set_formula("D6", "=SUM(D2:D5)")

# Format headers
client.format_cell("A1", {"bold": True})
client.format_cell("B1", {"bold": True})
client.format_cell("C1", {"bold": True})
client.format_cell("D1", {"bold": True})

# Format totals
client.format_cell("A6", {"bold": True})
client.format_cell("B6", {"bold": True})
client.format_cell("C6", {"bold": True})
client.format_cell("D6", {"bold": True})

# Save the workbook
client.save_file("/tmp/sales_analysis.ods")

# Export as PDF
client.export_pdf("/tmp/sales_analysis.pdf")

client.close()
```

### Multi-Sheet Workbook

```python
from envs.libreoffice_env import LibreOfficeEnv

client = LibreOfficeEnv.from_docker_image("libreoffice-env:latest")
client.reset()

# Rename the first sheet
client.rename_sheet("Sheet1", "Summary")

# Add data to summary sheet
client.set_cell("A1", "Monthly Sales Summary")
client.set_cell("A2", "Total Sales:")
client.set_formula("B2", "=SUM(January!B:B)")

# Add January sheet
client.add_sheet("January")
client.set_range("A1:B4", [
    ["Date", "Sales"],
    ["2024-01-01", 1000],
    ["2024-01-02", 1200],
    ["2024-01-03", 1500]
])

# Add February sheet
client.add_sheet("February")
client.set_range("A1:B4", [
    ["Date", "Sales"],
    ["2024-02-01", 1100],
    ["2024-02-02", 1300],
    ["2024-02-03", 1600]
])

# Update summary formula to include February
client.set_formula("B2", "=SUM(January!B:B)+SUM(February!B:B)")

# Save the workbook
client.save_file("/tmp/monthly_sales.ods")

client.close()
```

## Error Handling

The environment provides comprehensive error handling:

```python
# Invalid cell reference
result = client.set_cell("ZZ999", "test")
if not result.observation.success:
    print(f"Error: {result.observation.error_message}")

# Invalid sheet name
result = client.get_cell("A1", sheet="NonExistentSheet")
if not result.observation.success:
    print(f"Error: {result.observation.error_message}")

# File not found
result = client.open_file("/nonexistent/file.ods")
if not result.observation.success:
    print(f"Error: {result.observation.error_message}")
```

## Dependencies

- LibreOffice (with Calc)
- Python-UNO bridge
- uno-python package
- unohelper package

## Limitations

- Requires LibreOffice to be installed and running
- Some advanced LibreOffice features may not be available through UNO
- File operations are limited to the container's filesystem
- Performance may be slower than native spreadsheet applications for large datasets

## Troubleshooting

### LibreOffice Connection Issues
- Ensure LibreOffice is properly installed
- Check that the Python-UNO bridge is working
- Verify environment variables are set correctly

### File Permission Issues
- Ensure the container has write permissions to the target directory
- Check that file paths are accessible from within the container

### Memory Issues
- Large spreadsheets may require more memory
- Consider processing data in smaller chunks
- Monitor container memory usage

## Contributing

To contribute to the LibreOffice environment:

1. Follow the existing code patterns
2. Add comprehensive error handling
3. Include docstrings for all new methods
4. Add tests for new functionality
5. Update this README with new features

## License

This project is licensed under the BSD-style license found in the LICENSE file.
