#!/usr/bin/env python3
"""
Simple example of using the LibreOffice Spreadsheet Environment.

This example demonstrates basic spreadsheet operations using the LibreOffice environment.
"""

import sys
from pathlib import Path

# Add the src directory to the Python path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from envs.libreoffice_env import LibreOfficeAction, LibreOfficeEnv


def main():
    """Main function demonstrating LibreOffice environment usage."""
    print("LibreOffice Spreadsheet Environment Example")
    print("=" * 50)
    
    try:
        # Create environment from Docker image
        print("Creating LibreOffice environment...")
        client = LibreOfficeEnv.from_docker_image("libreoffice-env:latest")
        
        # Reset the environment
        print("Resetting environment...")
        result = client.reset()
        print(f"Reset result: {result.observation.result}")
        print(f"Success: {result.observation.success}")
        print(f"Current sheet: {result.observation.current_sheet}")
        
        if not result.observation.success:
            print("Failed to initialize LibreOffice environment")
            return
        
        # Create a simple data table
        print("\nCreating a data table...")
        data = [
            ["Product", "Quantity", "Price", "Total"],
            ["Widget A", 10, 15.50, "=B2*C2"],
            ["Widget B", 5, 25.00, "=B3*C3"],
            ["Widget C", 8, 12.75, "=B4*C4"],
            ["Widget D", 12, 8.25, "=B5*C5"]
        ]
        
        # Set the data in the spreadsheet
        result = client.set_range("A1:D5", data)
        print(f"Set range result: {result.observation.result}")
        print(f"Success: {result.observation.success}")
        
        # Format the header row
        print("\nFormatting header row...")
        for col in ["A", "B", "C", "D"]:
            result = client.format_cell(f"{col}1", {"bold": True})
            print(f"Formatted {col}1: {result.observation.success}")
        
        # Add a total row
        print("\nAdding total row...")
        result = client.set_cell("A6", "TOTAL")
        result = client.set_formula("B6", "=SUM(B2:B5)")
        result = client.set_formula("C6", "=AVERAGE(C2:C5)")
        result = client.set_formula("D6", "=SUM(D2:D5)")
        
        # Format the total row
        for col in ["A", "B", "C", "D"]:
            result = client.format_cell(f"{col}6", {"bold": True})
        
        # Get the calculated totals
        print("\nRetrieving calculated values...")
        result = client.get_cell("B6")
        print(f"Total quantity: {result.observation.data}")
        
        result = client.get_cell("C6")
        print(f"Average price: {result.observation.data}")
        
        result = client.get_cell("D6")
        print(f"Total revenue: {result.observation.data}")
        
        # Add a new sheet for analysis
        print("\nAdding analysis sheet...")
        result = client.add_sheet("Analysis")
        print(f"Add sheet result: {result.observation.result}")
        
        # Create a simple chart data
        chart_data = [
            ["Product", "Revenue"],
            ["Widget A", "=Sheet1.D2"],
            ["Widget B", "=Sheet1.D3"],
            ["Widget C", "=Sheet1.D4"],
            ["Widget D", "=Sheet1.D5"]
        ]
        
        result = client.set_range("A1:B5", chart_data)
        print(f"Set chart data: {result.observation.success}")
        
        # Save the workbook
        print("\nSaving workbook...")
        result = client.save_file("/tmp/example_spreadsheet.ods")
        print(f"Save result: {result.observation.result}")
        print(f"Success: {result.observation.success}")
        
        # Export as PDF
        print("\nExporting as PDF...")
        result = client.export_pdf("/tmp/example_spreadsheet.pdf")
        print(f"Export PDF result: {result.observation.result}")
        print(f"Success: {result.observation.success}")
        
        # Export as CSV
        print("\nExporting as CSV...")
        result = client.export_csv("/tmp/example_spreadsheet.csv")
        print(f"Export CSV result: {result.observation.result}")
        print(f"Success: {result.observation.success}")
        
        # Get final state
        print("\nFinal state:")
        state = client.state()
        print(f"Episode ID: {state.episode_id}")
        print(f"Step count: {state.step_count}")
        
        print("\nExample completed successfully!")
        print("Files created:")
        print("- /tmp/example_spreadsheet.ods")
        print("- /tmp/example_spreadsheet.pdf")
        print("- /tmp/example_spreadsheet.csv")
        
    except Exception as e:
        print(f"Error: {e}")
        print("Make sure the LibreOffice environment Docker image is built and available.")
        print("Run: docker build -t libreoffice-env:latest -f src/envs/libreoffice_env/server/Dockerfile .")
    
    finally:
        # Cleanup
        try:
            client.close()
        except:
            pass


if __name__ == "__main__":
    main()
