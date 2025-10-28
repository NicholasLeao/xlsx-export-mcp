#!/usr/bin/env python3
"""XLSX Export MCP Server - Python implementation."""

import json
import sys
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import io

# Export directory configuration
EXPORT_DIR = "/tmp/protex-intelligence-file-exports"


def convert_to_xlsx(
    data: List[Dict[str, Any]], 
    sheet_name: str = "Sheet1",
    headers: Optional[List[str]] = None
) -> bytes:
    """Convert array of objects to XLSX bytes."""
    if not data:
        return b""
    
    # Create workbook and worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    
    # Get headers from first object or use provided headers
    if headers:
        field_names = headers
    else:
        field_names = list(data[0].keys())
    
    # Write headers
    for col_idx, header in enumerate(field_names, 1):
        ws.cell(row=1, column=col_idx, value=header)
    
    # Write data rows
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, field_name in enumerate(field_names, 1):
            value = row_data.get(field_name, "")
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def get_file_size_string(content: bytes) -> str:
    """Calculate file size string from bytes content."""
    bytes_size = len(content)
    kb = bytes_size / 1024
    
    if kb < 1024:
        return f"{kb:.0f} KB" if kb >= 1 else "1 KB"
    else:
        return f"{kb / 1024:.2f} MB"


async def ensure_export_directory() -> None:
    """Ensure export directory exists, create if it doesn't."""
    export_path = Path(EXPORT_DIR)
    
    if export_path.exists():
        print(f"✓ Export directory exists: {EXPORT_DIR}", file=sys.stderr)
    else:
        try:
            export_path.mkdir(parents=True, exist_ok=True)
            print(f"✓ Created export directory: {EXPORT_DIR}", file=sys.stderr)
        except Exception as e:
            print(f"✗ Failed to create export directory: {e}", file=sys.stderr)
            raise


async def write_xlsx_to_file(xlsx_content: bytes, filename: str) -> str:
    """Write XLSX content to file system."""
    await ensure_export_directory()
    
    filepath = Path(EXPORT_DIR) / filename
    
    try:
        filepath.write_bytes(xlsx_content)
        print(f"✓ File written: {filepath}", file=sys.stderr)
        return str(filepath)
    except Exception as e:
        print(f"✗ Failed to write file: {e}", file=sys.stderr)
        raise


# Create FastMCP server
mcp = FastMCP("xlsx-export-mcp")


@mcp.tool()
async def xlsx_export(
    data: List[Dict[str, Any]],
    filename: str = "output",
    sheet_name: str = "Sheet1",
    description: str = None,
    headers: List[str] = None
) -> dict:
    """Export data to Excel (XLSX) format and save to filesystem.
    
    Args:
        data: Array of objects representing spreadsheet rows
        filename: Filename for the exported file (without extension)
        sheet_name: Name of the worksheet/sheet within the Excel file
        description: Optional description of the file contents
        headers: Optional custom column headers
        
    Returns:
        Dictionary with export results including path and file info
    """
    try:
        # Validate input
        if not data or not isinstance(data, list):
            raise ValueError("Data must be provided as an array of objects")
        
        if len(data) == 0:
            raise ValueError("Data array cannot be empty")
        
        # Convert to XLSX
        xlsx_content = convert_to_xlsx(data, sheet_name, headers)
        
        # Generate UUID and filename
        file_uuid = str(uuid.uuid4())
        sanitized_filename = "".join(c if c.isalnum() or c in "_-" else "_" for c in filename)
        full_filename = f"{sanitized_filename}_{file_uuid}.xlsx"
        file_size = get_file_size_string(xlsx_content)
        row_count = len(data)
        column_count = len(data[0].keys()) if data else 0
        
        # Write XLSX to file system
        filepath = await write_xlsx_to_file(xlsx_content, full_filename)
        
        print(f"✅ XLSX generated: {full_filename} ({file_size})", file=sys.stderr)
        print(f"   Rows: {row_count}, Columns: {column_count}, Sheet: {sheet_name}", file=sys.stderr)
        print(f"   Saved to: {filepath}", file=sys.stderr)
        
        # Return simplified response with essential information
        return {
            "path": full_filename,
            "filetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "filename": full_filename,
            "filesize": file_size,
        }
        
    except Exception as error:
        print(f"Error processing XLSX export: {error}", file=sys.stderr)
        
        return {
            "success": False,
            "error": str(error),
        }


def cli_main():
    """CLI entry point."""
    mcp.run()


if __name__ == "__main__":
    cli_main()