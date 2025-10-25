# XLSX Export MCP Server

A Model Context Protocol (MCP) server that provides Excel (XLSX) export functionality. This server allows you to export structured data to Excel format files.

## Features

- Export array of objects to XLSX format
- Customizable filename and sheet name
- Optional custom column headers
- File size reporting
- UUID-based unique filenames
- Saves files to `/tmp/protex-intelligence-file-exports/`

## Installation

```bash
# Clone and install
git clone <repository-url>
cd xlsx-export-mcp-py
uv pip install -e .
```

## Usage

The server provides one tool:

### `xlsx_export`

Exports data to Excel (XLSX) format.

**Parameters:**
- `data` (required): Array of objects representing spreadsheet rows
- `filename` (optional): Filename for the exported file (without extension), defaults to "output"
- `sheetName` (optional): Name of the worksheet/sheet within the Excel file, defaults to "Sheet1"
- `description` (optional): Description of the file contents
- `headers` (optional): Custom column headers array

**Example:**
```json
{
  "data": [
    {"name": "John", "age": 30, "city": "New York"},
    {"name": "Jane", "age": 25, "city": "Boston"}
  ],
  "filename": "people_data",
  "sheetName": "People",
  "headers": ["Name", "Age", "City"]
}
```

## Running the Server

```bash
xlsx-export-mcp
```

## Development

```bash
# Install in development mode
uv pip install -e .

# Run tests (if available)
python -m pytest
```

## License

MIT