#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { v4 as uuidv4 } from 'uuid';
import { promises as fs } from 'fs';
import path from 'path';
import XLSX from 'xlsx';

// Export directory configuration
const EXPORT_DIR = '/tmp/protex-intelligence-file-exports';

/**
 * Calculate file size string from buffer
 */
function getFileSizeString(buffer) {
  const bytes = buffer.length;
  const kb = Math.ceil(bytes / 1024);
  return kb < 1024 ? `${kb} KB` : `${(kb / 1024).toFixed(2)} MB`;
}

/**
 * Ensure export directory exists, create if it doesn't
 */
async function ensureExportDirectory() {
  try {
    await fs.access(EXPORT_DIR);
    console.error(`✓ Export directory exists: ${EXPORT_DIR}`);
  } catch (error) {
    try {
      await fs.mkdir(EXPORT_DIR, { recursive: true });
      console.error(`✓ Created export directory: ${EXPORT_DIR}`);
    } catch (mkdirError) {
      console.error(`✗ Failed to create export directory: ${mkdirError.message}`);
      throw mkdirError;
    }
  }
}

/**
 * Write XLSX buffer to file system
 */
async function writeXLSXToFile(xlsxBuffer, filename) {
  await ensureExportDirectory();

  const filepath = path.join(EXPORT_DIR, filename);

  try {
    await fs.writeFile(filepath, xlsxBuffer);
    console.error(`✓ File written: ${filepath}`);
    return filepath;
  } catch (error) {
    console.error(`✗ Failed to write file: ${error.message}`);
    throw error;
  }
}

// Create MCP server
const server = new Server(
  {
    name: 'xlsx-export-mcp',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// List available tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: 'xlsx_export',
        description: 'Export data to Excel (XLSX) format and save to filesystem',
        inputSchema: {
          type: 'object',
          properties: {
            data: {
              type: 'array',
              description: 'Array of objects representing spreadsheet rows',
              items: {
                type: 'object',
              },
            },
            filename: {
              type: 'string',
              description: 'Filename for the exported file (without extension)',
              default: 'output',
            },
            sheetName: {
              type: 'string',
              description: 'Name of the worksheet/sheet within the Excel file',
              default: 'Sheet1',
            },
            description: {
              type: 'string',
              description: 'Optional description of the file contents',
            },
            headers: {
              type: 'array',
              description: 'Optional custom column headers',
              items: {
                type: 'string',
              },
            },
          },
          required: ['data'],
        },
      },
    ],
  };
});

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  if (name === 'xlsx_export') {
    try {
      const {
        data,
        filename = 'output',
        sheetName = 'Sheet1',
        description,
        headers,
      } = args;

      // Validate input
      if (!data || !Array.isArray(data)) {
        throw new Error('Data must be provided as an array of objects');
      }

      if (data.length === 0) {
        throw new Error('Data array cannot be empty');
      }

      // Create workbook and worksheet
      const wb = XLSX.utils.book_new();

      // If custom headers are provided, use them
      let ws;
      if (headers && Array.isArray(headers)) {
        ws = XLSX.utils.json_to_sheet(data, { header: headers });
      } else {
        ws = XLSX.utils.json_to_sheet(data);
      }

      XLSX.utils.book_append_sheet(wb, ws, sheetName);

      // Generate UUID and filename
      const uuid = uuidv4();
      const sanitizedFilename = filename.replace(/[^a-z0-9_-]/gi, '_');
      const fullFilename = `${sanitizedFilename}_${uuid}.xlsx`;

      // Write XLSX to buffer (in-memory)
      const fileBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      const fileSize = getFileSizeString(fileBuffer);

      // Calculate row/column counts
      const rowCount = data.length;
      const columnCount = Object.keys(data[0] || {}).length;

      // Write XLSX to file system
      const filepath = await writeXLSXToFile(fileBuffer, fullFilename);

      console.error(`✅ XLSX generated: ${fullFilename} (${fileSize})`);
      console.error(`   Rows: ${rowCount}, Columns: ${columnCount}, Sheet: ${sheetName}`);
      console.error(`   Saved to: ${filepath}`);

      // Return simplified response with essential information
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                path: fullFilename,
                filetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                filename: fullFilename,
                filesize: fileSize,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      console.error('Error processing XLSX export:', error);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: false,
                error: error.message || 'Unknown error',
              },
              null,
              2
            ),
          },
        ],
        isError: true,
      };
    }
  }

  throw new Error(`Unknown tool: ${name}`);
});

// Start server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('XLSX Export MCP Server running on stdio');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
