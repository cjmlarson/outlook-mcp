import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { exec } from 'child_process';
import { promisify } from 'util';
import path from 'path';
import { fileURLToPath } from 'url';
import { z } from 'zod';

const execAsync = promisify(exec);
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Initialize MCP server
const server = new McpServer({
  name: 'outlook-mcp',
  version: '1.0.0'
}, {
  capabilities: {
    tools: {}
  }
});

// Tool: Read Outlook Item
server.registerTool('outlook_read', {
  description: 'Read full Outlook item content by EntryID (email, calendar, contact, task, etc.)',
  inputSchema: {
    entry_id: z.string().describe('The EntryID of the item to read'),
    json: z.boolean().default(false).describe('Output as JSON instead of formatted text'),
    save_attachments: z.boolean().default(false).describe('Save all attachments to temp folder'),
    save_html: z.string().default('').describe('Save HTML version to file (for emails)'),
    save_text: z.string().default('').describe('Save text version to file')
  }
}, async ({ entry_id, json = false, save_attachments = false, save_html = '', save_text = '' }) => {
  try {
    const scriptPath = path.join(__dirname, '..', 'python', 'outlook_read.py');
    let command = `py "${scriptPath}" "${entry_id}"`;
    
    if (json) {
      command += ' --json';
    }
    if (save_attachments) {
      command += ' --save-attachments';
    }
    if (save_html) {
      command += ` --save-html "${save_html}"`;
    }
    if (save_text) {
      command += ` --save-text "${save_text}"`;
    }
    
    const { stdout } = await execAsync(command);
    return {
      content: [
        {
          type: 'text',
          text: stdout
        }
      ]
    };
  } catch (error) {
    throw new Error(`Failed to read Outlook item: ${error.message}`);
  }
});

// Tool: Outlook List (unified listing)
server.registerTool('outlook_list', {
  description: 'List Outlook accounts, folders, or items (equivalent to Unix ls). Use this FIRST to identify folders before searching',
  inputSchema: {
    path: z.string().default('').describe('Path to list (e.g., "myaccount" or "myaccount/Inbox")'),
    all: z.boolean().default(false).describe('Show all items/folders including system folders'),
    count: z.number().default(50).describe('Number of items to show (default: 50)')
  }
}, async ({ path: outlookPath = '', all = false, count = 50 }) => {
  try {
    const scriptPath = path.join(__dirname, '..', 'python', 'outlook_list.py');
    let command = `py "${scriptPath}"`;
    
    if (outlookPath) {
      command += ` "${outlookPath}"`;
    }
    if (all) {
      command += ' --all';
    }
    if (count !== 50) {
      command += ` --count ${count}`;
    }
    
    const { stdout } = await execAsync(command);
    
    return {
      content: [
        {
          type: 'text',
          text: stdout || 'No items found'
        }
      ]
    };
  } catch (error) {
    throw new Error(`Failed to list Outlook items: ${error.message}`);
  }
});

// Tool: Outlook Filter (property-based filtering)
server.registerTool('outlook_filter', {
  description: 'Filter Outlook items by properties like date, sender, type (equivalent to Unix find). Faster than content search',
  inputSchema: {
    path: z.string().default('').describe('Path to filter (e.g., "myaccount" or "myaccount/Inbox")'),
    since: z.string().default('').describe('Start date (YYYY-MM-DD)'),
    until: z.string().default('').describe('End date (YYYY-MM-DD)'),
    days: z.number().optional().describe('Items from last N days'),
    from: z.string().default('').describe('Filter by sender/organizer'),
    type: z.string().default('').describe('Filter by type: email, event, contact, task'),
    unread: z.boolean().default(false).describe('Only unread items'),
    max_items: z.number().default(100).describe('Maximum items to return (default: 100)')
  }
}, async ({ path: outlookPath = '', since = '', until = '', days, from = '', type = '', unread = false, max_items = 100 }) => {
  try {
    const scriptPath = path.join(__dirname, '..', 'python', 'outlook_filter.py');
    let command = `py "${scriptPath}"`;
    
    if (outlookPath) {
      command += ` "${outlookPath}"`;
    }
    if (since) {
      command += ` --since "${since}"`;
    }
    if (until) {
      command += ` --until "${until}"`;
    }
    if (days !== undefined) {
      command += ` --days ${days}`;
    }
    if (from) {
      command += ` --from "${from}"`;
    }
    if (type) {
      command += ` --type "${type}"`;
    }
    if (unread) {
      command += ' --unread';
    }
    if (max_items !== 100) {
      command += ` --max-items ${max_items}`;
    }
    
    const { stdout } = await execAsync(command);
    
    return {
      content: [
        {
          type: 'text',
          text: stdout || 'No items found'
        }
      ]
    };
  } catch (error) {
    throw new Error(`Failed to filter Outlook items: ${error.message}`);
  }
});

// Tool: Outlook Search (content search)
server.registerTool('outlook_search', {
  description: 'Fast Outlook search using DASL queries. Search syntax: Space=OR ("United ZRH" finds either term), Ampersand=AND ("United&ZRH" finds both), Combined ("ZRH EWR&United" = (ZRH OR EWR) AND United). Path is REQUIRED - use outlook_list first to find folders.',
  inputSchema: {
    pattern: z.string().describe('Search pattern. Space=OR, Ampersand=AND. Examples: "United", "ZRH EWR", "flight&ZRH EWR"'),
    path: z.string().describe('Path to search (REQUIRED). Format: "account/folder". Use outlook_list to find paths.'),
    output_mode: z.enum(['list', 'content']).default('list').describe('Output mode: list (fast, metadata only) or content (with match snippets)'),
    since: z.string().default('').describe('Start date (YYYY-MM-DD)'),
    until: z.string().default('').describe('End date (YYYY-MM-DD)'),
    offset: z.number().default(0).describe('Pagination offset (results shown 25 at a time)')
  }
}, async ({ pattern, path: outlookPath, output_mode = 'list', since = '', until = '', offset = 0 }) => {
  try {
    const scriptPath = path.join(__dirname, '..', 'python', 'outlook_search.py');
    let command = `py "${scriptPath}" "${pattern}" "${outlookPath}"`;
    
    // Add parameters
    command += ` --output-mode ${output_mode}`;
    
    if (since) {
      command += ` --since "${since}"`;
    }
    if (until) {
      command += ` --until "${until}"`;
    }
    if (offset !== 0) {
      command += ` --offset ${offset}`;
    }
    
    const { stdout } = await execAsync(command);
    
    return {
      content: [
        {
          type: 'text',
          text: stdout || 'No matches found'
        }
      ]
    };
  } catch (error) {
    throw new Error(`Failed to search Outlook: ${error.message}`);
  }
});

// Start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Outlook MCP server running on stdio');
}

main().catch((error) => {
  console.error('Failed to start server:', error);
  process.exit(1);
});