# Outlook MCP Server

A Model Context Protocol (MCP) server that provides tools for Microsoft Outlook integration. Search, read, list, and filter emails, calendar events, contacts, and tasks directly from Claude Desktop or Claude Code.

## Features

- **outlook_list** - Browse accounts, folders, and items (like Unix `ls`)
- **outlook_filter** - Filter items by date, sender, type (like Unix `find`) 
- **outlook_search** - Full-text search with advanced query syntax
- **outlook_read** - Read complete item content including attachments

## Prerequisites

- **Windows OS** (required - uses Windows COM automation)
- **Microsoft Outlook** desktop application installed and configured
- **Python 3.8+** with `pywin32` package
- **Node.js 16+** for the MCP server

## Installation

### 1. Install Python dependencies

```bash
pip install pywin32
```

### 2. Install the MCP server

```bash
git clone https://github.com/cjmlarson/outlook-mcp.git
cd outlook-mcp
npm install
```

## Configuration

### Claude Desktop

Add to your Claude Desktop config file (`%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["C:/path/to/outlook-mcp/src/index.js"]
    }
  }
}
```

### Claude Code

```bash
# Add to global scope (available in all projects)
claude mcp add -s global outlook node C:/path/to/outlook-mcp/src/index.js

# Or add to current project only
claude mcp add outlook node C:/path/to/outlook-mcp/src/index.js
```

## Usage Examples

### List accounts and folders
```
Use outlook_list to see available accounts
Use outlook_list myaccount to see folders
Use outlook_list myaccount/Inbox to see recent emails
```

### Search for emails
```
Search for "flight confirmation" in Inbox:
outlook_search "flight confirmation" "myaccount/Inbox"

Search with AND logic:
outlook_search "United&ZRH" "myaccount/Inbox"

Search with OR logic (space = OR):
outlook_search "ZRH EWR JFK" "myaccount/Travel"
```

### Filter by properties
```
Get emails from last 7 days:
outlook_filter "myaccount/Inbox" --days 7

Get unread emails from specific sender:
outlook_filter "myaccount/Inbox" --unread --from "boss@company.com"

Get calendar events for next month:
outlook_filter "myaccount/Calendar" --type event --since 2025-02-01 --until 2025-02-28
```

### Read full content
```
Read email by EntryID:
outlook_read "00000000F0616777..."

Save attachments:
outlook_read "00000000F0616777..." --save-attachments
```

## Search Syntax

The search tool uses a powerful query syntax:
- **Space = OR**: `"United ZRH"` finds items with either "United" OR "ZRH"
- **Ampersand = AND**: `"United&ZRH"` finds items with both "United" AND "ZRH"
- **Combined**: `"ZRH EWR&United"` means (ZRH OR EWR) AND United

## Performance Tips

1. **Use outlook_list first** to identify the correct folder paths
2. **Filter is faster than search** for date/sender/type queries
3. **Use pagination** for large result sets (offset parameter)
4. **Be specific with paths** to avoid searching unnecessary folders

## Troubleshooting

### "Outlook.Application not found"
- Ensure Outlook desktop is installed (not just web/mobile)
- Outlook must have been opened at least once

### "Access denied" errors
- Run your terminal as Administrator if needed
- Check Outlook isn't showing security prompts

### Unicode/emoji issues
- The tools handle Unicode safely
- Emojis are stripped to prevent encoding errors

## Security Notes

- This tool accesses your local Outlook data via COM
- No data is sent to external servers
- EntryIDs are specific to your Outlook profile
- Always review the code before granting access

## License

MIT License - See LICENSE file for details

## Contributing

Pull requests welcome! Please ensure:
- Windows compatibility is maintained
- Error handling for COM exceptions
- Unicode text handling is robust

## Acknowledgments

Built for use with Anthropic's Claude via the Model Context Protocol.

## Author

Connor Larson ([@cjmlarson](https://github.com/cjmlarson))