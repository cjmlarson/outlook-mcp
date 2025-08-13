# Claude Code Setup for Outlook MCP

This guide shows how to configure the Outlook MCP server with Claude Code.

## Installation

### Option 1: Global npm install (Recommended)

```bash
# Install globally 
npm install -g outlook-mcp

# Configure with Claude Code
claude mcp add -s global outlook outlook-mcp
```

### Option 2: From source

```bash
git clone https://github.com/cjmlarson/outlook-mcp.git
cd outlook-mcp
npm install

# Global scope (available in all projects)
claude mcp add -s global outlook node C:/path/to/outlook-mcp/src/index.js

# Or project scope only
claude mcp add outlook node C:/path/to/outlook-mcp/src/index.js
```

## Verify Installation

Check that the server is properly configured:

```bash
claude mcp list
```

You should see `outlook` in the list with a ✓ Connected status.

## Usage

Once configured, you can use the Outlook tools in Claude Code:

- `outlook_list` - Browse accounts and folders (like `ls`)
- `outlook_filter` - Filter by date, sender, type (like `find`)
- `outlook_search` - Search email content with advanced syntax
- `outlook_read` - Read full email content including attachments

### Example Commands

```
"List my Outlook folders"
"Show emails from last week"
"Search for flight confirmations in my Travel folder"
"Read the latest email from my boss"
```

## Advanced Configuration

### Project-Specific Installation
```bash
# Add to current project only
claude mcp add outlook outlook-mcp

# This creates .claude/settings.json in your project
```

### With Environment Variables
```bash
# If you need to pass environment variables
claude mcp add outlook --env OUTLOOK_PROFILE=MyProfile outlook-mcp
```

### Removing the Server
```bash
# Remove from global scope
claude mcp remove -s global outlook

# Remove from project
claude mcp remove outlook
```

## Troubleshooting

If you encounter issues:

1. **Server not found**: Make sure you ran `npm install -g outlook-mcp`
2. **Tools not working**: Ensure Outlook is installed and opened at least once
3. **Permission errors**: Try running terminal as Administrator
4. **Path issues**: For source install, verify absolute paths are correct
5. **Python missing**: Ensure Python 3.8+ is installed (used by the COM scripts)

## Testing the Connection

After installation, test with:
```
Ask Claude: "Use outlook_list to show my Outlook accounts"
```