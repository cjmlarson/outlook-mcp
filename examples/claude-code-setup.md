# Claude Code Setup Examples

## Global Installation (Available in all projects)

```bash
# Add to global scope
claude mcp add -s global outlook node C:/Users/yourusername/outlook-mcp/src/index.js

# Verify installation
claude mcp list
```

## Project-Specific Installation

```bash
# Add to current project only
claude mcp add outlook node C:/Users/yourusername/outlook-mcp/src/index.js

# This creates .claude/settings.json in your project
```

## With Environment Variables

```bash
# If you need to pass environment variables
claude mcp add outlook --env OUTLOOK_PROFILE=MyProfile -- node C:/Users/yourusername/outlook-mcp/src/index.js
```

## Removing the Server

```bash
# Remove from global scope
claude mcp remove -s global outlook

# Remove from project
claude mcp remove outlook
```

## Testing the Connection

After installation, test with:
```
Ask Claude: "Use outlook_list to show my Outlook accounts"
```