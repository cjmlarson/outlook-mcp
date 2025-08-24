# Outlook MCP Server

A Model Context Protocol (MCP) server that provides tools for Microsoft Outlook integration. Search, read, list, and filter emails, calendar events, contacts, and tasks directly from Claude Desktop or Claude Code. Works with any accounts loaded into (classic) Outlook, incl. Outlook/Exchange, IMAP, etc.

## Features

- **outlook_list** - Browse accounts, folders, and items (like `ls`)
- **outlook_filter** - Filter items by date, sender, type (like `find`) 
- **outlook_search** - Full-text search with advanced query syntax (like `grep`)
- **outlook_read** - Read complete item content

## Prerequisites

- **Windows OS** (required - uses Windows COM automation)
- **Microsoft Outlook** desktop application (not "new" Outlook)
- **Python 3.8+**
- **Node.js 16+**

## Limitations

- **Windows Only**: Requires Windows OS (7/8/10/11) due to COM automation dependency
- **Desktop Outlook Required**: Web and mobile versions not supported, must have desktop application
- **Single Profile Access**: Accesses only the default Outlook profile at a time
- **No Real-Time Sync**: Does not monitor for new emails in real-time, requires explicit queries
- **Local Processing Only**: Cannot access cloud-only email providers without Outlook configuration
- **Language Support**: Best results with English content, international character support varies by system locale

## Installation

### For Claude Desktop (Recommended) 🎯

**One-click installation with .dxt file:**

1. Download the latest [outlook-mcp.dxt](https://github.com/cjmlarson/outlook-mcp/releases/latest) file
2. Double-click the `.dxt` file to install
   - If double-clicking doesn't work, open Claude Desktop and go to:
   - File → Settings... → Extensions → Advanced Settings → Install Extension...
   - Then select the downloaded .dxt file
3. Restart Claude Desktop
4. ✅ Done! Start using Outlook tools in Claude

### For Claude Code 💻

**Install globally via npm:**

```bash
npm install -g outlook-mcp

# Add to Claude Code
claude mcp add -s global outlook outlook-mcp
```

**Or install from source:**
```bash
git clone https://github.com/cjmlarson/outlook-mcp.git
cd outlook-mcp
npm install

# Add to Claude Code (global scope)
claude mcp add -s global outlook node /path/to/outlook-mcp/src/index.js

# Or add to current project only
claude mcp add outlook node /path/to/outlook-mcp/src/index.js
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
outlook_read "HflJFcbL9/Wq/H25e7YI6F0SU6px57ltK5XvAjg6JhoZXH2ImgroeUtUb6rWQnIUqdg4dSrX=="

Save attachments:
outlook_read "HflJFcbL9/Wq/H25e7YI6F0SU6px57ltK5XvAjg6JhoZXH2ImgroeUtUb6rWQnIUqdg4dSrX==" --save-attachments
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

### Performance Expectations
- **Simple queries** (< 100 results): Under 3 seconds
- **Complex searches** (> 1000 items): Under 10 seconds
- **Cross-folder operations**: Under 15 seconds
- **Attachment processing**: Under 5 seconds per file

## Architecture

This MCP server uses a hybrid approach:
- **Node.js** handles the MCP protocol and process management
- **Python** scripts handle the actual Outlook COM automation
- This provides the best of both worlds: reliable subprocess handling and robust COM interaction

## Troubleshooting

### Installation Issues

#### "Outlook.Application not found"
**Symptoms**: Error when trying to use any outlook tool
**Causes & Solutions**:
- **Outlook not installed**: Ensure Microsoft Outlook desktop application is installed (Office 365, Outlook 2019/2021)
- **Outlook never opened**: Outlook must be launched at least once to initialize COM components
- **Outlook running in safe mode**: Restart Outlook normally
- **Registry corruption**: Try running `outlook.exe /resetfolders` from command line

#### "Access denied" or "Permission denied"
**Symptoms**: COM automation fails with permission errors
**Solutions**:
- **Run as Administrator**: Launch your terminal/Claude Code as Administrator
- **Outlook security prompts**: Check if Outlook is showing security dialog boxes
- **Antivirus blocking**: Temporarily disable antivirus COM protection
- **User profile issues**: Try creating a new Outlook profile

#### Tools not found in Claude Desktop
**Symptoms**: MCP tools don't appear after .dxt installation
**Solutions**:
- **Restart required**: Always restart Claude Desktop after .dxt installation
- **Check logs**: Review logs in `%APPDATA%\Claude\logs\` for error messages
- **Corrupted .dxt**: Try downloading and installing the .dxt file again
- **Path issues**: Ensure .dxt was installed to correct Claude extensions directory

### Runtime Issues

#### Python execution errors
**Symptoms**: "Python not found" or subprocess errors
**Solutions**:
- **Python installation**: Ensure Python 3.8+ is installed and in PATH
- **Python packages**: Install required packages: `pip install pywin32`
- **COM registration**: Run `python Scripts/pywin32_postinstall.py -install` as Administrator
- **Path resolution**: Use absolute Python path in MCP configuration if needed

#### Unicode/emoji handling issues
**Symptoms**: Garbled text or encoding errors in email content
**Solutions**:
- The tools automatically handle Unicode safely by stripping problematic characters
- If issues persist, check Outlook language settings
- Ensure Windows locale supports Unicode (UTF-8)

#### Performance issues
**Symptoms**: Slow response times or timeouts

**Solutions**:
- **Large mailboxes**: Use filters to narrow search scope
- **Outlook indexing**: Ensure Outlook search indexing is enabled and up-to-date
- **Memory usage**: Close unnecessary Outlook add-ins that might slow COM access
- **Network delays**: For Exchange accounts, check network connectivity

### Configuration Issues

#### MCP server not connecting
**Symptoms**: Server shows as disconnected in `claude mcp list`
**Solutions**:
- **Check configuration**: Verify MCP configuration with `claude mcp list`
- **Port conflicts**: Ensure no other processes are using MCP ports
- **Node.js version**: Verify Node.js 16+ is installed
- **Dependencies**: Run `npm install` in server directory

#### Wrong Outlook profile accessed
**Symptoms**: Tools access unexpected mailbox or show no data
**Solutions**:
- **Default profile**: Ensure correct Outlook profile is set as default
- **Profile switching**: Close and restart Outlook with desired profile
- **Manual profile**: Set specific profile in Outlook before using tools

### Advanced Diagnostics

#### Enable detailed logging
For deeper troubleshooting, you can enable verbose logging:
```bash
# Set environment variable for detailed logs
set DEBUG=outlook-mcp*
claude
```

#### Test COM automation directly
Test if Outlook COM is working:
```python
python -c "
import win32com.client
outlook = win32com.client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
print(f'Outlook version: {outlook.Version}')
print(f'Folders: {namespace.Folders.Count}')
"
```

#### Manual server testing
Test the MCP server directly:
```bash
# Navigate to server directory
cd /path/to/outlook-mcp
node src/index.js
# Should start without errors and show MCP protocol messages
```

### Getting Help

If you encounter issues not covered here:

1. **Check existing issues**: Search [GitHub Issues](https://github.com/cjmlarson/outlook-mcp/issues)
2. **Create detailed report**: Include error messages, system info, and steps to reproduce
3. **System information**: Include Windows version, Outlook version, Python version, Node.js version
4. **Log files**: Attach relevant log files from Claude Desktop or MCP server

## Getting Help

### Support Channels

#### **GitHub Discussions** (Recommended)
- **Questions & Answers**: Ask questions and help other users
- **Feature Ideas**: Propose and discuss new features
- **Show and Tell**: Share interesting use cases and workflows
- Visit: [GitHub Discussions](https://github.com/cjmlarson/outlook-mcp/discussions)

#### **GitHub Issues**
- **Bug Reports**: Report bugs with detailed reproduction steps
- **Feature Requests**: Request specific new functionality
- **Documentation Issues**: Report problems with docs
- Visit: [GitHub Issues](https://github.com/cjmlarson/outlook-mcp/issues)

#### **Direct Contact**
- **Email**: connor@cjmlarson.com (for security issues or private concerns)
- **GitHub**: [@cjmlarson](https://github.com/cjmlarson)

### Before Asking for Help

1. **Check Documentation**: Review README, EXAMPLES.md, and troubleshooting guide
2. **Search Existing Issues**: Your problem may already be solved
3. **Try Basic Diagnostics**: Run the diagnostic commands in troubleshooting section
4. **Gather Information**: Prepare system info, error messages, and logs

### How to Report Issues Effectively

**Include This Information:**
- Windows version (e.g., Windows 11 22H2)
- Outlook version (e.g., Microsoft 365, Outlook 2021)
- Python version (`python --version`)
- Node.js version (`node --version`)
- Complete error message and stack trace
- Steps to reproduce the issue
- What you expected vs. what happened

## Development

### Building .dxt files
```bash
# TODO: Add build script
npm run build:dxt
```

### Testing locally
```bash
# Test the Node.js server
node src/index.js

# Test with Claude Code
claude mcp add outlook-test node ./src/index.js
```

## Security Notes

- This tool accesses your local Outlook data via COM
- No data is sent to external servers
- EntryIDs are specific to your Outlook profile
- Always review the code before granting access

## License

MIT License - See LICENSE file for details

## Maintenance and Support

This project is maintained as actively as possible. Community involvement welcomed. To report issues, use [GitHub Issues](https://github.com/cjmlarson/outlook-mcp/issues) with detailed information.

## Contributing

Pull requests welcome! Please ensure:
- Windows compatibility is maintained
- Error handling for COM exceptions
- Unicode text handling is robust
- All tests pass and new features include tests
- Documentation updated for any API changes

### Development Setup
```bash
git clone https://github.com/cjmlarson/outlook-mcp.git
cd outlook-mcp
npm install
npm test
```

### Code Standards
- ES6+ JavaScript with clear error handling
- Python 3.8+ with type hints where applicable
- Comprehensive error messages and logging
- Security-first approach to data handling

## Acknowledgments

Built for use with Anthropic's Claude via the Model Context Protocol.

## Author

Connor Larson ([@cjmlarson](https://github.com/cjmlarson))

**Contact**: 
- GitHub Issues for bug reports and feature requests
- GitHub Discussions for questions and community support