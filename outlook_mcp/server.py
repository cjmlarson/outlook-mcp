#!/usr/bin/env python3
"""
Outlook MCP Server - Python implementation
Provides tools for Microsoft Outlook integration via COM automation
"""

import sys
import asyncio
import subprocess
from pathlib import Path
from typing import Optional, List, Literal

try:
    from mcp.server.fastmcp import FastMCP
except ImportError:
    print("Error: MCP SDK not installed. Please run: pip install mcp", file=sys.stderr)
    sys.exit(1)

# Create the MCP server instance
mcp = FastMCP("outlook-mcp")

# Get the directory where this module is located
MODULE_DIR = Path(__file__).parent


def run_outlook_script_sync(script_name: str, args: List[str]) -> str:
    """
    Run an Outlook Python script and return its output (synchronous version).
    
    Args:
        script_name: Name of the script file (e.g., 'outlook_list.py')
        args: Command line arguments to pass to the script
        
    Returns:
        The stdout output from the script
    """
    script_path = MODULE_DIR / script_name
    
    # Build the command
    cmd = [sys.executable, str(script_path)] + args
    
    # Run the script
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='ignore'
    )
    
    if result.returncode != 0:
        raise RuntimeError(f"Script failed: {result.stderr}")
    
    return result.stdout


@mcp.tool()
def outlook_read(
    entry_id: str,
    json: bool = False,
    save_attachments: bool = False,
    save_html: str = "",
    save_text: str = ""
) -> str:
    """
    Read full Outlook item content by EntryID (email, calendar, contact, task, etc.)
    
    Args:
        entry_id: The EntryID of the item to read
        json: Output as JSON instead of formatted text
        save_attachments: Save all attachments to temp folder
        save_html: Save HTML version to file (for emails)
        save_text: Save text version to file
    """
    args = [entry_id]
    
    if json:
        args.append('--json')
    if save_attachments:
        args.append('--save-attachments')
    if save_html:
        args.extend(['--save-html', save_html])
    if save_text:
        args.extend(['--save-text', save_text])
    
    result = run_outlook_script_sync('outlook_read.py', args)
    return result


@mcp.tool()
def outlook_list(
    path: str = "",
    all: bool = False,
    count: int = 50
) -> str:
    """
    List Outlook accounts, folders, or items (equivalent to Unix ls).
    Use this FIRST to identify folders before searching.
    
    Args:
        path: Path to list (e.g., "myaccount" or "myaccount/Inbox")
        all: Show all items/folders including system folders
        count: Number of items to show (default: 50)
    """
    args = []
    
    if path:
        args.append(path)
    if all:
        args.append('--all')
    if count != 50:
        args.extend(['--count', str(count)])
    
    result = run_outlook_script_sync('outlook_list.py', args)
    return result or "No items found"


@mcp.tool()
def outlook_filter(
    path: str = "",
    since: str = "",
    until: str = "",
    days: Optional[int] = None,
    from_sender: str = "",  # 'from' is a Python keyword, so using 'from_sender'
    type: str = "",
    unread: bool = False,
    max_items: int = 100
) -> str:
    """
    Filter Outlook items by properties like date, sender, type (equivalent to Unix find).
    Faster than content search.
    
    Args:
        path: Path to filter (e.g., "myaccount" or "myaccount/Inbox")
        since: Start date (YYYY-MM-DD)
        until: End date (YYYY-MM-DD)
        days: Items from last N days
        from_sender: Filter by sender/organizer
        type: Filter by type: email, event, contact, task
        unread: Only unread items
        max_items: Maximum items to return (default: 100)
    """
    args = []
    
    if path:
        args.append(path)
    if since:
        args.extend(['--since', since])
    if until:
        args.extend(['--until', until])
    if days is not None:
        args.extend(['--days', str(days)])
    if from_sender:
        args.extend(['--from', from_sender])
    if type:
        args.extend(['--type', type])
    if unread:
        args.append('--unread')
    if max_items != 100:
        args.extend(['--max-items', str(max_items)])
    
    result = run_outlook_script_sync('outlook_filter.py', args)
    return result or "No items found"


@mcp.tool()
def outlook_search(
    pattern: str,
    path: str,
    output_mode: Literal["list", "content"] = "list",
    since: str = "",
    until: str = "",
    offset: int = 0
) -> str:
    """
    Fast Outlook search using DASL queries.
    
    Search syntax: Space=OR ("United ZRH" finds either term), 
    Ampersand=AND ("United&ZRH" finds both), 
    Combined ("ZRH EWR&United" = (ZRH OR EWR) AND United).
    Path is REQUIRED - use outlook_list first to find folders.
    
    Args:
        pattern: Search pattern. Space=OR, Ampersand=AND. Examples: "United", "ZRH EWR", "flight&ZRH EWR"
        path: Path to search (REQUIRED). Format: "account/folder". Use outlook_list to find paths.
        output_mode: Output mode: list (fast, metadata only) or content (with match snippets)
        since: Start date (YYYY-MM-DD)
        until: End date (YYYY-MM-DD)
        offset: Pagination offset (results shown 25 at a time)
    """
    args = [pattern, path]
    
    args.extend(['--output-mode', output_mode])
    
    if since:
        args.extend(['--since', since])
    if until:
        args.extend(['--until', until])
    if offset != 0:
        args.extend(['--offset', str(offset)])
    
    result = run_outlook_script_sync('outlook_search.py', args)
    return result or "No matches found"


def serve():
    """
    Run the MCP server.
    This is the entry point for the 'outlook-mcp' command.
    """
    try:
        # FastMCP handles its own server lifecycle
        mcp.run()
    except KeyboardInterrupt:
        print("\nServer stopped by user", file=sys.stderr)
        sys.exit(0)
    except Exception as e:
        print(f"Server error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    serve()