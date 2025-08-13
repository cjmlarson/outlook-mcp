#!/usr/bin/env python3
"""
Entry point for Desktop Extension (.dxt) packaging.
This wrapper ensures the server runs correctly when packaged as a DXT.
"""

import sys
import os
from pathlib import Path

# Add the server directory to Python path so we can import outlook_mcp
server_dir = Path(__file__).parent
sys.path.insert(0, str(server_dir))

# Add the lib directory for bundled dependencies
lib_dir = server_dir.parent / "lib"
if lib_dir.exists():
    sys.path.insert(0, str(lib_dir))

# Import and run the MCP server
try:
    from outlook_mcp.server import serve
    serve()
except ImportError as e:
    print(f"Error: Failed to import outlook_mcp server: {e}", file=sys.stderr)
    print(f"Python path: {sys.path}", file=sys.stderr)
    sys.exit(1)
except Exception as e:
    print(f"Error running server: {e}", file=sys.stderr)
    sys.exit(1)