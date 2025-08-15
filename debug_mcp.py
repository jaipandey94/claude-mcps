#!/usr/bin/env python3
"""
Minimal debug version of MCP server to test basic functionality
"""

import asyncio
import sys
from mcp.server import Server
from mcp.server.models import InitializationOptions
import mcp.server.stdio
import mcp.types as types

# Initialize MCP server
server = Server("outlook-connector")

@server.list_tools()
async def handle_list_tools():
    """List available tools"""
    print("📋 list_tools called", file=sys.stderr)
    return [
        types.Tool(
            name="test_tool",
            description="A simple test tool",
            inputSchema={
                "type": "object",
                "properties": {
                    "message": {
                        "type": "string",
                        "description": "Test message"
                    }
                }
            }
        )
    ]

@server.call_tool()
async def handle_call_tool(name: str, arguments: dict):
    """Handle tool calls"""
    print(f"🔧 call_tool: {name} with {arguments}", file=sys.stderr)
    
    if name == "test_tool":
        message = arguments.get("message", "Hello from MCP!")
        return [types.TextContent(
            type="text",
            text=f"✅ Test successful: {message}"
        )]
    else:
        return [types.TextContent(
            type="text",
            text=f"❌ Unknown tool: {name}"
        )]

async def main():
    """Main entry point"""
    try:
        print("🚀 Starting debug MCP server...", file=sys.stderr)
        
        # Set Windows event loop policy
        if sys.platform.startswith('win'):
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        
        # Run the MCP server
        async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
            print("📡 MCP server running...", file=sys.stderr)
            await server.run(
                read_stream,
                write_stream,
                InitializationOptions(
                    server_name="outlook-connector",
                    server_version="1.0.0"
                )
            )
    except Exception as e:
        print(f"❌ Server error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)

if __name__ == "__main__":
    asyncio.run(main())
