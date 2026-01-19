#!/usr/bin/env python3
"""
Test script for mailtool MCP server
Sends JSON-RPC requests to verify server functionality
"""

import asyncio
import json
import subprocess
import sys


async def test_mcp_server():
    """Test MCP server by sending requests"""

    print("🧪 Testing Mailtool MCP Server\n")

    # Start MCP server process
    process = subprocess.Popen(
        ["uv", "run", "--with", "pywin32", "mcp_server.py"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=0,  # Line buffered
    )

    async def send_request(request):
        """Send JSON-RPC request and get response"""
        request_str = json.dumps(request) + "\n"
        if process.stdin:
            process.stdin.write(request_str)
            process.stdin.flush()

        response_line = process.stdout.readline() if process.stdout else ""
        if not response_line:
            return None

        return json.loads(response_line.strip())

    try:
        # Test 1: Initialize
        print("📡 Test 1: Initialize server...")
        init_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "test-client", "version": "1.0"},
            },
        }
        response = await send_request(init_request)
        if response and "result" in response:
            print("✅ Server initialized successfully")
            server_info = response["result"].get("serverInfo", {})
            print(f"   Server: {server_info.get('name')} v{server_info.get('version')}")
        else:
            print("❌ Failed to initialize")
            print(f"   Response: {response}")
            return

        # Test 2: List tools
        print("\n📋 Test 2: List available tools...")
        list_tools_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list",
        }
        response = await send_request(list_tools_request)
        if response and "result" in response:
            tools = response["result"].get("tools", [])
            print(f"✅ Found {len(tools)} tools:")
            for tool in tools[:5]:  # Show first 5
                print(f"   - {tool['name']}: {tool['description'][:60]}...")
            if len(tools) > 5:
                print(f"   ... and {len(tools) - 5} more")
        else:
            print("❌ Failed to list tools")
            print(f"   Response: {response}")
            return

        # Test 3: List emails (if Outlook is running)
        print("\n📧 Test 3: List emails (requires Outlook running)...")
        list_emails_request = {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {
                "name": "list_emails",
                "arguments": {"limit": 3},
            },
        }
        response = await send_request(list_emails_request)
        if response and "result" in response:
            content = response["result"].get("content", [])
            if content and content[0].get("type") == "text":
                emails = json.loads(content[0]["text"])
                if isinstance(emails, list) and len(emails) > 0:
                    print(f"✅ Found {len(emails)} emails:")
                    for email in emails[:2]:
                        print(f"   - {email.get('subject', '(No subject)')[:50]}")
                else:
                    print("ℹ️  No emails found (Outlook might not be running)")
            else:
                print("⚠️  Unexpected response format")
        else:
            print("ℹ️  Could not list emails (Outlook might not be running)")
            if "error" in response:
                print(f"   Error: {response['error'].get('message')}")

        # Test 4: List calendar events
        print("\n📅 Test 4: List calendar events...")
        list_calendar_request = {
            "jsonrpc": "2.0",
            "id": 4,
            "method": "tools/call",
            "params": {
                "name": "list_calendar_events",
                "arguments": {"days": 7},
            },
        }
        response = await send_request(list_calendar_request)
        if response and "result" in response:
            content = response["result"].get("content", [])
            if content and content[0].get("type") == "text":
                events = json.loads(content[0]["text"])
                if isinstance(events, list) and len(events) > 0:
                    print(f"✅ Found {len(events)} calendar events")
                else:
                    print("ℹ️  No events found")
            else:
                print("⚠️  Unexpected response format")
        else:
            print("ℹ️  Could not list events")
            if "error" in response:
                print(f"   Error: {response['error'].get('message')}")

        # Test 5: List tasks
        print("\n✅ Test 5: List tasks...")
        list_tasks_request = {
            "jsonrpc": "2.0",
            "id": 5,
            "method": "tools/call",
            "params": {
                "name": "list_tasks",
                "arguments": {},
            },
        }
        response = await send_request(list_tasks_request)
        if response and "result" in response:
            content = response["result"].get("content", [])
            if content and content[0].get("type") == "text":
                tasks = json.loads(content[0]["text"])
                if isinstance(tasks, list):
                    print(f"✅ Found {len(tasks)} tasks")
                else:
                    print("⚠️  Unexpected response format")
            else:
                print("⚠️  Unexpected response format")
        else:
            print("ℹ️  Could not list tasks")
            if "error" in response:
                print(f"   Error: {response['error'].get('message')}")

        print("\n" + "=" * 50)
        print("✅ MCP server test completed!")
        print("=" * 50)

    finally:
        # Clean up
        try:
            if process.stdin:
                process.stdin.close()
            process.terminate()
            process.wait(timeout=5)
        except:
            process.kill()


if __name__ == "__main__":
    try:
        asyncio.run(test_mcp_server())
    except KeyboardInterrupt:
        print("\n\n⚠️  Test interrupted")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ Test failed with error: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)
