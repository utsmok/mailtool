#!/usr/bin/env python3
"""
Mailtool MCP Server
Exposes Outlook automation functionality via Model Context Protocol
"""

import asyncio
import json
import sys
from typing import Any

from mailtool.bridge import OutlookBridge


# Server implementation using stdio transport
class MCPServer:
    """Simple MCP server implementation"""

    def __init__(self):
        self.bridge = None
        self.initialized = False

    async def handle_request(self, request: dict[str, Any]) -> dict[str, Any]:
        """Handle incoming MCP request"""
        method = request.get("method")

        if method == "initialize":
            return await self.initialize(request)
        elif method == "tools/list":
            return await self.list_tools()
        elif method == "tools/call":
            return await self.call_tool(request)
        elif method == "ping":
            return {"result": {}}
        else:
            return {
                "error": {
                    "code": -32601,
                    "message": f"Method not found: {method}",
                }
            }

    async def initialize(self, request: dict[str, Any]) -> dict[str, Any]:
        """Initialize the server and connect to Outlook"""
        try:
            self.bridge = OutlookBridge()
            self.initialized = True
            return {
                "result": {
                    "protocolVersion": "2024-11-05",
                    "serverInfo": {
                        "name": "mailtool-outlook-bridge",
                        "version": "2.1.0",
                    },
                    "capabilities": {
                        "tools": {},
                    },
                }
            }
        except Exception as e:
            return {
                "error": {
                    "code": -32603,
                    "message": f"Failed to initialize Outlook bridge: {e}",
                }
            }

    async def list_tools(self) -> dict[str, Any]:
        """List all available tools"""
        tools = [
            # Email tools
            {
                "name": "list_emails",
                "description": "List recent emails from inbox or another folder",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "limit": {
                            "type": "number",
                            "description": "Maximum number of emails to return (default: 10)",
                        },
                        "folder": {
                            "type": "string",
                            "description": "Folder name (default: Inbox)",
                        },
                    },
                },
            },
            {
                "name": "get_email",
                "description": "Get full email body and details by entry ID",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {
                            "type": "string",
                            "description": "Outlook EntryID of the email",
                        },
                    },
                    "required": ["entry_id"],
                },
            },
            {
                "name": "send_email",
                "description": "Send a new email or save as draft",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "to": {
                            "type": "string",
                            "description": "Recipient email address",
                        },
                        "subject": {"type": "string", "description": "Email subject"},
                        "body": {
                            "type": "string",
                            "description": "Email body (plain text)",
                        },
                        "cc": {
                            "type": "string",
                            "description": "CC recipients (optional)",
                        },
                        "bcc": {
                            "type": "string",
                            "description": "BCC recipients (optional)",
                        },
                        "html_body": {
                            "type": "string",
                            "description": "HTML body (optional)",
                        },
                        "save_draft": {
                            "type": "boolean",
                            "description": "Save as draft instead of sending",
                        },
                    },
                    "required": ["to", "subject", "body"],
                },
            },
            {
                "name": "reply_email",
                "description": "Reply to an email",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {
                            "type": "string",
                            "description": "Email entry ID to reply to",
                        },
                        "body": {"type": "string", "description": "Reply body"},
                        "reply_all": {
                            "type": "boolean",
                            "description": "Reply to all (default: false)",
                        },
                    },
                    "required": ["entry_id", "body"],
                },
            },
            {
                "name": "forward_email",
                "description": "Forward an email",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {
                            "type": "string",
                            "description": "Email entry ID to forward",
                        },
                        "to": {
                            "type": "string",
                            "description": "Recipient to forward to",
                        },
                        "body": {
                            "type": "string",
                            "description": "Optional additional body text",
                        },
                    },
                    "required": ["entry_id", "to"],
                },
            },
            {
                "name": "mark_email",
                "description": "Mark an email as read or unread",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {"type": "string", "description": "Email entry ID"},
                        "unread": {
                            "type": "boolean",
                            "description": "True to mark as unread, False to mark as read",
                        },
                    },
                    "required": ["entry_id"],
                },
            },
            {
                "name": "move_email",
                "description": "Move an email to a different folder",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {"type": "string", "description": "Email entry ID"},
                        "folder": {
                            "type": "string",
                            "description": "Target folder name (e.g., Archive, Deleted Items)",
                        },
                    },
                    "required": ["entry_id", "folder"],
                },
            },
            {
                "name": "delete_email",
                "description": "Delete an email",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {"type": "string", "description": "Email entry ID"},
                    },
                    "required": ["entry_id"],
                },
            },
            {
                "name": "search_emails",
                "description": "Search emails using Outlook filter query",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "filter_query": {
                            "type": "string",
                            "description": "SQL-like filter query (e.g., \"[Subject] = 'meeting'\")",
                        },
                        "limit": {
                            "type": "number",
                            "description": "Max results (default: 100)",
                        },
                    },
                    "required": ["filter_query"],
                },
            },
            # Calendar tools
            {
                "name": "list_calendar_events",
                "description": "List calendar events for the next N days or all events",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "days": {
                            "type": "number",
                            "description": "Number of days ahead to look (default: 7)",
                        },
                        "all": {
                            "type": "boolean",
                            "description": "Return all events without date filtering",
                        },
                    },
                },
            },
            {
                "name": "create_appointment",
                "description": "Create a new calendar appointment",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "subject": {
                            "type": "string",
                            "description": "Appointment subject",
                        },
                        "start": {
                            "type": "string",
                            "description": "Start time (YYYY-MM-DD HH:MM:SS)",
                        },
                        "end": {
                            "type": "string",
                            "description": "End time (YYYY-MM-DD HH:MM:SS)",
                        },
                        "location": {"type": "string", "description": "Location"},
                        "body": {
                            "type": "string",
                            "description": "Appointment description",
                        },
                        "all_day": {
                            "type": "boolean",
                            "description": "All-day event",
                        },
                        "required_attendees": {
                            "type": "string",
                            "description": "Semicolon-separated list",
                        },
                        "optional_attendees": {
                            "type": "string",
                            "description": "Semicolon-separated list",
                        },
                    },
                    "required": ["subject", "start", "end"],
                },
            },
            {
                "name": "get_appointment",
                "description": "Get full appointment details by entry ID",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {
                            "type": "string",
                            "description": "Appointment entry ID",
                        },
                    },
                    "required": ["entry_id"],
                },
            },
            {
                "name": "edit_appointment",
                "description": "Edit an existing appointment",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {
                            "type": "string",
                            "description": "Appointment entry ID",
                        },
                        "subject": {"type": "string", "description": "New subject"},
                        "start": {
                            "type": "string",
                            "description": "New start time (YYYY-MM-DD HH:MM:SS)",
                        },
                        "end": {
                            "type": "string",
                            "description": "New end time (YYYY-MM-DD HH:MM:SS)",
                        },
                        "location": {"type": "string", "description": "New location"},
                        "body": {"type": "string", "description": "New body"},
                        "required_attendees": {
                            "type": "string",
                            "description": "Comma-separated list",
                        },
                        "optional_attendees": {
                            "type": "string",
                            "description": "Comma-separated list",
                        },
                    },
                    "required": ["entry_id"],
                },
            },
            {
                "name": "respond_to_meeting",
                "description": "Respond to a meeting invitation",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {
                            "type": "string",
                            "description": "Appointment entry ID",
                        },
                        "response": {
                            "type": "string",
                            "enum": ["accept", "decline", "tentative"],
                            "description": "Meeting response",
                        },
                    },
                    "required": ["entry_id", "response"],
                },
            },
            {
                "name": "delete_appointment",
                "description": "Delete an appointment",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {
                            "type": "string",
                            "description": "Appointment entry ID",
                        },
                    },
                    "required": ["entry_id"],
                },
            },
            {
                "name": "get_free_busy",
                "description": "Get free/busy status for an email address",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "email_address": {
                            "type": "string",
                            "description": "Email address to check (default: current user)",
                        },
                        "start_date": {
                            "type": "string",
                            "description": "Start date (YYYY-MM-DD, default: today)",
                        },
                        "end_date": {
                            "type": "string",
                            "description": "End date (YYYY-MM-DD, default: tomorrow)",
                        },
                    },
                },
            },
            # Task tools
            {
                "name": "list_tasks",
                "description": "List incomplete tasks (only active tasks by default)",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "include_completed": {
                            "type": "boolean",
                            "description": "Include completed tasks in results (default: false)",
                        },
                    },
                },
            },
            {
                "name": "list_all_tasks",
                "description": "List all tasks including completed ones",
                "inputSchema": {
                    "type": "object",
                    "properties": {},
                },
            },
            {
                "name": "create_task",
                "description": "Create a new task",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "subject": {"type": "string", "description": "Task subject"},
                        "body": {"type": "string", "description": "Task description"},
                        "due_date": {
                            "type": "string",
                            "description": "Due date (YYYY-MM-DD)",
                        },
                        "importance": {
                            "type": "number",
                            "enum": [0, 1, 2],
                            "description": "0=Low, 1=Normal, 2=High",
                        },
                    },
                    "required": ["subject"],
                },
            },
            {
                "name": "get_task",
                "description": "Get full task details by entry ID",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {"type": "string", "description": "Task entry ID"},
                    },
                    "required": ["entry_id"],
                },
            },
            {
                "name": "edit_task",
                "description": "Edit an existing task",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {"type": "string", "description": "Task entry ID"},
                        "subject": {"type": "string", "description": "New subject"},
                        "body": {"type": "string", "description": "New body"},
                        "due_date": {
                            "type": "string",
                            "description": "New due date (YYYY-MM-DD)",
                        },
                        "importance": {
                            "type": "number",
                            "enum": [0, 1, 2],
                            "description": "0=Low, 1=Normal, 2=High",
                        },
                        "percent_complete": {
                            "type": "number",
                            "description": "Percent complete (0-100)",
                        },
                        "complete": {
                            "type": "boolean",
                            "description": "Mark complete/incomplete",
                        },
                    },
                    "required": ["entry_id"],
                },
            },
            {
                "name": "complete_task",
                "description": "Mark a task as complete",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {"type": "string", "description": "Task entry ID"},
                    },
                    "required": ["entry_id"],
                },
            },
            {
                "name": "delete_task",
                "description": "Delete a task",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "entry_id": {"type": "string", "description": "Task entry ID"},
                    },
                    "required": ["entry_id"],
                },
            },
        ]

        return {"result": {"tools": tools}}

    async def call_tool(self, request: dict[str, Any]) -> dict[str, Any]:
        """Call a specific tool"""
        if not self.initialized or not self.bridge:
            return {
                "error": {
                    "code": -32603,
                    "message": "Server not initialized",
                }
            }

        params = request.get("params", {})
        tool_name = params.get("name")
        arguments = params.get("arguments", {})

        try:
            # Email operations
            if tool_name == "list_emails":
                result = self.bridge.list_emails(
                    limit=arguments.get("limit", 10),
                    folder=arguments.get("folder", "Inbox"),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "get_email":
                result = self.bridge.get_email_body(arguments["entry_id"])
                if result:
                    return {
                        "result": {
                            "content": [{"type": "text", "text": json.dumps(result)}]
                        }
                    }
                else:
                    return {"error": {"code": -1, "message": "Email not found"}}

            elif tool_name == "send_email":
                result = self.bridge.send_email(
                    to=arguments["to"],
                    subject=arguments["subject"],
                    body=arguments["body"],
                    cc=arguments.get("cc"),
                    bcc=arguments.get("bcc"),
                    html_body=arguments.get("html_body"),
                    save_draft=arguments.get("save_draft", False),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "reply_email":
                result = self.bridge.reply_email(
                    entry_id=arguments["entry_id"],
                    body=arguments["body"],
                    reply_all=arguments.get("reply_all", False),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "forward_email":
                result = self.bridge.forward_email(
                    entry_id=arguments["entry_id"],
                    to=arguments["to"],
                    body=arguments.get("body", ""),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "mark_email":
                result = self.bridge.mark_email_read(
                    entry_id=arguments["entry_id"],
                    unread=arguments.get("unread", False),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "move_email":
                result = self.bridge.move_email(
                    entry_id=arguments["entry_id"],
                    folder_name=arguments["folder"],
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "delete_email":
                result = self.bridge.delete_email(arguments["entry_id"])
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "search_emails":
                result = self.bridge.search_emails(
                    filter_query=arguments["filter_query"],
                    limit=arguments.get("limit", 100),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            # Calendar operations
            elif tool_name == "list_calendar_events":
                result = self.bridge.list_calendar_events(
                    days=arguments.get("days", 7),
                    all_events=arguments.get("all", False),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "create_appointment":
                result = self.bridge.create_appointment(
                    subject=arguments["subject"],
                    start=arguments["start"],
                    end=arguments["end"],
                    location=arguments.get("location", ""),
                    body=arguments.get("body", ""),
                    all_day=arguments.get("all_day", False),
                    required_attendees=arguments.get("required_attendees"),
                    optional_attendees=arguments.get("optional_attendees"),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "get_appointment":
                result = self.bridge.get_appointment(arguments["entry_id"])
                if result:
                    return {
                        "result": {
                            "content": [{"type": "text", "text": json.dumps(result)}]
                        }
                    }
                else:
                    return {"error": {"code": -1, "message": "Appointment not found"}}

            elif tool_name == "edit_appointment":
                result = self.bridge.edit_appointment(
                    entry_id=arguments["entry_id"],
                    subject=arguments.get("subject"),
                    start=arguments.get("start"),
                    end=arguments.get("end"),
                    location=arguments.get("location"),
                    body=arguments.get("body"),
                    required_attendees=arguments.get("required_attendees"),
                    optional_attendees=arguments.get("optional_attendees"),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "respond_to_meeting":
                result = self.bridge.respond_to_meeting(
                    entry_id=arguments["entry_id"],
                    response=arguments["response"],
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "delete_appointment":
                result = self.bridge.delete_appointment(arguments["entry_id"])
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "get_free_busy":
                result = self.bridge.get_free_busy(
                    email_address=arguments.get("email_address"),
                    start_date=arguments.get("start_date"),
                    end_date=arguments.get("end_date"),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            # Task operations
            elif tool_name == "list_tasks":
                result = self.bridge.list_tasks(
                    include_completed=arguments.get("include_completed", False)
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "list_all_tasks":
                result = self.bridge.list_all_tasks()
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "create_task":
                result = self.bridge.create_task(
                    subject=arguments["subject"],
                    body=arguments.get("body", ""),
                    due_date=arguments.get("due_date"),
                    importance=arguments.get("importance", 1),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "get_task":
                result = self.bridge.get_task(arguments["entry_id"])
                if result:
                    return {
                        "result": {
                            "content": [{"type": "text", "text": json.dumps(result)}]
                        }
                    }
                else:
                    return {"error": {"code": -1, "message": "Task not found"}}

            elif tool_name == "edit_task":
                result = self.bridge.edit_task(
                    entry_id=arguments["entry_id"],
                    subject=arguments.get("subject"),
                    body=arguments.get("body"),
                    due_date=arguments.get("due_date"),
                    importance=arguments.get("importance"),
                    percent_complete=arguments.get("percent_complete"),
                    complete=arguments.get("complete"),
                )
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "complete_task":
                result = self.bridge.complete_task(arguments["entry_id"])
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            elif tool_name == "delete_task":
                result = self.bridge.delete_task(arguments["entry_id"])
                return {
                    "result": {
                        "content": [{"type": "text", "text": json.dumps(result)}]
                    }
                }

            else:
                return {
                    "error": {
                        "code": -32601,
                        "message": f"Unknown tool: {tool_name}",
                    }
                }

        except Exception as e:
            return {
                "error": {
                    "code": -32603,
                    "message": f"Error executing tool: {e}",
                }
            }


async def main():
    """Main server loop"""
    server = MCPServer()

    # Read from stdin and write to stdout
    while True:
        try:
            line = await asyncio.get_event_loop().run_in_executor(
                None, sys.stdin.readline
            )
            if not line:
                break

            # Remove any trailing newline
            line = line.strip()
            if not line:
                continue

            # Parse JSON-RPC request
            request = json.loads(line)

            # Handle request
            response = await server.handle_request(request)

            # Add request ID to response if present
            if "id" in request:
                response["id"] = request["id"]

            # Write response
            print(json.dumps(response), flush=True)

        except json.JSONDecodeError:
            error_response = {
                "error": {
                    "code": -32700,
                    "message": "Parse error",
                }
            }
            print(json.dumps(error_response), flush=True)
        except Exception as e:
            error_response = {
                "error": {
                    "code": -32603,
                    "message": f"Internal error: {e}",
                }
            }
            print(json.dumps(error_response), flush=True)


if __name__ == "__main__":
    asyncio.run(main())
