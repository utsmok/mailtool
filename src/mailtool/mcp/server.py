"""Mailtool MCP Server

This module provides the main FastMCP server instance for Outlook automation.
It implements the Model Context Protocol (MCP) using the official MCP Python SDK v2
with the FastMCP framework.

The server provides 23 tools and 7 resources for Outlook email, calendar, and task management.
All tools return structured Pydantic models for type safety and LLM understanding.
"""

from typing import TYPE_CHECKING

from mcp import McpError
from mcp.server import FastMCP

from mailtool.mcp.lifespan import outlook_lifespan
from mailtool.mcp.models import AppointmentDetails, EmailDetails, TaskSummary

if TYPE_CHECKING:
    from mailtool.bridge import OutlookBridge

# Create FastMCP server instance
# The lifespan parameter manages Outlook COM bridge lifecycle (creation, warmup, cleanup)
mcp = FastMCP(
    name="mailtool-outlook-bridge",
    lifespan=outlook_lifespan,
)

# Module-level bridge instance (set by lifespan, accessed by tools)
_bridge: "OutlookBridge | None" = None


def _get_bridge():
    """Get the current bridge instance

    Returns:
        OutlookBridge: The bridge instance

    Raises:
        McpError: If bridge is not initialized (server not running)
    """
    global _bridge
    if _bridge is None:
        raise McpError("Outlook bridge not initialized. Is the server running?")
    return _bridge


# ============================================================================
# Email Tools (US-008: get_email)
# ============================================================================


@mcp.tool()
def get_email(entry_id: str) -> EmailDetails:
    """
    Get full email body and details by entry ID.

    Retrieves complete email information including body content (both plain text
    and HTML) using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)

    Returns:
        EmailDetails: Complete email details including body content

    Raises:
        McpError: If email not found or cannot be accessed
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Get email body from bridge
    result = bridge.get_email_body(entry_id)

    # Check if email was found
    if result is None:
        raise McpError(f"Email not found: {entry_id}")

    # Convert bridge result to EmailDetails model
    # Note: EmailDetails doesn't have 'unread' field (bridge.get_email_body doesn't return it)
    return EmailDetails(
        entry_id=result["entry_id"],
        subject=result["subject"],
        sender=result["sender"],
        sender_name=result["sender_name"],
        body=result["body"],
        html_body=result["html_body"],
        received_time=result["received_time"],
        has_attachments=result["has_attachments"],
    )


# ============================================================================
# Calendar Tools (US-009: get_appointment)
# ============================================================================


@mcp.tool()
def get_appointment(entry_id: str) -> AppointmentDetails:
    """
    Get full appointment details and body by entry ID.

    Retrieves complete appointment information including body content and all
    meeting metadata using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the appointment (O(1) direct access)

    Returns:
        AppointmentDetails: Complete appointment details including body content

    Raises:
        McpError: If appointment not found or cannot be accessed
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Get appointment details from bridge
    result = bridge.get_appointment(entry_id)

    # Check if appointment was found
    if result is None:
        raise McpError(f"Appointment not found: {entry_id}")

    # Convert bridge result to AppointmentDetails model
    # Note: AppointmentDetails extends AppointmentSummary, adding 'body' field
    return AppointmentDetails(
        entry_id=result["entry_id"],
        subject=result["subject"],
        start=result["start"],
        end=result["end"],
        location=result["location"],
        organizer=result["organizer"],
        body=result["body"],
        all_day=result["all_day"],
        required_attendees=result["required_attendees"],
        optional_attendees=result["optional_attendees"],
        response_status=result["response_status"],
        meeting_status=result["meeting_status"],
        response_requested=result["response_requested"],
    )


# ============================================================================
# Task Tools (US-010: get_task)
# ============================================================================


@mcp.tool()
def get_task(entry_id: str) -> TaskSummary:
    """
    Get full task details and body by entry ID.

    Retrieves complete task information including body content and all
    task metadata using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the task (O(1) direct access)

    Returns:
        TaskSummary: Complete task details including body content

    Raises:
        McpError: If task not found or cannot be accessed
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Get task details from bridge
    result = bridge.get_task(entry_id)

    # Check if task was found
    if result is None:
        raise McpError(f"Task not found: {entry_id}")

    # Convert bridge result to TaskSummary model
    # Note: TaskSummary includes all fields from bridge.get_task()
    return TaskSummary(
        entry_id=result["entry_id"],
        subject=result["subject"],
        body=result["body"],
        due_date=result["due_date"],
        status=result["status"],
        priority=result["priority"],
        complete=result["complete"],
        percent_complete=result["percent_complete"],
    )


if __name__ == "__main__":
    # Run the MCP server with stdio transport
    # This is the standard transport for MCP clients like Claude Code
    mcp.run(transport="stdio")
