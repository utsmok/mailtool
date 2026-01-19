"""MCP Resources

This module provides MCP resource implementations for read-only data access.
Resources are registered with the FastMCP server and return formatted text/JSON content.

Email resources:
- inbox://emails - List recent emails from inbox
- inbox://unread - List unread emails from inbox
- email://{entry_id} - Get full email details by entry ID (template resource)

Calendar resources will be added in US-028
Task resources will be added in US-033
"""

from typing import TYPE_CHECKING

from mcp.server import FastMCP

from mailtool.mcp.models import EmailDetails, EmailSummary

if TYPE_CHECKING:
    from mailtool.bridge import OutlookBridge

# Module-level bridge instance (set by server.py)
_bridge: "OutlookBridge | None" = None


def _get_bridge() -> "OutlookBridge":
    """Get the current bridge instance.

    Returns:
        OutlookBridge: The bridge instance

    Raises:
        RuntimeError: If bridge is not initialized
    """
    global _bridge
    if _bridge is None:
        raise RuntimeError("Outlook bridge not initialized")
    return _bridge


def _set_bridge(bridge: "OutlookBridge") -> None:
    """Set the bridge instance (called by server.py lifespan).

    Args:
        bridge: The bridge instance
    """
    global _bridge
    _bridge = bridge


def _format_email_summary(email: EmailSummary) -> str:
    """Format an email summary as readable text.

    Args:
        email: EmailSummary model

    Returns:
        Formatted text representation
    """
    return f"""Subject: {email.subject}
From: {email.sender_name} <{email.sender}>
Received: {email.received_time}
Unread: {"Yes" if email.unread else "No"}
Attachments: {"Yes" if email.has_attachments else "No"}
Entry ID: {email.entry_id}
"""


def _format_email_details(email: EmailDetails) -> str:
    """Format full email details as readable text.

    Args:
        email: EmailDetails model

    Returns:
        Formatted text representation with body
    """
    return f"""Subject: {email.subject}
From: {email.sender_name} <{email.sender}>
Received: {email.received_time}
Attachments: {"Yes" if email.has_attachments else "No"}
Entry ID: {email.entry_id}

Body:
{email.body}
"""


def _email_summary_to_dict(email: EmailSummary) -> dict:
    """Convert EmailSummary to dict for JSON serialization.

    Args:
        email: EmailSummary model

    Returns:
        Dictionary representation
    """
    return {
        "entry_id": email.entry_id,
        "subject": email.subject,
        "sender": email.sender,
        "sender_name": email.sender_name,
        "received_time": email.received_time.isoformat()
        if email.received_time
        else None,
        "unread": email.unread,
        "has_attachments": email.has_attachments,
    }


def _email_details_to_dict(email: EmailDetails) -> dict:
    """Convert EmailDetails to dict for JSON serialization.

    Args:
        email: EmailDetails model

    Returns:
        Dictionary representation
    """
    return {
        "entry_id": email.entry_id,
        "subject": email.subject,
        "sender": email.sender,
        "sender_name": email.sender_name,
        "body": email.body,
        "html_body": email.html_body,
        "received_time": email.received_time.isoformat()
        if email.received_time
        else None,
        "has_attachments": email.has_attachments,
    }


def register_email_resources(mcp: FastMCP) -> None:
    """Register all email resources with the FastMCP server.

    Args:
        mcp: FastMCP server instance

    This function registers three email resources:
    1. inbox://emails - Lists recent emails (limit 50)
    2. inbox://unread - Lists unread emails (limit 50)
    3. email://{entry_id} - Gets full email details (template resource)
    """

    @mcp.resource(
        uri="inbox://emails",
        name="inbox_emails",
        title="Inbox Emails",
        description="List recent emails from the inbox (max 50)",
    )
    def inbox_emails() -> str:
        """Get recent emails from inbox.

        Returns:
            Formatted text with email summaries
        """
        bridge = _get_bridge()

        # Get emails from bridge
        emails_data = bridge.list_emails(limit=50, folder="Inbox")

        # Convert to EmailSummary models
        emails = [
            EmailSummary(
                entry_id=email["entry_id"],
                subject=email["subject"],
                sender=email["sender"],
                sender_name=email["sender_name"],
                received_time=email["received_time"],
                unread=email["unread"],
                has_attachments=email["has_attachments"],
            )
            for email in emails_data
        ]

        # Format as text
        if not emails:
            return "No emails found in inbox"

        lines = [f"Inbox Emails ({len(emails)} items)", ""]
        for email in emails:
            lines.append(_format_email_summary(email))
            lines.append("-" * 60)

        return "\n".join(lines)

    @mcp.resource(
        uri="inbox://unread",
        name="inbox_unread",
        title="Unread Inbox Emails",
        description="List unread emails from the inbox (max 50)",
    )
    def inbox_unread() -> str:
        """Get unread emails from inbox.

        Returns:
            Formatted text with unread email summaries
        """
        bridge = _get_bridge()

        # Get all emails from bridge
        emails_data = bridge.list_emails(limit=50, folder="Inbox")

        # Filter to only unread emails
        unread_emails_data = [
            email for email in emails_data if email.get("unread", False)
        ]

        # Convert to EmailSummary models
        emails = [
            EmailSummary(
                entry_id=email["entry_id"],
                subject=email["subject"],
                sender=email["sender"],
                sender_name=email["sender_name"],
                received_time=email["received_time"],
                unread=email["unread"],
                has_attachments=email["has_attachments"],
            )
            for email in unread_emails_data
        ]

        # Format as text
        if not emails:
            return "No unread emails in inbox"

        lines = [f"Unread Emails ({len(emails)} items)", ""]
        for email in emails:
            lines.append(_format_email_summary(email))
            lines.append("-" * 60)

        return "\n".join(lines)

    @mcp.resource(
        uri="email://{entry_id}",
        name="email_details",
        title="Email Details",
        description="Get full email details by entry ID",
    )
    def email_details(entry_id: str) -> str:
        """Get full email details by entry ID.

        Args:
            entry_id: Outlook EntryID of the email

        Returns:
            Formatted text with full email details including body

        Raises:
            RuntimeError: If email not found
        """
        bridge = _get_bridge()

        # Get email details from bridge
        email_data = bridge.get_email_body(entry_id)

        if email_data is None:
            return f"Email not found: {entry_id}"

        # Convert to EmailDetails model
        email = EmailDetails(
            entry_id=email_data["entry_id"],
            subject=email_data["subject"],
            sender=email_data["sender"],
            sender_name=email_data["sender_name"],
            body=email_data["body"],
            html_body=email_data["html_body"],
            received_time=email_data["received_time"],
            has_attachments=email_data["has_attachments"],
        )

        # Format as text
        return _format_email_details(email)
