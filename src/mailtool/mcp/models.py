"""
Pydantic Models for MCP Tools

This module contains all Pydantic models for structured output from MCP tools.
Models provide type safety, automatic schema generation, and descriptive field metadata.
"""

from pydantic import BaseModel, Field

# ============================================================================
# Email Models (US-004)
# ============================================================================


class EmailSummary(BaseModel):
    """Summary representation of an email for list views"""

    entry_id: str = Field(description="Outlook EntryID for O(1) direct access")
    subject: str = Field(description="Email subject line")
    sender: str = Field(description="SMTP email address of sender")
    sender_name: str = Field(description="Display name of sender")
    received_time: str | None = Field(
        default=None,
        description="Received timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    unread: bool = Field(description="Whether the email is unread")
    has_attachments: bool = Field(description="Whether the email has attachments")


class EmailDetails(BaseModel):
    """Full email details including body content"""

    entry_id: str = Field(description="Outlook EntryID for O(1) direct access")
    subject: str = Field(description="Email subject line")
    sender: str = Field(description="SMTP email address of sender")
    sender_name: str = Field(description="Display name of sender")
    body: str = Field(description="Plain text email body")
    html_body: str = Field(description="HTML email body")
    received_time: str | None = Field(
        default=None,
        description="Received timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    has_attachments: bool = Field(description="Whether the email has attachments")


class SendEmailResult(BaseModel):
    """Result of sending an email or saving a draft"""

    success: bool = Field(description="Whether the operation succeeded")
    entry_id: str | None = Field(
        default=None, description="EntryID of saved draft (None if sent or failed)"
    )
    message: str = Field(description="Human-readable result message")


# ============================================================================
# Calendar Models (US-005)
# ============================================================================
# TODO: Implement AppointmentSummary, AppointmentDetails, CreateAppointmentResult, FreeBusyInfo


# ============================================================================
# Task Models (US-006)
# ============================================================================
# TODO: Implement TaskSummary, CreateTaskResult


# ============================================================================
# Common Result Models (US-007)
# ============================================================================
# TODO: Implement OperationResult
