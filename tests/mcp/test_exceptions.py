"""Tests for custom exception classes.

This module tests the custom exception classes that extend McpError
for better error handling in the MCP server.
"""

import pytest
from mcp import McpError

from mailtool.mcp.exceptions import (
    OutlookComError,
    OutlookNotFoundError,
    OutlookValidationError,
)


class TestOutlookNotFoundError:
    """Tests for OutlookNotFoundError class."""

    def test_is_mcp_error_subclass(self) -> None:
        """Test that OutlookNotFoundError extends McpError."""
        error = OutlookNotFoundError("Test error")
        assert isinstance(error, McpError)

    def test_message_only(self) -> None:
        """Test exception with message only."""
        error = OutlookNotFoundError("Item not found")
        assert str(error) == "Item not found"
        assert error.entry_id is None

    def test_message_with_entry_id(self) -> None:
        """Test exception with message and entry_id."""
        error = OutlookNotFoundError("Item not found", entry_id="ABC123")
        assert "Item not found" in str(error)
        assert "ABC123" in str(error)
        assert error.entry_id == "ABC123"

    def test_entry_id_in_message(self) -> None:
        """Test that entry_id is included in the exception message."""
        error = OutlookNotFoundError("Email not found", entry_id="0000000000000000")
        error_message = str(error)
        assert "Email not found" in error_message
        assert "EntryID: 0000000000000000" in error_message


class TestOutlookComError:
    """Tests for OutlookComError class."""

    def test_is_mcp_error_subclass(self) -> None:
        """Test that OutlookComError extends McpError."""
        error = OutlookComError("COM error")
        assert isinstance(error, McpError)

    def test_message_only(self) -> None:
        """Test exception with message only."""
        error = OutlookComError("Bridge not initialized")
        assert str(error) == "Bridge not initialized"
        assert error.details is None

    def test_message_with_details(self) -> None:
        """Test exception with message and details."""
        error = OutlookComError("COM error", details="Outlook not running")
        assert "COM error" in str(error)
        assert "Outlook not running" in str(error)
        assert error.details == "Outlook not running"

    def test_details_in_message(self) -> None:
        """Test that details are included in the exception message."""
        error = OutlookComError(
            "Bridge initialization failed", details="Timeout after 5 retries"
        )
        error_message = str(error)
        assert "Bridge initialization failed" in error_message
        assert "Timeout after 5 retries" in error_message

    def test_bridge_not_initialized_message(self) -> None:
        """Test common bridge not initialized error."""
        error = OutlookComError(
            "Outlook bridge not initialized. Is the server running?"
        )
        assert "Outlook bridge not initialized" in str(error)
        assert "server running" in str(error)


class TestOutlookValidationError:
    """Tests for OutlookValidationError class."""

    def test_is_mcp_error_subclass(self) -> None:
        """Test that OutlookValidationError extends McpError."""
        error = OutlookValidationError("Validation failed")
        assert isinstance(error, McpError)

    def test_message_only(self) -> None:
        """Test exception with message only."""
        error = OutlookValidationError("Invalid date format")
        assert "Validation failed" in str(error)
        assert "Invalid date format" in str(error)
        assert error.field is None

    def test_message_with_field(self) -> None:
        """Test exception with message and field."""
        error = OutlookValidationError("Invalid date format", field="start_date")
        error_message = str(error)
        assert "Validation failed" in error_message
        assert "start_date" in error_message
        assert "Invalid date format" in error_message
        assert error.field == "start_date"

    def test_field_in_message(self) -> None:
        """Test that field is included in the exception message."""
        error = OutlookValidationError(
            "Must be a valid email address", field="recipient"
        )
        error_message = str(error)
        assert "Validation failed" in error_message
        assert "'recipient'" in error_message
        assert "Must be a valid email address" in error_message

    def test_priority_validation(self) -> None:
        """Test validation error for priority field."""
        error = OutlookValidationError("Priority must be 0, 1, or 2", field="priority")
        error_message = str(error)
        assert "priority" in error_message
        assert "0, 1, or 2" in error_message

    def test_date_format_validation(self) -> None:
        """Test validation error for date format."""
        error = OutlookValidationError(
            "Date must be in YYYY-MM-DD format", field="due_date"
        )
        error_message = str(error)
        assert "due_date" in error_message
        assert "YYYY-MM-DD" in error_message


class TestExceptionAttributes:
    """Tests for exception attributes and behavior."""

    def test_all_exceptions_have_custom_attributes(self) -> None:
        """Test that all custom exceptions have their specific attributes."""
        not_found = OutlookNotFoundError("Test", entry_id="123")
        assert hasattr(not_found, "entry_id")
        assert not_found.entry_id == "123"

        com_error = OutlookComError("Test", details="Details")
        assert hasattr(com_error, "details")
        assert com_error.details == "Details"

        validation_error = OutlookValidationError("Test", field="field1")
        assert hasattr(validation_error, "field")
        assert validation_error.field == "field1"

    def test_exceptions_can_be_raised_and_caught(self) -> None:
        """Test that exceptions can be raised and caught properly."""
        # Test OutlookNotFoundError
        with pytest.raises(OutlookNotFoundError) as exc_info:
            raise OutlookNotFoundError("Not found", entry_id="ABC")
        assert exc_info.value.entry_id == "ABC"

        # Test OutlookComError
        with pytest.raises(OutlookComError) as exc_info:
            raise OutlookComError("COM failed", details="Timeout")
        assert exc_info.value.details == "Timeout"

        # Test OutlookValidationError
        with pytest.raises(OutlookValidationError) as exc_info:
            raise OutlookValidationError("Invalid", field="test")
        assert exc_info.value.field == "test"

    def test_exceptions_can_be_caught_as_mcp_error(self) -> None:
        """Test that all custom exceptions can be caught as McpError."""
        with pytest.raises(McpError):
            raise OutlookNotFoundError("Not found")

        with pytest.raises(McpError):
            raise OutlookComError("COM error")

        with pytest.raises(McpError):
            raise OutlookValidationError("Invalid")

    def test_exception_messages_are_descriptive(self) -> None:
        """Test that exception messages are descriptive and useful."""
        # OutlookNotFoundError with entry_id
        error1 = OutlookNotFoundError("Email not found", entry_id="ABC123")
        message1 = str(error1)
        assert "Email not found" in message1
        assert "ABC123" in message1

        # OutlookComError with details
        error2 = OutlookComError(
            "Bridge initialization failed", details="COM not available"
        )
        message2 = str(error2)
        assert "Bridge initialization failed" in message2
        assert "COM not available" in message2

        # OutlookValidationError with field
        error3 = OutlookValidationError("Invalid priority value", field="priority")
        message3 = str(error3)
        assert "priority" in message3
        assert "Invalid priority value" in message3
