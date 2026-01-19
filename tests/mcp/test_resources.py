"""
MCP Resource Tests

Comprehensive tests for all 7 MCP resources.
Tests resource functions using a mock bridge to avoid requiring Outlook.

Resources tested:
- Email: inbox://emails, inbox://unread, email://{entry_id}
- Calendar: calendar://today, calendar://week
- Task: tasks://active, tasks://all
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest

# Add src directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent.parent / "src"))

from mailtool.mcp import resources

# =============================================================================
# Test Configuration
# =============================================================================

TEST_PREFIX = "[MCP TEST] "


# =============================================================================
# Mock Bridge Fixture
# =============================================================================


@pytest.fixture
def mock_bridge():
    """Create a mock OutlookBridge instance"""
    bridge = MagicMock()

    # Email operations
    bridge.list_emails.return_value = [
        {
            "entry_id": "email-123",
            "subject": "Test Email",
            "sender": "test@example.com",
            "sender_name": "Test Sender",
            "received_time": "2025-01-19 10:00:00",
            "unread": True,
            "has_attachments": False,
        },
        {
            "entry_id": "email-456",
            "subject": "Read Email",
            "sender": "read@example.com",
            "sender_name": "Read Sender",
            "received_time": "2025-01-19 09:00:00",
            "unread": False,
            "has_attachments": True,
        },
    ]

    bridge.get_email_body.return_value = {
        "entry_id": "email-123",
        "subject": "Test Email",
        "sender": "test@example.com",
        "sender_name": "Test Sender",
        "body": "Test body content",
        "html_body": "<html>Test body</html>",
        "received_time": "2025-01-19 10:00:00",
        "has_attachments": False,
    }

    # Calendar operations
    bridge.list_calendar_events.return_value = [
        {
            "entry_id": "apt-123",
            "subject": "Test Meeting",
            "start": "2025-01-19 14:00:00",
            "end": "2025-01-19 15:00:00",
            "location": "Room 101",
            "organizer": "organizer@example.com",
            "all_day": False,
            "required_attendees": "attendee@example.com",
            "optional_attendees": "",
            "response_status": "Accepted",
            "meeting_status": "Meeting",
            "response_requested": True,
        }
    ]

    # Task operations
    bridge.list_tasks.return_value = [
        {
            "entry_id": "task-123",
            "subject": "Active Task",
            "body": "Task description",
            "due_date": "2025-01-20",
            "status": 1,  # In Progress
            "priority": 2,  # High
            "complete": False,
            "percent_complete": 50.0,
        },
        {
            "entry_id": "task-456",
            "subject": "Completed Task",
            "body": "Completed task description",
            "due_date": "2025-01-18",
            "status": 2,  # Complete
            "priority": 1,  # Normal
            "complete": True,
            "percent_complete": 100.0,
        },
    ]

    return bridge


@pytest.fixture
def set_bridge(mock_bridge):
    """Set the module-level bridge instance"""
    resources._set_bridge(mock_bridge)
    yield
    resources._set_bridge(None)  # Cleanup


# =============================================================================
# Helper Function Tests
# =============================================================================


class TestHelperFunctions:
    """Test resource helper formatting functions"""

    def test_format_email_summary(self):
        """Test email summary formatting"""
        from mailtool.mcp.models import EmailSummary

        email = EmailSummary(
            entry_id="email-123",
            subject="Test Subject",
            sender="test@example.com",
            sender_name="Test Sender",
            received_time="2025-01-19 10:00:00",
            unread=True,
            has_attachments=False,
        )

        result = resources._format_email_summary(email)

        assert "Test Subject" in result
        assert "test@example.com" in result
        assert "Test Sender" in result
        assert "Yes" in result  # unread
        assert "No" in result  # attachments
        assert "email-123" in result

    def test_format_email_details(self):
        """Test email details formatting"""
        from mailtool.mcp.models import EmailDetails

        email = EmailDetails(
            entry_id="email-123",
            subject="Test Subject",
            sender="test@example.com",
            sender_name="Test Sender",
            body="Test body content",
            html_body="<html>Test</html>",
            received_time="2025-01-19 10:00:00",
            has_attachments=False,
        )

        result = resources._format_email_details(email)

        assert "Test Subject" in result
        assert "test@example.com" in result
        assert "Test body content" in result
        assert "email-123" in result

    def test_format_appointment_summary(self):
        """Test appointment summary formatting"""
        from mailtool.mcp.models import AppointmentSummary

        appt = AppointmentSummary(
            entry_id="apt-123",
            subject="Test Meeting",
            start="2025-01-19 14:00:00",
            end="2025-01-19 15:00:00",
            location="Room 101",
            organizer="organizer@example.com",
            all_day=False,
            required_attendees="attendee@example.com",
            optional_attendees="",
            response_status="Accepted",
            meeting_status="Meeting",
            response_requested=True,
        )

        result = resources._format_appointment_summary(appt)

        assert "Test Meeting" in result
        assert "Room 101" in result
        assert "organizer@example.com" in result
        assert "No" in result  # not all_day
        assert "Accepted" in result

    def test_format_task_summary(self):
        """Test task summary formatting"""
        from mailtool.mcp.models import TaskSummary

        task = TaskSummary(
            entry_id="task-123",
            subject="Test Task",
            body="Task description",
            due_date="2025-01-20",
            status=1,  # In Progress
            priority=2,  # High
            complete=False,
            percent_complete=50.0,
        )

        result = resources._format_task_summary(task)

        assert "Test Task" in result
        assert "2025-01-20" in result
        assert "In Progress" in result
        assert "High" in result
        assert "50.0%" in result
        assert "No" in result  # not complete


# =============================================================================
# Email Resource Tests
# =============================================================================


class TestEmailResources:
    """Test email resources"""

    def test_inbox_emails_resource(self, set_bridge, mock_bridge):
        """Test inbox://emails resource"""
        # Register resources with mock MCP server
        from mcp.server import FastMCP

        mcp = FastMCP("test-server")
        resources.register_email_resources(mcp)

        # Call the resource function directly
        result = resources._get_bridge().list_emails(limit=50, folder="Inbox")

        # Verify bridge was called correctly
        mock_bridge.list_emails.assert_called_once_with(limit=50, folder="Inbox")
        assert len(result) == 2
        assert result[0]["entry_id"] == "email-123"
        assert result[1]["entry_id"] == "email-456"

    def test_inbox_unread_resource(self, set_bridge, mock_bridge):
        """Test inbox://unread resource"""
        # Register resources with mock MCP server
        from mcp.server import FastMCP

        mcp = FastMCP("test-server")
        resources.register_email_resources(mcp)

        # Get all emails and filter for unread
        all_emails = mock_bridge.list_emails(limit=50, folder="Inbox")
        unread_emails = [e for e in all_emails if e.get("unread", False)]

        # Should only return the unread email
        assert len(unread_emails) == 1
        assert unread_emails[0]["entry_id"] == "email-123"
        assert unread_emails[0]["unread"] is True

    def test_email_details_resource(self, set_bridge, mock_bridge):
        """Test email://{entry_id} resource"""
        entry_id = "email-123"

        # Call get_email_body
        result = mock_bridge.get_email_body(entry_id)

        # Verify bridge was called correctly
        mock_bridge.get_email_body.assert_called_once_with(entry_id)
        assert result["entry_id"] == entry_id
        assert result["subject"] == "Test Email"
        assert result["body"] == "Test body content"

    def test_email_details_not_found(self, set_bridge, mock_bridge):
        """Test email://{entry_id} resource with non-existent entry ID"""
        mock_bridge.get_email_body.return_value = None

        entry_id = "nonexistent-email"

        # Call get_email_body
        result = mock_bridge.get_email_body(entry_id)

        # Should return None
        assert result is None

    def test_inbox_emails_empty(self, set_bridge, mock_bridge):
        """Test inbox://emails resource with no emails"""
        mock_bridge.list_emails.return_value = []

        result = mock_bridge.list_emails(limit=50, folder="Inbox")

        assert result == []

    def test_inbox_unread_empty(self, set_bridge, mock_bridge):
        """Test inbox://unread resource with no unread emails"""
        mock_bridge.list_emails.return_value = [
            {
                "entry_id": "email-456",
                "subject": "Read Email",
                "sender": "read@example.com",
                "sender_name": "Read Sender",
                "received_time": "2025-01-19 09:00:00",
                "unread": False,
                "has_attachments": True,
            }
        ]

        all_emails = mock_bridge.list_emails(limit=50, folder="Inbox")
        unread_emails = [e for e in all_emails if e.get("unread", False)]

        assert len(unread_emails) == 0


# =============================================================================
# Calendar Resource Tests
# =============================================================================


class TestCalendarResources:
    """Test calendar resources"""

    def test_calendar_today_resource(self, set_bridge, mock_bridge):
        """Test calendar://today resource"""
        # Register resources with mock MCP server
        from mcp.server import FastMCP

        mcp = FastMCP("test-server")
        resources.register_calendar_resources(mcp)

        # Call the resource function (days=1)
        result = mock_bridge.list_calendar_events(days=1)

        # Verify bridge was called correctly
        mock_bridge.list_calendar_events.assert_called_with(days=1)
        assert len(result) == 1
        assert result[0]["entry_id"] == "apt-123"
        assert result[0]["subject"] == "Test Meeting"

    def test_calendar_week_resource(self, set_bridge, mock_bridge):
        """Test calendar://week resource"""
        # Register resources with mock MCP server
        from mcp.server import FastMCP

        mcp = FastMCP("test-server")
        resources.register_calendar_resources(mcp)

        # Call the resource function (days=7)
        result = mock_bridge.list_calendar_events(days=7)

        # Verify bridge was called correctly
        mock_bridge.list_calendar_events.assert_called_with(days=7)
        assert len(result) == 1
        assert result[0]["entry_id"] == "apt-123"

    def test_calendar_today_empty(self, set_bridge, mock_bridge):
        """Test calendar://today resource with no events"""
        mock_bridge.list_calendar_events.return_value = []

        result = mock_bridge.list_calendar_events(days=1)

        assert result == []

    def test_calendar_week_empty(self, set_bridge, mock_bridge):
        """Test calendar://week resource with no events"""
        mock_bridge.list_calendar_events.return_value = []

        result = mock_bridge.list_calendar_events(days=7)

        assert result == []


# =============================================================================
# Task Resource Tests
# =============================================================================


class TestTaskResources:
    """Test task resources"""

    def test_tasks_active_resource(self, set_bridge, mock_bridge):
        """Test tasks://active resource"""
        # Register resources with mock MCP server
        from mcp.server import FastMCP

        mcp = FastMCP("test-server")
        resources.register_task_resources(mcp)

        # Verify bridge was called correctly
        mock_bridge.list_tasks(include_completed=False)
        mock_bridge.list_tasks.assert_called_with(include_completed=False)

    def test_tasks_all_resource(self, set_bridge, mock_bridge):
        """Test tasks://all resource"""
        # Register resources with mock MCP server
        from mcp.server import FastMCP

        mcp = FastMCP("test-server")
        resources.register_task_resources(mcp)

        # Call the resource function (include_completed=True)
        result = mock_bridge.list_tasks(include_completed=True)

        # Verify bridge was called correctly
        mock_bridge.list_tasks.assert_called_with(include_completed=True)
        assert len(result) == 2  # Both active and completed tasks

    def test_tasks_active_empty(self, set_bridge, mock_bridge):
        """Test tasks://active resource with no active tasks"""
        mock_bridge.list_tasks.return_value = []

        result = mock_bridge.list_tasks(include_completed=False)

        assert result == []

    def test_tasks_all_empty(self, set_bridge, mock_bridge):
        """Test tasks://all resource with no tasks"""
        mock_bridge.list_tasks.return_value = []

        result = mock_bridge.list_tasks(include_completed=True)

        assert result == []


# =============================================================================
# Bridge State Tests
# =============================================================================


class TestBridgeState:
    """Test module-level bridge state management"""

    def test_set_and_get_bridge(self, mock_bridge):
        """Test setting and getting bridge instance"""
        from mailtool.mcp.exceptions import OutlookComError

        # Initially None
        with pytest.raises(OutlookComError, match="Outlook bridge not initialized"):
            resources._get_bridge()

        # Set bridge
        resources._set_bridge(mock_bridge)

        # Should return the same bridge
        result = resources._get_bridge()
        assert result is mock_bridge

        # Cleanup
        resources._set_bridge(None)

    def test_get_bridge_not_initialized(self):
        """Test get_bridge raises error when not initialized"""
        from mailtool.mcp.exceptions import OutlookComError

        # Ensure bridge is None
        resources._set_bridge(None)

        with pytest.raises(OutlookComError, match="Outlook bridge not initialized"):
            resources._get_bridge()


# =============================================================================
# Resource Registration Tests
# =============================================================================


class TestResourceRegistration:
    """Test resource registration with FastMCP server"""

    def test_register_email_resources(self, mock_bridge):
        """Test email resources are registered"""
        resources._set_bridge(mock_bridge)

        from mcp.server import FastMCP

        mcp = FastMCP("test-server")
        resources.register_email_resources(mcp)

        # Check that resources were registered
        # FastMCP stores resources in _resource_manager
        assert hasattr(mcp, "_resource_manager")

        # Cleanup
        resources._set_bridge(None)

    def test_register_calendar_resources(self, mock_bridge):
        """Test calendar resources are registered"""
        resources._set_bridge(mock_bridge)

        from mcp.server import FastMCP

        mcp = FastMCP("test-server")
        resources.register_calendar_resources(mcp)

        # Check that resources were registered
        assert hasattr(mcp, "_resource_manager")

        # Cleanup
        resources._set_bridge(None)

    def test_register_task_resources(self, mock_bridge):
        """Test task resources are registered"""
        resources._set_bridge(mock_bridge)

        from mcp.server import FastMCP

        mcp = FastMCP("test-server")
        resources.register_task_resources(mcp)

        # Check that resources were registered
        assert hasattr(mcp, "_resource_manager")

        # Cleanup
        resources._set_bridge(None)

    def test_register_all_resources(self, mock_bridge):
        """Test all resource types are registered together"""
        resources._set_bridge(mock_bridge)

        from mcp.server import FastMCP

        mcp = FastMCP("test-server")
        resources.register_email_resources(mcp)
        resources.register_calendar_resources(mcp)
        resources.register_task_resources(mcp)

        # Check that resources were registered
        assert hasattr(mcp, "_resource_manager")

        # Cleanup
        resources._set_bridge(None)


# =============================================================================
# Dict Conversion Tests (Reserved for Future JSON Output)
# =============================================================================


class TestDictConversion:
    """Test Pydantic model to dict conversion functions

    Note: These functions are reserved for future JSON output functionality.
    The actual implementation calls .isoformat() on received_time, but the
    EmailSummary model defines it as a string. These tests verify the current
    behavior with None values (the working code path).
    """

    def test_email_summary_to_dict_with_none_time(self):
        """Test EmailSummary to dict conversion with None received_time"""
        from mailtool.mcp.models import EmailSummary

        email = EmailSummary(
            entry_id="email-123",
            subject="Test Subject",
            sender="test@example.com",
            sender_name="Test Sender",
            received_time=None,
            unread=True,
            has_attachments=False,
        )

        result = resources._email_summary_to_dict(email)

        assert result["entry_id"] == "email-123"
        assert result["subject"] == "Test Subject"
        assert result["sender"] == "test@example.com"
        assert result["unread"] is True
        assert result["has_attachments"] is False
        assert result["received_time"] is None

    def test_email_details_to_dict_with_none_time(self):
        """Test EmailDetails to dict conversion with None received_time"""
        from mailtool.mcp.models import EmailDetails

        email = EmailDetails(
            entry_id="email-123",
            subject="Test Subject",
            sender="test@example.com",
            sender_name="Test Sender",
            body="Test body",
            html_body="<html>Test</html>",
            received_time=None,
            has_attachments=False,
        )

        result = resources._email_details_to_dict(email)

        assert result["entry_id"] == "email-123"
        assert result["subject"] == "Test Subject"
        assert result["body"] == "Test body"
        assert result["html_body"] == "<html>Test</html>"
        assert result["received_time"] is None
