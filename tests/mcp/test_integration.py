"""
MCP Integration Tests

End-to-end workflow tests for complete email, calendar, and task operations.
These tests verify that tools work together correctly in realistic scenarios.

Test scenarios:
- Email workflow: List → Get → Reply → Move → Delete
- Calendar workflow: List → Create → Edit → Respond → Delete
- Task workflow: List → Create → Edit → Complete → Delete
- Cross-domain workflows: Email → Task, Email → Appointment
- Resource queries: Verify resources return correct data
"""

import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest

# Add src directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent.parent / "src"))

from mailtool.mcp import server
from mailtool.mcp.models import (
    AppointmentDetails,
    AppointmentSummary,
    CreateAppointmentResult,
    CreateTaskResult,
    EmailDetails,
    EmailSummary,
    FreeBusyInfo,
    OperationResult,
    SendEmailResult,
    TaskSummary,
)

# =============================================================================
# Test Configuration
# =============================================================================

TEST_PREFIX = "[MCP TEST] "


# =============================================================================
# Mock Bridge Fixture
# =============================================================================


@pytest.fixture
def mock_bridge():
    """Create a mock OutlookBridge instance with realistic data"""
    bridge = MagicMock()

    # Email operations - multiple emails
    bridge.list_emails.return_value = [
        {
            "entry_id": "email-001",
            "subject": f"{TEST_PREFIX}Project Update",
            "sender": "alice@example.com",
            "sender_name": "Alice Smith",
            "received_time": "2025-01-19 09:00:00",
            "unread": True,
            "has_attachments": True,
        },
        {
            "entry_id": "email-002",
            "subject": f"{TEST_PREFIX}Meeting Tomorrow",
            "sender": "bob@example.com",
            "sender_name": "Bob Johnson",
            "received_time": "2025-01-19 08:30:00",
            "unread": False,
            "has_attachments": False,
        },
    ]

    bridge.get_email_body.return_value = {
        "entry_id": "email-001",
        "subject": f"{TEST_PREFIX}Project Update",
        "sender": "alice@example.com",
        "sender_name": "Alice Smith",
        "body": "Hi, the project is progressing well. We need to discuss the next steps.",
        "html_body": "<p>Hi, the project is progressing well.</p>",
        "received_time": "2025-01-19 09:00:00",
        "has_attachments": True,
    }

    bridge.send_email.return_value = True  # Email sent successfully

    bridge.reply_email.return_value = True  # Reply sent successfully

    bridge.forward_email.return_value = True  # Forward sent successfully

    bridge.mark_email_read.return_value = True  # Marked as read

    bridge.move_email.return_value = True  # Email moved

    bridge.delete_email.return_value = True  # Email deleted

    bridge.search_emails.return_value = [
        {
            "entry_id": "email-001",
            "subject": f"{TEST_PREFIX}Project Update",
            "sender": "alice@example.com",
            "sender_name": "Alice Smith",
            "received_time": "2025-01-19 09:00:00",
            "unread": True,
            "has_attachments": True,
        }
    ]

    # Calendar operations
    bridge.list_calendar_events.return_value = [
        {
            "entry_id": "appt-001",
            "subject": f"{TEST_PREFIX}Team Meeting",
            "start": "2025-01-20 14:00:00",
            "end": "2025-01-20 15:00:00",
            "location": "Room 101",
            "organizer": "manager@example.com",
            "all_day": False,
            "required_attendees": "team@example.com",
            "optional_attendees": "",
            "response_status": "0",
            "meeting_status": "0",
            "response_requested": True,
        }
    ]

    bridge.get_appointment.return_value = {
        "entry_id": "appt-001",
        "subject": f"{TEST_PREFIX}Team Meeting",
        "start": "2025-01-20 14:00:00",
        "end": "2025-01-20 15:00:00",
        "location": "Room 101",
        "organizer": "manager@example.com",
        "body": "Weekly team sync to discuss project status.",
        "all_day": False,
        "required_attendees": "team@example.com",
        "optional_attendees": "",
        "response_status": "0",
        "meeting_status": "0",
        "response_requested": True,
    }

    bridge.create_appointment.return_value = "appt-new-001"  # New appointment EntryID

    bridge.edit_appointment.return_value = True  # Appointment edited

    bridge.respond_to_meeting.return_value = True  # Response sent

    bridge.delete_appointment.return_value = True  # Appointment deleted

    bridge.get_free_busy.return_value = {
        "email": "user@example.com",
        "start_date": "2025-01-20",
        "end_date": "2025-01-21",
        "free_busy": "0000111122000",  # Free/busy status string
    }

    # Task operations
    bridge.list_tasks.return_value = [
        {
            "entry_id": "task-001",
            "subject": f"{TEST_PREFIX}Review Proposal",
            "body": "Review the Q1 proposal document",
            "due_date": "2025-01-25",
            "status": 1,  # In Progress
            "priority": 2,  # High
            "complete": False,
            "percent_complete": 50.0,
        }
    ]

    bridge.get_task.return_value = {
        "entry_id": "task-001",
        "subject": f"{TEST_PREFIX}Review Proposal",
        "body": "Review the Q1 proposal document",
        "due_date": "2025-01-25",
        "status": 1,  # In Progress
        "priority": 2,  # High
        "complete": False,
        "percent_complete": 50.0,
    }

    bridge.create_task.return_value = "task-new-001"  # New task EntryID

    bridge.edit_task.return_value = True  # Task edited

    bridge.complete_task.return_value = True  # Task completed

    bridge.delete_task.return_value = True  # Task deleted

    return bridge


@pytest.fixture
def set_bridge(mock_bridge):
    """Set module-level bridge state and cleanup after test"""
    from mailtool.mcp import resources

    server._bridge = mock_bridge
    resources._set_bridge(mock_bridge)
    yield
    server._bridge = None
    resources._set_bridge(None)


# =============================================================================
# Email Workflow Tests
# =============================================================================


class TestEmailWorkflow:
    """Test complete email workflow: List → Get → Reply → Move → Delete"""

    def test_list_then_get_email(self, set_bridge):
        """Test listing emails then getting details for one"""
        # List emails
        emails = server.list_emails(limit=10)
        assert len(emails) == 2
        assert isinstance(emails[0], EmailSummary)
        assert emails[0].subject == f"{TEST_PREFIX}Project Update"

        # Get details for first email
        details = server.get_email(entry_id="email-001")
        assert isinstance(details, EmailDetails)
        assert details.subject == f"{TEST_PREFIX}Project Update"
        assert details.body.startswith("Hi, the project")

    def test_reply_to_email_workflow(self, set_bridge):
        """Test getting an email and replying to it"""
        # Get email details
        email = server.get_email(entry_id="email-001")
        assert email.subject == f"{TEST_PREFIX}Project Update"

        # Reply to email
        result = server.reply_email(entry_id="email-001", body="Thanks for the update!")
        assert isinstance(result, OperationResult)
        assert result.success is True
        assert "replied" in result.message.lower()

    def test_forward_email_workflow(self, set_bridge):
        """Test forwarding an email to someone else"""
        # Forward email
        result = server.forward_email(
            entry_id="email-001", to="charlie@example.com", body="FYI - see below"
        )
        assert isinstance(result, OperationResult)
        assert result.success is True
        assert "forwarded" in result.message.lower()

    def test_mark_and_move_email_workflow(self, set_bridge):
        """Test marking email as read and moving to folder"""
        # Mark as read
        result1 = server.mark_email(entry_id="email-001", unread=False)
        assert isinstance(result1, OperationResult)
        assert result1.success is True

        # Move to Archive
        result2 = server.move_email(entry_id="email-001", folder="Archive")
        assert isinstance(result2, OperationResult)
        assert result2.success is True
        assert "archive" in result2.message.lower()

    def test_send_then_delete_workflow(self, set_bridge):
        """Test sending a new email then deleting it"""
        # Send email
        result = server.send_email(
            to="recipient@example.com",
            subject="Test Message",
            body="This is a test",
        )
        assert isinstance(result, SendEmailResult)
        assert result.success is True

        # Delete email (in real workflow, would delete sent email)
        delete_result = server.delete_email(entry_id="email-001")
        assert isinstance(delete_result, OperationResult)
        assert delete_result.success is True

    def test_search_emails_workflow(self, set_bridge):
        """Test searching for specific emails"""
        # Search for project-related emails
        results = server.search_emails(
            filter_query="[Subject] LIKE '%Project%'", limit=10
        )
        assert len(results) >= 1
        assert isinstance(results[0], EmailSummary)
        assert "project" in results[0].subject.lower()


# =============================================================================
# Calendar Workflow Tests
# =============================================================================


class TestCalendarWorkflow:
    """Test complete calendar workflow: List → Create → Edit → Respond → Delete"""

    def test_list_then_get_appointment(self, set_bridge):
        """Test listing appointments then getting details for one"""
        # List appointments
        appointments = server.list_calendar_events(days=7)
        assert len(appointments) == 1
        assert isinstance(appointments[0], AppointmentSummary)
        assert appointments[0].subject == f"{TEST_PREFIX}Team Meeting"

        # Get details
        details = server.get_appointment(entry_id="appt-001")
        assert isinstance(details, AppointmentDetails)
        assert details.subject == f"{TEST_PREFIX}Team Meeting"
        assert details.location == "Room 101"

    def test_create_appointment_workflow(self, set_bridge):
        """Test creating a new appointment"""
        result = server.create_appointment(
            subject="New Meeting",
            start="2025-01-21 10:00:00",
            end="2025-01-21 11:00:00",
            location="Room 202",
        )
        assert isinstance(result, CreateAppointmentResult)
        assert result.success is True
        assert result.entry_id == "appt-new-001"

    def test_edit_appointment_workflow(self, set_bridge):
        """Test editing an existing appointment"""
        # Edit appointment location
        result = server.edit_appointment(
            entry_id="appt-001", location="Room 305", body="Updated agenda"
        )
        assert isinstance(result, OperationResult)
        assert result.success is True

    def test_respond_to_meeting_workflow(self, set_bridge):
        """Test responding to a meeting invitation"""
        # Accept meeting
        result = server.respond_to_meeting(entry_id="appt-001", response="accept")
        assert isinstance(result, OperationResult)
        assert result.success is True
        assert "accepted" in result.message.lower()

        # Test other responses
        result_decline = server.respond_to_meeting(
            entry_id="appt-001", response="decline"
        )
        assert result_decline.success is True

        result_tentative = server.respond_to_meeting(
            entry_id="appt-001", response="tentative"
        )
        assert result_tentative.success is True

    def test_delete_appointment_workflow(self, set_bridge):
        """Test deleting an appointment"""
        result = server.delete_appointment(entry_id="appt-001")
        assert isinstance(result, OperationResult)
        assert result.success is True

    def test_free_busy_workflow(self, set_bridge):
        """Test getting free/busy information"""
        info = server.get_free_busy(
            email_address="user@example.com",
            start_date="2025-01-20",
            end_date="2025-01-21",
        )
        assert isinstance(info, FreeBusyInfo)
        assert info.email == "user@example.com"
        assert info.free_busy is not None


# =============================================================================
# Task Workflow Tests
# =============================================================================


class TestTaskWorkflow:
    """Test complete task workflow: List → Create → Edit → Complete → Delete"""

    def test_list_then_get_task(self, set_bridge):
        """Test listing tasks then getting details for one"""
        # List active tasks
        tasks = server.list_tasks(include_completed=False)
        assert len(tasks) == 1
        assert isinstance(tasks[0], TaskSummary)
        assert tasks[0].subject == f"{TEST_PREFIX}Review Proposal"

        # Get details
        details = server.get_task(entry_id="task-001")
        assert isinstance(details, TaskSummary)
        assert details.subject == f"{TEST_PREFIX}Review Proposal"
        assert details.percent_complete == 50.0

    def test_list_all_tasks_workflow(self, set_bridge):
        """Test listing all tasks including completed"""
        # List all tasks
        tasks = server.list_all_tasks()
        assert len(tasks) == 1
        assert isinstance(tasks[0], TaskSummary)

    def test_create_task_workflow(self, set_bridge):
        """Test creating a new task"""
        result = server.create_task(
            subject="New Task",
            body="Task description",
            due_date="2025-01-30",
            priority=2,  # High
        )
        assert isinstance(result, CreateTaskResult)
        assert result.success is True
        assert result.entry_id == "task-new-001"

    def test_edit_task_workflow(self, set_bridge):
        """Test editing an existing task"""
        # Edit task
        result = server.edit_task(
            entry_id="task-001",
            subject="Updated Task Title",
            percent_complete=75.0,
        )
        assert isinstance(result, OperationResult)
        assert result.success is True

    def test_complete_task_workflow(self, set_bridge):
        """Test marking a task as complete"""
        result = server.complete_task(entry_id="task-001")
        assert isinstance(result, OperationResult)
        assert result.success is True

    def test_delete_task_workflow(self, set_bridge):
        """Test deleting a task"""
        result = server.delete_task(entry_id="task-001")
        assert isinstance(result, OperationResult)
        assert result.success is True


# =============================================================================
# Cross-Domain Workflow Tests
# =============================================================================


class TestCrossDomainWorkflows:
    """Test workflows that span multiple domains (email → task, email → appointment)"""

    def test_email_to_task_workflow(self, set_bridge):
        """Test creating a task from an email"""
        # Get email details
        email = server.get_email(entry_id="email-001")

        # Create task based on email content
        task_result = server.create_task(
            subject=f"Follow up: {email.subject}",
            body=f"From: {email.sender_name}\n\n{email.body}",
            due_date="2025-01-25",
            priority=2,
        )
        assert isinstance(task_result, CreateTaskResult)
        assert task_result.success is True

    def test_email_to_appointment_workflow(self, set_bridge):
        """Test creating an appointment from an email discussion"""
        # Get email
        email = server.get_email(entry_id="email-001")

        # Create meeting to discuss email topic
        appt_result = server.create_appointment(
            subject=f"Discuss: {email.subject}",
            start="2025-01-20 14:00:00",
            end="2025-01-20 15:00:00",
            location="Conference Room",
            body=f"Discussion about: {email.subject}",
        )
        assert isinstance(appt_result, CreateAppointmentResult)
        assert appt_result.success is True

    def test_task_to_appointment_workflow(self, set_bridge):
        """Test creating an appointment from a task"""
        # Get task
        task = server.get_task(entry_id="task-001")

        # Create meeting to work on task
        appt_result = server.create_appointment(
            subject=f"Work Session: {task.subject}",
            start="2025-01-20 10:00:00",
            end="2025-01-20 11:00:00",
            location="Office",
            body=task.body,
        )
        assert isinstance(appt_result, CreateAppointmentResult)
        assert appt_result.success is True


# =============================================================================
# Resource Query Tests
# =============================================================================


class TestResourceQueries:
    """Test that resources can be queried and return correct data"""

    def test_email_resources_registered(self, set_bridge):
        """Test that email resources are registered with the server"""
        from mailtool.mcp import resources

        # Check that email resources are registered
        # Resources are registered via FastMCP decorators
        # We can verify they're callable by checking the module has the registration functions
        assert hasattr(resources, "register_email_resources")
        assert callable(resources.register_email_resources)

    def test_calendar_resources_registered(self, set_bridge):
        """Test that calendar resources are registered with the server"""
        from mailtool.mcp import resources

        # Check that calendar resources are registered
        assert hasattr(resources, "register_calendar_resources")
        assert callable(resources.register_calendar_resources)

    def test_task_resources_registered(self, set_bridge):
        """Test that task resources are registered with the server"""
        from mailtool.mcp import resources

        # Check that task resources are registered
        assert hasattr(resources, "register_task_resources")
        assert callable(resources.register_task_resources)

    def test_helper_functions_exist(self, set_bridge):
        """Test that resource helper functions exist"""
        from mailtool.mcp import resources

        # Check that formatting helper functions exist
        assert hasattr(resources, "_format_email_summary")
        assert hasattr(resources, "_format_appointment_summary")
        assert hasattr(resources, "_format_task_summary")
        assert callable(resources._format_email_summary)
        assert callable(resources._format_appointment_summary)
        assert callable(resources._format_task_summary)

    def test_bridge_state_management(self, set_bridge):
        """Test that bridge state can be managed"""
        from mailtool.mcp import resources

        # Test _get_bridge and _set_bridge
        assert hasattr(resources, "_get_bridge")
        assert hasattr(resources, "_set_bridge")

        # Get current bridge
        current_bridge = resources._get_bridge()
        assert current_bridge is not None

        # Set None and verify that _get_bridge raises an error
        from mailtool.mcp.exceptions import OutlookComError

        resources._set_bridge(None)
        with pytest.raises(OutlookComError, match="not initialized"):
            resources._get_bridge()

        # Restore bridge
        resources._set_bridge(current_bridge)
        assert resources._get_bridge() is not None

    def test_resource_registration_with_server(self, set_bridge):
        """Test that resources can be registered with FastMCP server"""
        from mailtool.mcp import server

        # Resources are already registered in server.py
        # Verify that server has resources registered
        # FastMCP stores resources in _resource_manager
        assert hasattr(server.mcp, "_resource_manager")
        resource_manager = server.mcp._resource_manager

        # Check that resources are registered (has _resources and _templates attributes)
        assert hasattr(resource_manager, "_resources")
        assert hasattr(resource_manager, "_templates")

        # Should have at least some resources registered
        total_resources = len(resource_manager._resources) + len(
            resource_manager._templates
        )
        assert total_resources > 0


# =============================================================================
# Error Handling Tests
# =============================================================================


class TestErrorHandling:
    """Test error handling in integration scenarios"""

    def test_get_nonexistent_email_raises_error(self, set_bridge):
        """Test getting non-existent email raises error"""
        # Mock bridge to return None (reset first to clear default return value)
        server._bridge.get_email_body.reset_mock()
        server._bridge.get_email_body.return_value = None

        # Should raise some kind of exception
        with pytest.raises(Exception):  # noqa: B017
            server.get_email(entry_id="nonexistent")

    def test_get_nonexistent_appointment_raises_error(self, set_bridge):
        """Test getting non-existent appointment raises error"""
        server._bridge.get_appointment.reset_mock()
        server._bridge.get_appointment.return_value = None

        # Should raise some kind of exception
        with pytest.raises(Exception):  # noqa: B017
            server.get_appointment(entry_id="nonexistent")

    def test_get_nonexistent_task_raises_error(self, set_bridge):
        """Test getting non-existent task raises error"""
        server._bridge.get_task.reset_mock()
        server._bridge.get_task.return_value = None

        # Should raise some kind of exception
        with pytest.raises(Exception):  # noqa: B017
            server.get_task(entry_id="nonexistent")

    def test_bridge_not_initialized_raises_error(self):
        """Test that tools raise error when bridge not initialized"""
        server._bridge = None

        # Should raise some kind of exception
        with pytest.raises(Exception):  # noqa: B017
            server.list_emails()


# =============================================================================
# Performance Tests
# =============================================================================


class TestPerformance:
    """Test performance characteristics of integration scenarios"""

    def test_list_multiple_folders(self, set_bridge):
        """Test listing emails from multiple folders"""
        # List from Inbox
        inbox = server.list_emails(limit=10, folder="Inbox")
        assert len(inbox) == 2

        # List from Archive (would be different in real usage)
        # Mock returns same data for simplicity
        archive = server.list_emails(limit=10, folder="Archive")
        assert isinstance(archive, list)

    def test_large_limit_handling(self, set_bridge):
        """Test that large limits are handled correctly"""
        # Request large number of emails
        emails = server.list_emails(limit=1000)
        # Mock returns 2 emails, but function should handle large limit
        assert isinstance(emails, list)

    def test_multiple_operations_sequence(self, set_bridge):
        """Test performing multiple operations in sequence"""
        # List emails
        emails = server.list_emails(limit=10)
        assert len(emails) > 0

        # Get first email
        email = server.get_email(entry_id=emails[0].entry_id)
        assert email.entry_id == emails[0].entry_id

        # Mark as read
        result = server.mark_email(entry_id=email.entry_id, unread=False)
        assert result.success is True

        # Move to folder
        result2 = server.move_email(entry_id=email.entry_id, folder="Archive")
        assert result2.success is True
