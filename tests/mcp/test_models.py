"""
Model Validation Tests

Tests all Pydantic models to ensure they validate data correctly.
This test suite runs without requiring Outlook to be running.
"""

import pytest
from pydantic import ValidationError

from mailtool.mcp.models import EmailDetails, EmailSummary, SendEmailResult


class TestEmailSummary:
    """Test EmailSummary model validation"""

    def test_valid_email_summary(self):
        """Test EmailSummary accepts valid data"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "received_time": "2025-01-19 10:30:00",
            "unread": True,
            "has_attachments": False,
        }
        email = EmailSummary(**data)
        assert email.entry_id == "test-entry-id-123"
        assert email.subject == "Test Subject"
        assert email.sender == "sender@example.com"
        assert email.sender_name == "Test Sender"
        assert email.received_time == "2025-01-19 10:30:00"
        assert email.unread is True
        assert email.has_attachments is False

    def test_email_summary_with_none_received_time(self):
        """Test EmailSummary accepts None for received_time"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "received_time": None,
            "unread": False,
            "has_attachments": True,
        }
        email = EmailSummary(**data)
        assert email.received_time is None

    def test_email_summary_missing_required_fields(self):
        """Test EmailSummary raises ValidationError for missing required fields"""
        data = {
            "entry_id": "test-entry-id-123",
            # Missing: subject, sender, sender_name, unread, has_attachments
        }
        with pytest.raises(ValidationError):
            EmailSummary(**data)


class TestEmailDetails:
    """Test EmailDetails model validation"""

    def test_valid_email_details(self):
        """Test EmailDetails accepts valid data"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "body": "Plain text body",
            "html_body": "<html><body>HTML body</body></html>",
            "received_time": "2025-01-19 10:30:00",
            "has_attachments": False,
        }
        email = EmailDetails(**data)
        assert email.entry_id == "test-entry-id-123"
        assert email.subject == "Test Subject"
        assert email.body == "Plain text body"
        assert email.html_body == "<html><body>HTML body</body></html>"
        assert email.received_time == "2025-01-19 10:30:00"
        assert email.has_attachments is False

    def test_email_details_with_none_received_time(self):
        """Test EmailDetails accepts None for received_time"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "body": "Plain text body",
            "html_body": "<html><body>HTML body</body></html>",
            "received_time": None,
            "has_attachments": False,
        }
        email = EmailDetails(**data)
        assert email.received_time is None

    def test_email_details_missing_required_fields(self):
        """Test EmailDetails raises ValidationError for missing required fields"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            # Missing: sender, sender_name, body, html_body, has_attachments
        }
        with pytest.raises(ValidationError):
            EmailDetails(**data)


class TestSendEmailResult:
    """Test SendEmailResult model validation"""

    def test_sent_email_result(self):
        """Test SendEmailResult for successfully sent email"""
        data = {
            "success": True,
            "entry_id": None,
            "message": "Email sent successfully",
        }
        result = SendEmailResult(**data)
        assert result.success is True
        assert result.entry_id is None
        assert result.message == "Email sent successfully"

    def test_draft_saved_result(self):
        """Test SendEmailResult for saved draft"""
        data = {
            "success": True,
            "entry_id": "draft-entry-id-456",
            "message": "Draft saved successfully",
        }
        result = SendEmailResult(**data)
        assert result.success is True
        assert result.entry_id == "draft-entry-id-456"
        assert result.message == "Draft saved successfully"

    def test_failed_send_result(self):
        """Test SendEmailResult for failed email"""
        data = {
            "success": False,
            "entry_id": None,
            "message": "Failed to send email",
        }
        result = SendEmailResult(**data)
        assert result.success is False
        assert result.entry_id is None
        assert result.message == "Failed to send email"

    def test_send_result_default_entry_id(self):
        """Test SendEmailResult entry_id defaults to None"""
        data = {
            "success": True,
            "message": "Email sent successfully",
        }
        result = SendEmailResult(**data)
        assert result.success is True
        assert result.entry_id is None
        assert result.message == "Email sent successfully"

    def test_send_result_missing_required_fields(self):
        """Test SendEmailResult raises ValidationError for missing required fields"""
        data = {
            "success": True,
            # Missing: message
        }
        with pytest.raises(ValidationError):
            SendEmailResult(**data)


class TestModelSerialization:
    """Test model serialization and deserialization"""

    def test_email_summary_serialization(self):
        """Test EmailSummary can be serialized to dict and JSON"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "received_time": "2025-01-19 10:30:00",
            "unread": True,
            "has_attachments": False,
        }
        email = EmailSummary(**data)
        # Test model_dump
        dumped = email.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = email.model_dump_json()
        assert "test-entry-id-123" in json_str
        assert "Test Subject" in json_str

    def test_email_details_serialization(self):
        """Test EmailDetails can be serialized to dict and JSON"""
        data = {
            "entry_id": "test-entry-id-123",
            "subject": "Test Subject",
            "sender": "sender@example.com",
            "sender_name": "Test Sender",
            "body": "Plain text body",
            "html_body": "<html><body>HTML body</body></html>",
            "received_time": "2025-01-19 10:30:00",
            "has_attachments": False,
        }
        email = EmailDetails(**data)
        # Test model_dump
        dumped = email.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = email.model_dump_json()
        assert "Plain text body" in json_str

    def test_send_email_result_serialization(self):
        """Test SendEmailResult can be serialized to dict and JSON"""
        data = {
            "success": True,
            "entry_id": "draft-entry-id-456",
            "message": "Draft saved successfully",
        }
        result = SendEmailResult(**data)
        # Test model_dump
        dumped = result.model_dump()
        assert dumped == data
        # Test model_dump_json
        json_str = result.model_dump_json()
        assert "draft-entry-id-456" in json_str
