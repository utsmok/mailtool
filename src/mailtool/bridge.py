#!/usr/bin/env python3
"""
Outlook COM Automation Bridge
Runs on Windows, callable from WSL2
Provides CLI interface to Outlook email and calendar via COM

Requirements (Windows):
    pip install pywin32

Usage:
    # List recent emails
    python outlook_com_bridge.py emails --limit 10

    # List calendar events
    python outlook_com_bridge.py calendar --days 7

    # Get email body
    python outlook_com_bridge.py email --id <entry_id>
"""

# Modified to test pre-commit hook

import sys
from datetime import datetime, timedelta

import win32com.client


class OutlookBridge:
    """Bridge to Outlook application via COM"""

    @staticmethod
    def _safe_get_attr(obj, attr, default=None):
        """
        Safely get an attribute from a COM object, handling COM errors gracefully

        Args:
            obj: COM object
            attr: Attribute name to get
            default: Default value if attribute access fails

        Returns:
            Attribute value or default
        """
        try:
            return getattr(obj, attr, default)
        # Catch pywintypes.com_error if available, otherwise fall back to Exception
        except Exception:
            return default

    def __init__(self):
        """
        Connect to running Outlook instance or start it

        Launch Logic:
        1. Try GetActiveObject first (if Outlook is running)
        2. Fall back to Dispatch if Outlook is closed
        """
        try:
            self.outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception as e:
            # Outlook might not be running - try to launch it
            try:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception:
                print("Error: Could not connect to or launch Outlook.", file=sys.stderr)
                print(f"Details: {e}", file=sys.stderr)
                print(
                    "Hint: Make sure Outlook is installed and you can launch it manually.",
                    file=sys.stderr,
                )
                sys.exit(1)

        self.namespace = self.outlook.GetNamespace("MAPI")

    def get_inbox(self):
        """Get the inbox folder"""
        inbox = self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        return inbox

    def get_calendar(self):
        """Get the calendar folder"""
        calendar = self.namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
        return calendar

    def get_tasks(self):
        """Get the tasks folder"""
        tasks = self.namespace.GetDefaultFolder(13)  # 13 = olFolderTasks
        return tasks

    def get_folder_by_name(self, folder_name):
        """
        Get a folder by name (e.g., "Sent Items", "Archive", etc.)

        Args:
            folder_name: Name of the folder

        Returns:
            Folder object or None
        """
        try:
            inbox = self.get_inbox()
            # Try to get subfolder of inbox
            folder = inbox.Folders[folder_name]
            return folder
        except Exception:
            try:
                # Try to get from root
                folder = self.namespace.Folders.Item(1).Folders[folder_name]
                return folder
            except Exception:
                return None

    def get_item_by_id(self, entry_id):
        """
        Get any Outlook item by EntryID (O(1) direct access)

        Args:
            entry_id: Outlook EntryID

        Returns:
            Outlook item (MailItem, AppointmentItem, TaskItem, etc.) or None
        """
        try:
            return self.namespace.GetItemFromID(entry_id)
        except Exception:
            return None

    def resolve_smtp_address(self, mail_item):
        """
        Get SMTP address from Exchange address (EX type)

        Args:
            mail_item: Outlook MailItem

        Returns:
            SMTP email address string
        """
        try:
            if (
                (
                    hasattr(mail_item, "SenderEmailType")
                    and mail_item.SenderEmailType == "EX"
                )
                and hasattr(mail_item, "Sender")
                and hasattr(mail_item.Sender, "GetExchangeUser")
            ):
                exchange_user = mail_item.Sender.GetExchangeUser()
                if hasattr(exchange_user, "PrimarySmtpAddress"):
                    return exchange_user.PrimarySmtpAddress
            return (
                mail_item.SenderEmailAddress
                if hasattr(mail_item, "SenderEmailAddress")
                else ""
            )
        except Exception:
            return (
                mail_item.SenderEmailAddress
                if hasattr(mail_item, "SenderEmailAddress")
                else ""
            )

    def list_emails(self, limit=10, folder="Inbox"):
        """
        List emails from the specified folder

        Args:
            limit: Maximum number of emails to return
            folder: Folder name (default: Inbox)

        Returns:
            List of email dictionaries
        """
        inbox = self.get_folder_by_name(folder)
        if not inbox:
            inbox = self.get_inbox()

        items = inbox.Items

        # Sort by received time, most recent first
        items.Sort("[ReceivedTime]", True)

        emails = []
        count = 0
        for item in items:
            if count >= limit:
                break

            try:
                email = {
                    "entry_id": item.EntryID,
                    "subject": item.Subject,
                    "sender": self.resolve_smtp_address(item),
                    "sender_name": item.SenderName,
                    "received_time": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                    if item.ReceivedTime
                    else None,
                    "unread": item.Unread,
                    "has_attachments": item.Attachments.Count > 0,
                }
                emails.append(email)
                count += 1
            except Exception:
                # Skip items that can't be accessed
                continue

        return emails

    def get_email_body(self, entry_id):
        """
        Get the full body of an email by entry ID (O(1) direct access)

        Args:
            entry_id: Outlook EntryID of the email

        Returns:
            Email dictionary with body
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                return {
                    "entry_id": item.EntryID,
                    "subject": item.Subject,
                    "sender": self.resolve_smtp_address(item),
                    "sender_name": item.SenderName,
                    "body": item.Body,
                    "html_body": item.HTMLBody,
                    "received_time": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                    if item.ReceivedTime
                    else None,
                    "has_attachments": item.Attachments.Count > 0,
                }
            except Exception:
                return None
        return None

    def list_calendar_events(self, days=7, all_events=False):
        """
        List calendar events for the next N days

        Args:
            days: Number of days ahead to look
            all_events: If True, return all events without date filtering

        Returns:
            List of event dictionaries
        """
        calendar = self.get_calendar()
        items = calendar.Items

        # CRITICAL: Filter to only appointment items before any other operations
        # This prevents COM errors when encountering meeting requests/responses
        items = items.Restrict(
            "[MessageClass] >= 'IPM.Appointment' AND [MessageClass] < 'IPM.Appointment{'"
        )

        # CRITICAL: Enable recurrence expansion BEFORE sorting
        # Must sort ascending for recurrence to work properly
        items.IncludeRecurrences = True
        items.Sort("[Start]")  # Ascending for recurrence

        # CRITICAL FIX: Apply Restrict BEFORE iterating to avoid "Calendar Bomb"
        # Without this, recurring meetings without end dates generate infinite items
        if not all_events:
            start_date = datetime.now()
            end_date = start_date + timedelta(days=days)
            # Jet SQL format for dates: MM/DD/YYYY HH:MM
            # Use Restrict to filter at COM level before Python iteration
            filter_str = (
                f"[Start] <= '{end_date.strftime('%m/%d/%Y %H:%M')}' "
                f"AND [End] >= '{start_date.strftime('%m/%d/%Y %H:%M')}'"
            )
            items = items.Restrict(filter_str)

        events = []
        for item in items:
            try:
                # Use safe attribute access to handle COM errors
                start = self._safe_get_attr(item, "Start")
                end = self._safe_get_attr(item, "End")

                # Skip if no start time
                if not start:
                    continue

                # Additional Python-level filtering for safety (in case Restrict wasn't applied)
                if not all_events:
                    start_date = datetime.now()
                    end_date = start_date + timedelta(days=days)
                    if not (start >= start_date and start <= end_date):
                        continue

                # Get attendees (safe access)
                required_attendees = self._safe_get_attr(item, "RequiredAttendees", "")
                optional_attendees = self._safe_get_attr(item, "OptionalAttendees", "")

                # Get meeting status
                # ResponseStatus: 0=None, 1=Organizer, 2=Tentative, 3=Accepted, 4=Declined, 5=NotResponded
                response_status = self._safe_get_attr(item, "ResponseStatus")
                response_status_map = {
                    0: "None",
                    1: "Organizer",
                    2: "Tentative",
                    3: "Accepted",
                    4: "Declined",
                    5: "NotResponded",
                }

                # MeetingStatus: 0=Non-meeting, 1=Meeting, 2=Received, 3=Canceled
                meeting_status = self._safe_get_attr(item, "MeetingStatus")
                meeting_status_map = {
                    0: "NonMeeting",
                    1: "Meeting",
                    2: "Received",
                    3: "Canceled",
                }

                event = {
                    "entry_id": self._safe_get_attr(item, "EntryID", ""),
                    "subject": self._safe_get_attr(item, "Subject", "(No Subject)"),
                    "start": start.strftime("%Y-%m-%d %H:%M:%S") if start else None,
                    "end": end.strftime("%Y-%m-%d %H:%M:%S") if end else None,
                    "location": self._safe_get_attr(item, "Location", ""),
                    "organizer": self._safe_get_attr(item, "Organizer"),
                    "all_day": self._safe_get_attr(item, "AllDayEvent", False),
                    "required_attendees": required_attendees,
                    "optional_attendees": optional_attendees,
                    "response_status": response_status_map.get(
                        response_status, "Unknown"
                    ),
                    "meeting_status": meeting_status_map.get(meeting_status, "Unknown"),
                    "response_requested": self._safe_get_attr(
                        item, "ResponseRequested", False
                    ),
                }
                events.append(event)
            except (Exception, BaseException):
                # Skip items that cause errors (including COM fatal errors)
                continue

        return events

    def send_email(
        self,
        to,
        subject,
        body,
        cc=None,
        bcc=None,
        html_body=None,
        file_paths=None,
        save_draft=False,
    ):
        """
        Send an email (or save as draft)

        Args:
            to: Recipient email address
            subject: Email subject
            body: Email body (plain text)
            cc: CC recipients (optional)
            bcc: BCC recipients (optional)
            html_body: HTML body (optional)
            file_paths: List of file paths to attach (optional)
            save_draft: If True, save to Drafts instead of sending

        Returns:
            Draft entry ID if saved, True if sent, False if failed
        """
        try:
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = to
            mail.Subject = subject
            if html_body:
                mail.HTMLBody = html_body
            else:
                mail.Body = body
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc

            # Add attachments
            # NOTE: WSL path translation handled in outlook.sh wrapper
            # Windows Python expects Windows paths (C:\path\to\file)
            if file_paths:
                for file_path in file_paths:
                    mail.Attachments.Add(file_path)

            if save_draft:
                mail.Save()
                return mail.EntryID
            else:
                mail.Send()
                return True
        except Exception as e:
            print(f"Error sending email: {e}", file=sys.stderr)
            return False

    def reply_email(self, entry_id, body, reply_all=False):
        """
        Reply to an email (O(1) direct access)

        Args:
            entry_id: Email entry ID
            body: Reply body
            reply_all: True to reply all, False to reply sender only

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                if reply_all:
                    reply = item.ReplyAll()
                else:
                    reply = item.Reply()
                reply.Body = body
                reply.Send()
                return True
            except Exception as e:
                print(f"Error replying to email: {e}", file=sys.stderr)
                return False
        return False

    def forward_email(self, entry_id, to, body=""):
        """
        Forward an email (O(1) direct access)

        Args:
            entry_id: Email entry ID
            to: Recipient to forward to
            body: Optional additional body text

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                forward = item.Forward()
                forward.To = to
                if body:
                    forward.Body = body + "\n\n" + forward.Body
                forward.Send()
                return True
            except Exception as e:
                print(f"Error forwarding email: {e}", file=sys.stderr)
                return False
        return False

    def mark_email_read(self, entry_id, unread=False):
        """
        Mark an email as read or unread (O(1) direct access)

        Args:
            entry_id: Email entry ID
            unread: True to mark as unread, False to mark as read

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Unread = unread
                item.Save()
                return True
            except Exception:
                return False
        return False

    def move_email(self, entry_id, folder_name):
        """
        Move an email to a different folder (O(1) direct access)

        Args:
            entry_id: Email entry ID
            folder_name: Target folder name

        Returns:
            True if successful
        """
        try:
            item = self.get_item_by_id(entry_id)
            if not item:
                return False

            target_folder = self.get_folder_by_name(folder_name)
            if not target_folder:
                print(f"Error: Folder '{folder_name}' not found", file=sys.stderr)
                return False

            item.Move(target_folder)
            return True
        except Exception as e:
            print(f"Error moving email: {e}", file=sys.stderr)
            return False

    def delete_email(self, entry_id):
        """
        Delete an email (O(1) direct access)

        Args:
            entry_id: Email entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Delete()
                return True
            except Exception:
                return False
        return False

    def download_attachments(self, entry_id, download_dir):
        """
        Download all attachments from an email

        Args:
            entry_id: Email entry ID
            download_dir: Directory to save attachments

        Returns:
            List of downloaded file paths
        """
        item = self.get_item_by_id(entry_id)
        if not item or item.Attachments.Count == 0:
            return []

        downloaded = []
        try:
            import os

            os.makedirs(download_dir, exist_ok=True)

            for i in range(item.Attachments.Count):
                attachment = item.Attachments.Item(i + 1)  # COM is 1-indexed
                filename = attachment.FileName
                filepath = os.path.join(download_dir, filename)
                attachment.SaveAsFile(filepath)
                downloaded.append(filepath)
            return downloaded
        except Exception as e:
            print(f"Error downloading attachments: {e}", file=sys.stderr)
            return []

    def create_appointment(
        self,
        subject,
        start,
        end,
        location="",
        body="",
        all_day=False,
        required_attendees=None,
        optional_attendees=None,
    ):
        """
        Create a calendar appointment

        Args:
            subject: Appointment subject
            start: Start time (YYYY-MM-DD HH:MM:SS)
            end: End time (YYYY-MM-DD HH:MM:SS)
            location: Location
            body: Appointment body/description
            all_day: True for all-day event
            required_attendees: Semicolon-separated list of required attendees
            optional_attendees: Semicolon-separated list of optional attendees

        Returns:
            Appointment entry ID if successful
        """
        try:
            appointment = self.outlook.CreateItem(1)  # 1 = olAppointmentItem
            appointment.Subject = subject
            appointment.Start = datetime.strptime(start, "%Y-%m-%d %H:%M:%S")
            appointment.End = datetime.strptime(end, "%Y-%m-%d %H:%M:%S")
            appointment.Location = location
            appointment.Body = body
            appointment.AllDayEvent = all_day
            if required_attendees:
                appointment.RequiredAttendees = required_attendees
            if optional_attendees:
                appointment.OptionalAttendees = optional_attendees
            appointment.Save()
            return appointment.EntryID
        except Exception as e:
            print(f"Error creating appointment: {e}", file=sys.stderr)
            return None

    def edit_appointment(
        self,
        entry_id,
        required_attendees=None,
        optional_attendees=None,
        subject=None,
        start=None,
        end=None,
        location=None,
        body=None,
    ):
        """
        Edit an existing appointment

        Args:
            entry_id: Appointment entry ID
            required_attendees: Comma-separated list of required attendees
            optional_attendees: Comma-separated list of optional attendees
            subject: New subject (optional)
            start: New start time (optional)
            end: New end time (optional)
            location: New location (optional)
            body: New body (optional)

        Returns:
            True if successful
        """
        try:
            calendar = self.get_calendar()
            for item in calendar.Items:
                if item.EntryID == entry_id:
                    if required_attendees:
                        item.RequiredAttendees = required_attendees
                    if optional_attendees:
                        item.OptionalAttendees = optional_attendees
                    if subject:
                        item.Subject = subject
                    if start:
                        item.Start = datetime.strptime(start, "%Y-%m-%d %H:%M:%S")
                    if end:
                        item.End = datetime.strptime(end, "%Y-%m-%d %H:%M:%S")
                    if location is not None:
                        item.Location = location
                    if body is not None:
                        item.Body = body
                    item.Save()
                    return True
            return False
        except Exception as e:
            print(f"Error editing appointment: {e}", file=sys.stderr)
            return False

    def get_appointment(self, entry_id):
        """
        Get full appointment details by entry ID (O(1) direct access)

        Args:
            entry_id: Appointment entry ID

        Returns:
            Appointment dictionary with full details
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                required_attendees = (
                    item.RequiredAttendees if hasattr(item, "RequiredAttendees") else ""
                )
                optional_attendees = (
                    item.OptionalAttendees if hasattr(item, "OptionalAttendees") else ""
                )
                response_status = (
                    item.ResponseStatus if hasattr(item, "ResponseStatus") else None
                )
                response_status_map = {
                    0: "None",
                    1: "Organizer",
                    2: "Tentative",
                    3: "Accepted",
                    4: "Declined",
                    5: "NotResponded",
                }
                meeting_status = (
                    item.MeetingStatus if hasattr(item, "MeetingStatus") else None
                )
                meeting_status_map = {
                    0: "NonMeeting",
                    1: "Meeting",
                    2: "Received",
                    3: "Canceled",
                }

                return {
                    "entry_id": item.EntryID,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "start": item.Start.strftime("%Y-%m-%d %H:%M:%S")
                    if hasattr(item, "Start") and item.Start
                    else None,
                    "end": item.End.strftime("%Y-%m-%d %H:%M:%S")
                    if hasattr(item, "End") and item.End
                    else None,
                    "location": item.Location if hasattr(item, "Location") else "",
                    "organizer": item.Organizer if hasattr(item, "Organizer") else None,
                    "body": item.Body if hasattr(item, "Body") else "",
                    "all_day": item.AllDayEvent
                    if hasattr(item, "AllDayEvent")
                    else False,
                    "required_attendees": required_attendees,
                    "optional_attendees": optional_attendees,
                    "response_status": response_status_map.get(
                        response_status, "Unknown"
                    ),
                    "meeting_status": meeting_status_map.get(meeting_status, "Unknown"),
                    "response_requested": item.ResponseRequested
                    if hasattr(item, "ResponseRequested")
                    else False,
                }
            except Exception:
                pass
        return None

    def respond_to_meeting(self, entry_id, response):
        """
        Respond to a meeting invitation (O(1) direct access)

        Args:
            entry_id: Appointment entry ID
            response: Response - "accept", "decline", "tentative"

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                response_map = {
                    "accept": 3,  # olResponseAccepted
                    "decline": 4,  # olResponseDeclined
                    "tentative": 2,  # olResponseTentative
                }
                if response.lower() in response_map:
                    item.Response(response_map[response.lower()])
                    item.Send()
                    return True
            except Exception:
                pass
        return False

    def delete_appointment(self, entry_id):
        """
        Delete an appointment (O(1) direct access)

        Args:
            entry_id: Appointment entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Delete()
                return True
            except Exception:
                pass
        return False

    def list_tasks(self, include_completed=False):
        """
        List tasks (only incomplete by default)

        Args:
            include_completed: If True, return all tasks. If False (default), return only incomplete tasks.

        Returns:
            List of task dictionaries
        """
        tasks_folder = self.get_tasks()
        items = tasks_folder.Items

        tasks = []
        for item in items:
            try:
                # Skip completed tasks unless include_completed is True
                if not include_completed and item.Complete:
                    continue

                task = {
                    "entry_id": item.EntryID,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "body": item.Body if hasattr(item, "Body") else "",
                    "due_date": item.DueDate.strftime("%Y-%m-%d")
                    if hasattr(item, "DueDate") and item.DueDate
                    else None,
                    "status": item.Status if hasattr(item, "Status") else None,
                    "priority": item.Importance
                    if hasattr(item, "Importance")
                    else None,
                    "complete": item.Complete if hasattr(item, "Complete") else False,
                    "percent_complete": item.PercentComplete
                    if hasattr(item, "PercentComplete")
                    else 0,
                }
                tasks.append(task)
            except Exception:
                continue

        return tasks

    def list_all_tasks(self):
        """
        List all tasks including completed ones

        Returns:
            List of task dictionaries
        """
        return self.list_tasks(include_completed=True)

    def create_task(self, subject, body="", due_date=None, importance=1):
        """
        Create a new task

        Args:
            subject: Task subject
            body: Task description
            due_date: Due date (YYYY-MM-DD) or None
            importance: 0=Low, 1=Normal, 2=High

        Returns:
            Task entry ID if successful
        """
        try:
            task = self.outlook.CreateItem(3)  # 3 = olTaskItem
            task.Subject = subject
            task.Body = body
            if due_date:
                # Use noon to avoid timezone boundary issues
                task.DueDate = datetime.strptime(
                    f"{due_date} 12:00:00", "%Y-%m-%d %H:%M:%S"
                )
            task.Importance = importance
            task.Save()
            return task.EntryID
        except Exception as e:
            print(f"Error creating task: {e}", file=sys.stderr)
            return None

    def get_task(self, entry_id):
        """
        Get full task details by entry ID (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            Task dictionary with full details
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                return {
                    "entry_id": item.EntryID,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "body": item.Body if hasattr(item, "Body") else "",
                    "due_date": item.DueDate.strftime("%Y-%m-%d")
                    if hasattr(item, "DueDate") and item.DueDate
                    else None,
                    "status": item.Status if hasattr(item, "Status") else None,
                    "priority": item.Importance
                    if hasattr(item, "Importance")
                    else None,
                    "complete": item.Complete if hasattr(item, "Complete") else False,
                    "percent_complete": item.PercentComplete
                    if hasattr(item, "PercentComplete")
                    else 0,
                }
            except Exception:
                pass
        return None

    def edit_task(
        self,
        entry_id,
        subject=None,
        body=None,
        due_date=None,
        importance=None,
        percent_complete=None,
        complete=None,
    ):
        """
        Edit an existing task (O(1) direct access)

        Args:
            entry_id: Task entry ID
            subject: New subject (optional)
            body: New body (optional)
            due_date: New due date YYYY-MM-DD (optional)
            importance: New importance 0=Low, 1=Normal, 2=High (optional)
            percent_complete: New percent complete 0-100 (optional)
            complete: Mark complete/incomplete True/False (optional)

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                if subject:
                    item.Subject = subject
                if body is not None:
                    item.Body = body
                if due_date:
                    # Use noon to avoid timezone boundary issues
                    item.DueDate = datetime.strptime(
                        f"{due_date} 12:00:00", "%Y-%m-%d %H:%M:%S"
                    )
                if importance is not None:
                    item.Importance = importance
                if percent_complete is not None:
                    item.PercentComplete = percent_complete
                    # Update status based on percent_complete
                    if percent_complete == 100:
                        item.Status = 2  # Complete
                        item.Complete = True
                    elif percent_complete == 0:
                        item.Status = 0  # Not started
                    else:
                        item.Status = 1  # In progress
                if complete is not None:
                    item.Complete = complete
                    if complete:
                        item.PercentComplete = 100
                        item.Status = 2
                    else:
                        item.PercentComplete = 0
                        item.Status = 0
                item.Save()
                return True
            except Exception as e:
                print(f"Error editing task: {e}", file=sys.stderr)
            pass
        return False

    def complete_task(self, entry_id):
        """
        Mark a task as complete (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Complete = True
                item.PercentComplete = 100
                item.Status = 2  # olTaskComplete
                item.Save()
                return True
            except Exception:
                pass
        return False

    def delete_task(self, entry_id):
        """
        Delete a task (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Delete()
                return True
            except Exception:
                pass
        return False

    def search_emails(self, filter_query, limit=100):
        """
        Search emails using Outlook Restriction filter (O(1) search, no iteration)

        Args:
            filter_query: SQL query string for filtering
            limit: Max results to return

        Returns:
            List of email dictionaries
        """
        try:
            # First get the folder
            folder = self.get_folder_by_name("Inbox")
            if not folder:
                folder = self.get_inbox()

            items = folder.Items
            # Apply restriction filter
            items = items.Restrict(filter_query)

            # Sort by received time, most recent first
            items.Sort("[ReceivedTime]", True)

            emails = []
            count = 0
            for item in items:
                if count >= limit:
                    break

                try:
                    email = {
                        "entry_id": item.EntryID,
                        "subject": item.Subject,
                        "sender": self.resolve_smtp_address(item),
                        "sender_name": item.SenderName,
                        "received_time": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                        if item.ReceivedTime
                        else None,
                        "unread": item.Unread,
                        "has_attachments": item.Attachments.Count > 0,
                    }
                    emails.append(email)
                    count += 1
                except Exception:
                    # Skip items that can't be accessed
                    continue

            return emails
        except Exception as e:
            print(f"Error searching emails: {e}", file=sys.stderr)
            return []

    def get_free_busy(
        self, email_address=None, start_date=None, end_date=None, entry_id=None
    ):
        """
        Get free/busy status for an email address

        Args:
            email_address: Email address to check (optional, defaults to current user)
            start_date: Start date (YYYY-MM-DD) or datetime object (optional, defaults to today)
            end_date: End date (YYYY-MM-DD) or datetime object (optional, defaults to start + 1 day)
            entry_id: DEPRECATED - Appointment entry ID (legacy, use email_address instead)

        Returns:
            Dictionary with free/busy information
        """
        try:
            # Handle legacy entry_id parameter (extract first required attendee)
            if entry_id and not email_address:
                item = self.get_item_by_id(entry_id)
                if (
                    item
                    and hasattr(item, "RequiredAttendees")
                    and item.RequiredAttendees
                ):
                    attendees = item.RequiredAttendees.split(";")
                    if attendees:
                        email_address = attendees[0].strip()

            # Default to current user if no email provided
            if not email_address:
                email_address = self.namespace.CurrentUser.Address

            # Default to today if no dates provided
            if not start_date:
                start_date = datetime.now()
            elif isinstance(start_date, str):
                start_date = datetime.strptime(start_date, "%Y-%m-%d")

            if not end_date:
                end_date = start_date + timedelta(days=1)
            elif isinstance(end_date, str):
                end_date = datetime.strptime(end_date, "%Y-%m-%d")

            # Create recipient and get free/busy
            recipient = self.namespace.CreateRecipient(email_address)
            if recipient.Resolve():
                # FreeBusy returns a string with time slots and status
                # 0=Free, 1=Tentative, 2=Busy, 3=Out of Office, 4=Working Elsewhere
                freebusy = recipient.FreeBusy(
                    start_date, 60 * 24
                )  # 1440 minutes = 1 day
                return {
                    "email": email_address,
                    "start_date": start_date.strftime("%Y-%m-%d"),
                    "end_date": end_date.strftime("%Y-%m-%d"),
                    "free_busy": freebusy,
                    "resolved": True,
                }
            else:
                return {
                    "email": email_address,
                    "start_date": start_date.strftime("%Y-%m-%d"),
                    "end_date": end_date.strftime("%Y-%m-%d"),
                    "error": "Could not resolve email address",
                    "resolved": False,
                }
        except Exception as e:
            return {
                "email": email_address if email_address else "unknown",
                "error": str(e),
                "resolved": False,
            }
