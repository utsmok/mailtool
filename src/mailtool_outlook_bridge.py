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

import win32com.client
import json
import sys
import argparse
from datetime import datetime, timedelta
from pathlib import Path


class OutlookBridge:
    """Bridge to Outlook application via COM"""

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
            except Exception as e2:
                print(f"Error: Could not connect to or launch Outlook.", file=sys.stderr)
                print(f"Details: {e}", file=sys.stderr)
                print(f"Hint: Make sure Outlook is installed and you can launch it manually.", file=sys.stderr)
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
        except:
            try:
                # Try to get from root
                folder = self.namespace.Folders.Item(1).Folders[folder_name]
                return folder
            except:
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
        except Exception as e:
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
            if hasattr(mail_item, 'SenderEmailType') and mail_item.SenderEmailType == "EX":
                if hasattr(mail_item, 'Sender') and hasattr(mail_item.Sender, 'GetExchangeUser'):
                    exchange_user = mail_item.Sender.GetExchangeUser()
                    if hasattr(exchange_user, 'PrimarySmtpAddress'):
                        return exchange_user.PrimarySmtpAddress
            return mail_item.SenderEmailAddress if hasattr(mail_item, 'SenderEmailAddress') else ""
        except Exception:
            return mail_item.SenderEmailAddress if hasattr(mail_item, 'SenderEmailAddress') else ""

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
                    "received_time": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S") if item.ReceivedTime else None,
                    "unread": item.Unread,
                    "has_attachments": item.Attachments.Count > 0
                }
                emails.append(email)
                count += 1
            except Exception as e:
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
                    "received_time": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S") if item.ReceivedTime else None,
                    "has_attachments": item.Attachments.Count > 0
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

        # CRITICAL: Enable recurrence expansion BEFORE sorting
        # Must sort ascending for recurrence to work properly
        items.IncludeRecurrences = True
        items.Sort("[Start]")  # Ascending for recurrence

        events = []
        for item in items:
            try:
                start = item.Start if hasattr(item, 'Start') else None
                end = item.End if hasattr(item, 'End') else None

                # Skip if no start time
                if not start:
                    continue

                # Filter by date range unless --all is specified
                if not all_events:
                    start_date = datetime.now()
                    end_date = start_date + timedelta(days=days)
                    if not (start >= start_date and start <= end_date):
                        continue

                # Get attendees
                required_attendees = item.RequiredAttendees if hasattr(item, 'RequiredAttendees') else ""
                optional_attendees = item.OptionalAttendees if hasattr(item, 'OptionalAttendees') else ""

                # Get meeting status
                # ResponseStatus: 0=None, 1=Organizer, 2=Tentative, 3=Accepted, 4=Declined, 5=NotResponded
                response_status = item.ResponseStatus if hasattr(item, 'ResponseStatus') else None
                response_status_map = {
                    0: "None",
                    1: "Organizer",
                    2: "Tentative",
                    3: "Accepted",
                    4: "Declined",
                    5: "NotResponded"
                }

                # MeetingStatus: 0=Non-meeting, 1=Meeting, 2=Received, 3=Canceled
                meeting_status = item.MeetingStatus if hasattr(item, 'MeetingStatus') else None
                meeting_status_map = {
                    0: "NonMeeting",
                    1: "Meeting",
                    2: "Received",
                    3: "Canceled"
                }

                event = {
                    "entry_id": item.EntryID,
                    "subject": item.Subject if hasattr(item, 'Subject') else "(No Subject)",
                    "start": start.strftime("%Y-%m-%d %H:%M:%S") if start else None,
                    "end": end.strftime("%Y-%m-%d %H:%M:%S") if end else None,
                    "location": item.Location if hasattr(item, 'Location') else "",
                    "organizer": item.Organizer if hasattr(item, 'Organizer') else None,
                    "all_day": item.AllDayEvent if hasattr(item, 'AllDayEvent') else False,
                    "required_attendees": required_attendees,
                    "optional_attendees": optional_attendees,
                    "response_status": response_status_map.get(response_status, "Unknown"),
                    "meeting_status": meeting_status_map.get(meeting_status, "Unknown"),
                    "response_requested": item.ResponseRequested if hasattr(item, 'ResponseRequested') else False
                }
                events.append(event)
            except Exception as e:
                # Skip items that cause errors
                continue

        return events

    def send_email(self, to, subject, body, cc=None, bcc=None, html_body=None,
                 file_paths=None, save_draft=False):
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
            True if successful
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

    def create_appointment(self, subject, start, end, location="", body="", all_day=False,
                         required_attendees=None, optional_attendees=None):
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

    def edit_appointment(self, entry_id, required_attendees=None, optional_attendees=None,
                        subject=None, start=None, end=None, location=None, body=None):
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
                required_attendees = item.RequiredAttendees if hasattr(item, 'RequiredAttendees') else ""
                optional_attendees = item.OptionalAttendees if hasattr(item, 'OptionalAttendees') else ""
                response_status = item.ResponseStatus if hasattr(item, 'ResponseStatus') else None
                response_status_map = {0: "None", 1: "Organizer", 2: "Tentative", 3: "Accepted", 4: "Declined", 5: "NotResponded"}
                meeting_status = item.MeetingStatus if hasattr(item, 'MeetingStatus') else None
                meeting_status_map = {0: "NonMeeting", 1: "Meeting", 2: "Received", 3: "Canceled"}

                return {
                    "entry_id": item.EntryID,
                    "subject": item.Subject if hasattr(item, 'Subject') else "(No Subject)",
                    "start": item.Start.strftime("%Y-%m-%d %H:%M:%S") if hasattr(item, 'Start') and item.Start else None,
                    "end": item.End.strftime("%Y-%m-%d %H:%M:%S") if hasattr(item, 'End') and item.End else None,
                    "location": item.Location if hasattr(item, 'Location') else "",
                    "organizer": item.Organizer if hasattr(item, 'Organizer') else None,
                    "body": item.Body if hasattr(item, 'Body') else "",
                    "all_day": item.AllDayEvent if hasattr(item, 'AllDayEvent') else False,
                    "required_attendees": required_attendees,
                    "optional_attendees": optional_attendees,
                    "response_status": response_status_map.get(response_status, "Unknown"),
                    "meeting_status": meeting_status_map.get(meeting_status, "Unknown"),
                    "response_requested": item.ResponseRequested if hasattr(item, 'ResponseRequested') else False
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
                    "tentative": 2  # olResponseTentative
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

    def list_tasks(self):
        """
        List all tasks

        Returns:
            List of task dictionaries
        """
        tasks_folder = self.get_tasks()
        items = tasks_folder.Items

        tasks = []
        for item in items:
            try:
                task = {
                    "entry_id": item.EntryID,
                    "subject": item.Subject if hasattr(item, 'Subject') else "(No Subject)",
                    "body": item.Body if hasattr(item, 'Body') else "",
                    "due_date": item.DueDate.strftime("%Y-%m-%d") if hasattr(item, 'DueDate') and item.DueDate else None,
                    "status": item.Status if hasattr(item, 'Status') else None,
                    "priority": item.Importance if hasattr(item, 'Importance') else None,
                    "complete": item.Complete if hasattr(item, 'Complete') else False,
                    "percent_complete": item.PercentComplete if hasattr(item, 'PercentComplete') else 0
                }
                tasks.append(task)
            except Exception as e:
                continue

        return tasks

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
                task.DueDate = datetime.strptime(due_date, "%Y-%m-%d")
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
                    "subject": item.Subject if hasattr(item, 'Subject') else "(No Subject)",
                    "body": item.Body if hasattr(item, 'Body') else "",
                    "due_date": item.DueDate.strftime("%Y-%m-%d") if hasattr(item, 'DueDate') and item.DueDate else None,
                    "status": item.Status if hasattr(item, 'Status') else None,
                    "priority": item.Importance if hasattr(item, 'Importance') else None,
                    "complete": item.Complete if hasattr(item, 'Complete') else False,
                    "percent_complete": item.PercentComplete if hasattr(item, 'PercentComplete') else 0
                }
            except Exception:
                pass
        return None

    def edit_task(self, entry_id, subject=None, body=None, due_date=None, importance=None,
                 percent_complete=None, complete=None):
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
                    item.DueDate = datetime.strptime(due_date, "%Y-%m-%d")
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
                        "received_time": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S") if item.ReceivedTime else None,
                        "unread": item.Unread,
                        "has_attachments": item.Attachments.Count > 0
                    }
                    emails.append(email)
                    count += 1
                except Exception as e:
                    # Skip items that can't be accessed
                    continue

            return emails
        except Exception as e:
            print(f"Error searching emails: {e}", file=sys.stderr)
            return []

    def get_free_busy(self, entry_id, start_date, end_date):
        """
        Get free/busy status for an appointment

        Args:
            entry_id: Appointment entry ID to check
            start_date: Start date (YYYY-MM-DD)
            end_date: End date (YYYY-MM-DD)

        Returns:
            Dictionary with free/busy information
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                # Get recipient
                if hasattr(item, 'RequiredAttendees') and item.RequiredAttendees:
                    # Parse first attendee
                    attendees = item.RequiredAttendees.split(';')
                    if attendees:
                        first_attendee = attendees[0].strip()
                        # Try to get their free/busy
                        try:
                            recipient = self.namespace.CreateRecipient("olExchange", first_attendee)
                            freebusy = recipient.FreeBusy(start_date, end_date)
                            return {
                                "entry_id": entry_id,
                                "attendee": first_attendee,
                                "start_date": start_date,
                                "end_date": end_date,
                                "free_busy": freebusy
                            }
                        except Exception as e:
                            print(f"Error getting free/busy: {e}", file=sys.stderr)
            except Exception as e:
                print(f"Error with get_free_busy: {e}", file=sys.stderr)
        return None


def main():
    parser = argparse.ArgumentParser(description="Outlook COM Bridge")
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # Emails command
    email_parser = subparsers.add_parser("emails", help="List emails")
    email_parser.add_argument("--limit", type=int, default=10, help="Max emails to return")
    email_parser.add_argument("--folder", default="Inbox", help="Folder name (default: Inbox)")

    # Calendar command
    cal_parser = subparsers.add_parser("calendar", help="List calendar events")
    cal_parser.add_argument("--days", type=int, default=7, help="Days ahead to look")
    cal_parser.add_argument("--all", action="store_true", help="Show all events without date filtering")

    # Get email command
    get_parser = subparsers.add_parser("email", help="Get email details")
    get_parser.add_argument("--id", required=True, help="Email entry ID")

    # Send email command
    send_parser = subparsers.add_parser("send", help="Send an email")
    send_parser.add_argument("--to", required=True, help="Recipient email address")
    send_parser.add_argument("--subject", required=True, help="Email subject")
    send_parser.add_argument("--body", required=True, help="Email body")
    send_parser.add_argument("--cc", help="CC recipients")
    send_parser.add_argument("--bcc", help="BCC recipients")
    send_parser.add_argument("--html", help="HTML body (rich text)")
    send_parser.add_argument("--attach", nargs="+", help="File paths to attach")
    send_parser.add_argument("--draft", action="store_true", help="Save as draft instead of sending")

    # Download attachments command
    attach_parser = subparsers.add_parser("attachments", help="Download email attachments")
    attach_parser.add_argument("--id", required=True, help="Email entry ID")
    attach_parser.add_argument("--dir", required=True, help="Directory to save attachments")

    # Reply email command
    reply_parser = subparsers.add_parser("reply", help="Reply to an email")
    reply_parser.add_argument("--id", required=True, help="Email entry ID")
    reply_parser.add_argument("--body", required=True, help="Reply body")
    reply_parser.add_argument("--all", action="store_true", help="Reply all instead of just sender")

    # Forward email command
    forward_parser = subparsers.add_parser("forward", help="Forward an email")
    forward_parser.add_argument("--id", required=True, help="Email entry ID")
    forward_parser.add_argument("--to", required=True, help="Recipient to forward to")
    forward_parser.add_argument("--body", default="", help="Additional body text")

    # Search emails command
    search_parser = subparsers.add_parser("search", help="Search emails using Restriction")
    search_parser.add_argument("--query", required=True, help="SQL filter query (e.g., urn:schemas:httpmail:subject LIKE '%keyword%')")
    search_parser.add_argument("--limit", type=int, default=100, help="Max results to return")

    # Mark email command
    mark_parser = subparsers.add_parser("mark", help="Mark email as read/unread")
    mark_parser.add_argument("--id", required=True, help="Email entry ID")
    mark_parser.add_argument("--unread", action="store_true", help="Mark as unread (default: read)")

    # Move email command
    move_parser = subparsers.add_parser("move", help="Move email to folder")
    move_parser.add_argument("--id", required=True, help="Email entry ID")
    move_parser.add_argument("--folder", required=True, help="Target folder name")

    # Delete email command
    del_email_parser = subparsers.add_parser("delete-email", help="Delete an email")
    del_email_parser.add_argument("--id", required=True, help="Email entry ID")

    # Create appointment command
    create_appt_parser = subparsers.add_parser("create-appt", help="Create calendar appointment")
    create_appt_parser.add_argument("--subject", required=True, help="Appointment subject")
    create_appt_parser.add_argument("--start", required=True, help="Start time (YYYY-MM-DD HH:MM:SS)")
    create_appt_parser.add_argument("--end", required=True, help="End time (YYYY-MM-DD HH:MM:SS)")
    create_appt_parser.add_argument("--location", default="", help="Location")
    create_appt_parser.add_argument("--body", default="", help="Appointment description")
    create_appt_parser.add_argument("--all-day", action="store_true", help="All-day event")
    create_appt_parser.add_argument("--required", help="Required attendees (semicolon-separated)")
    create_appt_parser.add_argument("--optional", help="Optional attendees (semicolon-separated)")

    # Get appointment command
    get_appt_parser = subparsers.add_parser("appointment", help="Get appointment details")
    get_appt_parser.add_argument("--id", required=True, help="Appointment entry ID")

    # Delete appointment command
    del_appt_parser = subparsers.add_parser("delete-appt", help="Delete an appointment")
    del_appt_parser.add_argument("--id", required=True, help="Appointment entry ID")

    # Respond to meeting command
    respond_parser = subparsers.add_parser("respond", help="Respond to meeting invitation")
    respond_parser.add_argument("--id", required=True, help="Appointment entry ID")
    respond_parser.add_argument("--response", required=True, choices=["accept", "decline", "tentative"],
                              help="Meeting response")

    # Free/busy command
    freebusy_parser = subparsers.add_parser("freebusy", help="Get free/busy status")
    freebusy_parser.add_argument("--id", required=True, help="Appointment or email entry ID")
    freebusy_parser.add_argument("--start", required=True, help="Start date (YYYY-MM-DD)")
    freebusy_parser.add_argument("--end", required=True, help="End date (YYYY-MM-DD)")

    # Edit appointment command
    edit_appt_parser = subparsers.add_parser("edit-appt", help="Edit an appointment")
    edit_appt_parser.add_argument("--id", required=True, help="Appointment entry ID")
    edit_appt_parser.add_argument("--required", help="Required attendees (comma-separated)")
    edit_appt_parser.add_argument("--optional", help="Optional attendees (comma-separated)")
    edit_appt_parser.add_argument("--subject", help="New subject")
    edit_appt_parser.add_argument("--start", help="New start time (YYYY-MM-DD HH:MM:SS)")
    edit_appt_parser.add_argument("--end", help="New end time (YYYY-MM-DD HH:MM:SS)")
    edit_appt_parser.add_argument("--location", help="New location")
    edit_appt_parser.add_argument("--body", help="New body/description")

    # Tasks command
    tasks_parser = subparsers.add_parser("tasks", help="List all tasks")

    # Get task command
    get_task_parser = subparsers.add_parser("task", help="Get task details")
    get_task_parser.add_argument("--id", required=True, help="Task entry ID")

    # Create task command
    create_task_parser = subparsers.add_parser("create-task", help="Create a new task")
    create_task_parser.add_argument("--subject", required=True, help="Task subject")
    create_task_parser.add_argument("--body", default="", help="Task description")
    create_task_parser.add_argument("--due", help="Due date (YYYY-MM-DD)")
    create_task_parser.add_argument("--priority", type=int, choices=[0, 1, 2], default=1,
                                   help="Priority: 0=Low, 1=Normal, 2=High")

    # Edit task command
    edit_task_parser = subparsers.add_parser("edit-task", help="Edit a task")
    edit_task_parser.add_argument("--id", required=True, help="Task entry ID")
    edit_task_parser.add_argument("--subject", help="New subject")
    edit_task_parser.add_argument("--body", help="New description")
    edit_task_parser.add_argument("--due", help="New due date (YYYY-MM-DD)")
    edit_task_parser.add_argument("--priority", type=int, choices=[0, 1, 2], help="New priority")
    edit_task_parser.add_argument("--percent", type=int, choices=range(0, 101), help="Percent complete (0-100)")
    edit_task_parser.add_argument("--complete", type=bool, help="Mark complete/incomplete (true/false)")

    # Complete task command
    complete_task_parser = subparsers.add_parser("complete-task", help="Mark task as complete")
    complete_task_parser.add_argument("--id", required=True, help="Task entry ID")

    # Delete task command
    del_task_parser = subparsers.add_parser("delete-task", help="Delete a task")
    del_task_parser.add_argument("--id", required=True, help="Task entry ID")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    bridge = OutlookBridge()

    if args.command == "emails":
        emails = bridge.list_emails(limit=args.limit, folder=args.folder)
        print(json.dumps(emails, indent=2))

    elif args.command == "calendar":
        events = bridge.list_calendar_events(days=args.days, all_events=args.all)
        print(json.dumps(events, indent=2))

    elif args.command == "email":
        email = bridge.get_email_body(entry_id=args.id)
        if email:
            print(json.dumps(email, indent=2))
        else:
            print("Email not found", file=sys.stderr)
            sys.exit(1)

    elif args.command == "send":
        result = bridge.send_email(args.to, args.subject, args.body, args.cc, args.bcc,
                                     html_body=args.html, file_paths=args.attach, save_draft=args.draft)
        if result:
            if args.draft:
                print(json.dumps({"status": "success", "entry_id": result, "message": "Draft saved"}))
            else:
                print(json.dumps({"status": "success", "message": "Email sent"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to send email"}))
            sys.exit(1)

    elif args.command == "attachments":
        downloaded = bridge.download_attachments(args.id, args.dir)
        if downloaded:
            print(json.dumps({"status": "success", "attachments": downloaded}))
        else:
            print(json.dumps({"status": "error", "message": "No attachments found or failed to download"}))
            sys.exit(1)

    elif args.command == "reply":
        result = bridge.reply_email(args.id, args.body, reply_all=args.all)
        if result:
            print(json.dumps({"status": "success", "message": "Reply sent"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to send reply"}))
            sys.exit(1)

    elif args.command == "forward":
        result = bridge.forward_email(args.id, args.to, args.body)
        if result:
            print(json.dumps({"status": "success", "message": "Email forwarded"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to forward email"}))
            sys.exit(1)

    elif args.command == "search":
        emails = bridge.search_emails(args.query, limit=args.limit)
        print(json.dumps(emails, indent=2))

    elif args.command == "mark":
        result = bridge.mark_email_read(args.id, unread=args.unread)
        if result:
            status = "unread" if args.unread else "read"
            print(json.dumps({"status": "success", "message": f"Email marked as {status}"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to mark email"}))
            sys.exit(1)

    elif args.command == "move":
        result = bridge.move_email(args.id, args.folder)
        if result:
            print(json.dumps({"status": "success", "message": f"Email moved to {args.folder}"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to move email"}))
            sys.exit(1)

    elif args.command == "delete-email":
        result = bridge.delete_email(args.id)
        if result:
            print(json.dumps({"status": "success", "message": "Email deleted"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to delete email"}))
            sys.exit(1)

    elif args.command == "create-appt":
        entry_id = bridge.create_appointment(args.subject, args.start, args.end, args.location, args.body,
                                              args.all_day, args.required, args.optional)
        if entry_id:
            print(json.dumps({"status": "success", "entry_id": entry_id, "message": "Appointment created"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to create appointment"}))
            sys.exit(1)

    elif args.command == "appointment":
        appointment = bridge.get_appointment(args.id)
        if appointment:
            print(json.dumps(appointment, indent=2))
        else:
            print("Appointment not found", file=sys.stderr)
            sys.exit(1)

    elif args.command == "delete-appt":
        result = bridge.delete_appointment(args.id)
        if result:
            print(json.dumps({"status": "success", "message": "Appointment deleted"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to delete appointment"}))
            sys.exit(1)

    elif args.command == "edit-appt":
        result = bridge.edit_appointment(
            args.id,
            required_attendees=args.required,
            optional_attendees=args.optional,
            subject=args.subject,
            start=args.start,
            end=args.end,
            location=args.location,
            body=args.body
        )
        if result:
            print(json.dumps({"status": "success", "message": "Appointment updated"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to edit appointment"}))
            sys.exit(1)

    elif args.command == "respond":
        result = bridge.respond_to_meeting(args.id, args.response)
        if result:
            print(json.dumps({"status": "success", "message": f"Meeting {args.response}ed"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to respond to meeting"}))
            sys.exit(1)

    elif args.command == "freebusy":
        freebusy = bridge.get_free_busy(args.id, args.start, args.end)
        if freebusy:
            print(json.dumps(freebusy, indent=2))
        else:
            print(json.dumps({"status": "error", "message": "Could not get free/busy information"}))
            sys.exit(1)

    elif args.command == "tasks":
        tasks = bridge.list_tasks()
        print(json.dumps(tasks, indent=2))

    elif args.command == "task":
        task = bridge.get_task(args.id)
        if task:
            print(json.dumps(task, indent=2))
        else:
            print("Task not found", file=sys.stderr)
            sys.exit(1)

    elif args.command == "create-task":
        entry_id = bridge.create_task(args.subject, args.body, args.due, args.priority)
        if entry_id:
            print(json.dumps({"status": "success", "entry_id": entry_id, "message": "Task created"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to create task"}))
            sys.exit(1)

    elif args.command == "edit-task":
        result = bridge.edit_task(args.id, args.subject, args.body, args.due, args.priority, args.percent, args.complete)
        if result:
            print(json.dumps({"status": "success", "message": "Task updated"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to edit task"}))
            sys.exit(1)

    elif args.command == "complete-task":
        result = bridge.complete_task(args.id)
        if result:
            print(json.dumps({"status": "success", "message": "Task marked as complete"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to complete task"}))
            sys.exit(1)

    elif args.command == "delete-task":
        result = bridge.delete_task(args.id)
        if result:
            print(json.dumps({"status": "success", "message": "Task deleted"}))
        else:
            print(json.dumps({"status": "error", "message": "Failed to delete task"}))
            sys.exit(1)


if __name__ == "__main__":
    main()
