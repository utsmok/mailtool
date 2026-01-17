# Mailtool Command Reference

Complete list of all commands for the Outlook automation bridge.

## Email Commands

### List Emails
```bash
./outlook.sh emails [--limit N]
```
List recent emails from inbox.

**Example:**
```bash
./outlook.sh emails --limit 5
```

### Get Email Details
```bash
./outlook.sh email --id <entry_id>
```
Get full email body and details.

**Example:**
```bash
./outlook.sh email --id "00000000604FC3F48B..."
```

### Send Email
```bash
./outlook.sh send --to <email> --subject <text> --body <text> [--cc <emails>] [--bcc <emails>]
```
Send a new email.

**Example:**
```bash
./outlook.sh send \
  --to "someone@example.com" \
  --subject "Meeting tomorrow" \
  --body "Let's meet at 2pm" \
  --cc "other@example.com"
```

### Mark Email as Read/Unread
```bash
./outlook.sh mark --id <entry_id> [--unread]
```
Mark an email as read (default) or unread.

**Examples:**
```bash
# Mark as read
./outlook.sh mark --id "00000000604FC3F48B..."

# Mark as unread
./outlook.sh mark --id "00000000604FC3F48B..." --unread
```

### Move Email to Folder
```bash
./outlook.sh move --id <entry_id> --folder <folder_name>
```
Move an email to a different folder.

**Example:**
```bash
./outlook.sh move --id "00000000604FC3F48B..." --folder "Archive"
```

**Common folder names:** `Archive`, `Deleted Items`, `Sent Items`, `Junk Email`, or any custom folder.

### Delete Email
```bash
./outlook.sh delete-email --id <entry_id>
```
Delete an email.

**Example:**
```bash
./outlook.sh delete-email --id "00000000604FC3F48B..."
```

---

## Calendar Commands

### List Calendar Events
```bash
./outlook.sh calendar [--days N] [--all]
```
List calendar events.

**Examples:**
```bash
# Next 7 days (default)
./outlook.sh calendar

# Next 30 days
./outlook.sh calendar --days 30

# ALL events (past and future)
./outlook.sh calendar --all
```

### Create Appointment
```bash
./outlook.sh create-appt \
  --subject <text> \
  --start "YYYY-MM-DD HH:MM:SS" \
  --end "YYYY-MM-DD HH:MM:SS" \
  [--location <text>] \
  [--body <text>] \
  [--all-day]
```
Create a new calendar appointment.

**Example:**
```bash
./outlook.sh create-appt \
  --subject "Team Meeting" \
  --start "2026-01-25 14:00:00" \
  --end "2026-01-25 15:00:00" \
  --location "Room 101" \
  --body "Discuss Q1 goals"
```

**All-day event:**
```bash
./outlook.sh create-appt \
  --subject "Company Holiday" \
  --start "2026-12-25 00:00:00" \
  --end "2026-12-25 23:59:59" \
  --all-day
```

### Delete Appointment
```bash
./outlook.sh delete-appt --id <entry_id>
```
Delete an appointment.

**Example:**
```bash
./outlook.sh delete-appt --id "00000000604FC3F48B..."
```

---

## Task Commands

### List Tasks
```bash
./outlook.sh tasks
```
List all tasks.

**Example:**
```bash
./outlook.sh tasks
```

**Response fields:**
- `subject`: Task title
- `body`: Task description
- `due_date`: Due date (YYYY-MM-DD)
- `status`: 0=Not started, 1=In progress, 2=Complete
- `priority`: 0=Low, 1=Normal, 2=High
- `complete`: Boolean
- `percent_complete`: 0-100

### Create Task
```bash
./outlook.sh create-task \
  --subject <text> \
  [--body <text>] \
  [--due YYYY-MM-DD] \
  [--priority N]
```
Create a new task.

**Example:**
```bash
./outlook.sh create-task \
  --subject "Review project proposal" \
  --body "Check the Q1 budget and provide feedback" \
  --due "2026-01-30" \
  --priority 2
```

**Priority levels:** 0=Low, 1=Normal (default), 2=High

### Complete Task
```bash
./outlook.sh complete-task --id <entry_id>
```
Mark a task as complete.

**Example:**
```bash
./outlook.sh complete-task --id "00000000604FC3F48B..."
```

### Delete Task
```bash
./outlook.sh delete-task --id <entry_id>
```
Delete a task.

**Example:**
```bash
./outlook.sh delete-task --id "00000000604FC3F48B..."
```

---

## Common Patterns

### Get Entry ID
Most commands require an `entry_id`. Get it from list commands:

```bash
# Get email ID
./outlook.sh emails --limit 1

# Get appointment ID
./outlook.sh calendar --all

# Get task ID
./outlook.sh tasks
```

### Pipe to jq for JSON Processing
```bash
# Get unread emails
./outlook.sh emails | jq '.[] | select(.unread == true)'

# Get specific fields
./outlook.sh emails | jq '.[] | {subject, sender, date: .received_time}'

# Get incomplete tasks
./outlook.sh tasks | jq '.[] | select(.complete == false)'
```

### Chain Commands (Bash)
```bash
# Mark all unread emails as read
for id in $(./outlook.sh emails | jq -r '.[] | select(.unread == true) | .entry_id'); do
  ./outlook.sh mark --id "$id"
done

# Complete all overdue tasks
for id in $(./outlook.sh tasks | jq -r '.[] | select(.due_date < "2026-01-20") | .entry_id'); do
  ./outlook.sh complete-task --id "$id"
done
```

---

## Tips

1. **Use --all flag for calendar** to see all events when date filtering isn't working
2. **Folder names** are case-sensitive for move command
3. **Dates** must be in format: `YYYY-MM-DD HH:MM:SS` for appointments, `YYYY-MM-DD` for tasks
4. **Entry IDs** are long - use quotes when passing them
5. **Test first** - create test items before operating on important data

---

## Getting Help

```bash
# Show all commands
./outlook.sh

# Show help for specific command (not yet implemented, but coming)
./outlook.sh <command> --help
```
