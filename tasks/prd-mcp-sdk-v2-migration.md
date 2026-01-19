# PRD: MCP SDK v2 Migration for Mailtool

## Introduction

Migrate the mailtool MCP server from a hand-rolled JSON-RPC implementation to the official **MCP Python SDK v2** with FastMCP framework. This migration will replace manual JSON-RPC handling with declarative decorators, add structured output using Pydantic models, implement MCP Resources for read-only data access, and improve error handling—all while maintaining 100% feature parity with the existing 23 tools.

**Current State**: `mcp_server.py` (856 lines) with manual JSON-RPC implementation
**Target State**: `src/mailtool/mcp/` package with FastMCP decorators and Pydantic models

**Key Benefits**:
- ~70% less code through declarative patterns
- Type safety via automatic schema generation
- Better separation of concerns (tools, resources, models, lifespan)
- Production-ready patterns for future enhancements

---

## Goals

- ✅ Replace hand-rolled JSON-RPC implementation with FastMCP decorators
- ✅ Add structured output for all 23 tools using Pydantic models
- ✅ Implement MCP Resources for email, calendar, and tasks (read-only data access)
- ✅ Achieve 100% feature parity with existing implementation
- ✅ Add comprehensive error handling with custom exception types
- ✅ Maintain COM threading model (apartment-threaded)
- ✅ Provide full test coverage for all tools and resources
- ✅ Complete migration in single 3-4 week release

---

## User Stories

### Foundation & Infrastructure

### US-001: Set up MCP SDK infrastructure
**Description:** As a developer, I need to add MCP SDK dependencies and create the new package structure so that I can start building the FastMCP-based server.

**Acceptance Criteria:**
- [ ] Add `mcp>=0.9.0` to `pyproject.toml` dependencies
- [ ] Create `src/mailtool/mcp/` package with `__init__.py`
- [ ] Create empty `server.py`, `models.py`, `resources.py`, `lifespan.py` files
- [ ] Update `pyproject.toml` version to 2.3.0
- [ ] Run `uv sync --all-groups` successfully
- [ ] Commit: "feat(mcp): add MCP SDK v2 dependency and package structure"

### US-002: Implement lifespan management
**Description:** As a developer, I need to manage Outlook bridge lifecycle so that COM objects are properly initialized and cleaned up.

**Acceptance Criteria:**
- [ ] Create `OutlookContext` dataclass with bridge attribute
- [ ] Implement `outlook_lifespan()` async context manager
- [ ] Create `OutlookBridge` instance on startup with warmup (5 retry attempts)
- [ ] Release COM objects and force garbage collection on shutdown
- [ ] Add warmup connection test (real COM call to Inbox.Items.Count)
- [ ] Test lifespan starts and shuts down without errors
- [ ] Typecheck passes

### US-003: Create FastMCP server skeleton
**Description:** As a developer, I need to create the basic FastMCP server instance so that tools and resources can be registered.

**Acceptance Criteria:**
- [ ] Create FastMCP instance with name "mailtool-outlook-bridge"
- [ ] Attach `outlook_lifespan` to server
- [ ] Add basic `if __name__ == "__main__"` block calling `mcp.run()`
- [ ] Test server starts via `uv run --with mcp --with pywin32 -m mailtool.mcp.server`
- [ ] Verify server responds to MCP initialize handshake
- [ ] Typecheck passes

---

### Pydantic Models

### US-004: Define email Pydantic models
**Description:** As a developer, I need strongly-typed email models so that tools return validated structured output.

**Acceptance Criteria:**
- [ ] Create `EmailSummary` model (entry_id, subject, sender, sender_name, received_time, unread, has_attachments)
- [ ] Create `EmailDetails` model (extends EmailSummary with body, html_body)
- [ ] Create `SendEmailResult` model (success, entry_id, message)
- [ ] All fields have descriptive Field() descriptions for LLM understanding
- [ ] Optional fields match bridge behavior (received_time, unread can be None)
- [ ] Add model validation tests in `tests/mcp/test_models.py`
- [ ] Typecheck passes

### US-005: Define calendar Pydantic models
**Description:** As a developer, I need strongly-typed calendar models so that appointment tools return structured output.

**Acceptance Criteria:**
- [ ] Create `AppointmentSummary` model (entry_id, subject, start, end, location, organizer, all_day, required_attendees, optional_attendees, response_status, meeting_status, response_requested)
- [ ] Create `AppointmentDetails` model (extends AppointmentSummary with body)
- [ ] Create `CreateAppointmentResult` model (success, entry_id, message)
- [ ] Create `FreeBusyInfo` model (email, start_date, end_date, free_busy, resolved, error)
- [ ] All fields have descriptive Field() descriptions
- [ ] Optional fields match bridge behavior (start, end, organizer can be None)
- [ ] Add model validation tests
- [ ] Typecheck passes

### US-006: Define task Pydantic models
**Description:** As a developer, I need strongly-typed task models so that task tools return structured output.

**Acceptance Criteria:**
- [ ] Create `TaskSummary` model (entry_id, subject, body, due_date, status, priority, complete, percent_complete)
- [ ] Create `CreateTaskResult` model (success, entry_id, message)
- [ ] Handle None values for status and priority (bridge returns None for some tasks)
- [ ] All fields have descriptive Field() descriptions
- [ ] Add model validation tests
- [ ] Typecheck passes

### US-007: Define common result models
**Description:** As a developer, I need generic result models so that boolean operations return consistent structured output.

**Acceptance Criteria:**
- [ ] Create `OperationResult` model (success, message)
- [ ] Use for all tools that currently return True/False
- [ ] Add model validation tests
- [ ] Typecheck passes

---

### Simple Tool Migration (8 tools)

### US-008: Implement get_email tool
**Description:** As a Claude Code user, I want to retrieve full email details with structured output so that I can read email content.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` decorated function `get_email(entry_id: str)`
- [ ] Access bridge from lifespan context
- [ ] Convert bridge dict to `EmailDetails` Pydantic model
- [ ] Raise `McpError` if email not found
- [ ] Handle missing `unread` field (bridge doesn't return it in get_email_body)
- [ ] Return structured `EmailDetails` object
- [ ] Add test verifying structured output in `tests/mcp/test_tools.py`
- [ ] Manual test with Claude Code: retrieve real email
- [ ] Typecheck passes

### US-009: Implement get_appointment tool
**Description:** As a Claude Code user, I want to retrieve calendar event details so that I can view meeting information.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function `get_appointment(entry_id: str)`
- [ ] Convert bridge dict to `AppointmentDetails` model
- [ ] Raise `McpError` if appointment not found
- [ ] Return structured `AppointmentDetails` object
- [ ] Add test verifying structured output
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-010: Implement get_task tool
**Description:** As a Claude Code user, I want to retrieve task details so that I can view task information.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function `get_task(entry_id: str)`
- [ ] Convert bridge dict to `TaskSummary` model
- [ ] Handle None status and priority
- [ ] Raise `McpError` if task not found
- [ ] Return structured `TaskSummary` object
- [ ] Add test verifying structured output
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-011: Implement mark_email tool
**Description:** As a Claude Code user, I want to mark emails as read/unread so that I can manage my inbox.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function `mark_email(entry_id: str, unread: bool = False)`
- [ ] Call `bridge.mark_email_read(entry_id, unread=unread)`
- [ ] Return `OperationResult(success=True, message="Email marked as read/unread")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add test for both read and unread cases
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-012: Implement delete_email tool
**Description:** As a Claude Code user, I want to delete emails so that I can clean my inbox.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function `delete_email(entry_id: str)`
- [ ] Call `bridge.delete_email(entry_id)`
- [ ] Return `OperationResult(success=True, message="Email deleted successfully")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add test verifying deletion
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-013: Implement delete_appointment tool
**Description:** As a Claude Code user, I want to delete calendar events so that I can remove cancelled meetings.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function `delete_appointment(entry_id: str)`
- [ ] Call `bridge.delete_appointment(entry_id)`
- [ ] Return `OperationResult(success=True, message="Appointment deleted successfully")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add test verifying deletion
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-014: Implement complete_task tool
**Description:** As a Claude Code user, I want to mark tasks as complete so that I can track my progress.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function `complete_task(entry_id: str)`
- [ ] Call `bridge.complete_task(entry_id)`
- [ ] Return `OperationResult(success=True, message="Task completed successfully")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add test verifying completion
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-015: Implement delete_task tool
**Description:** As a Claude Code user, I want to delete tasks so that I can remove obsolete items.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function `delete_task(entry_id: str)`
- [ ] Call `bridge.delete_task(entry_id)`
- [ ] Return `OperationResult(success=True, message="Task deleted successfully")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add test verifying deletion
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

---

### Email Tools & Resources (9 tools total)

### US-016: Implement list_emails tool
**Description:** As a Claude Code user, I want to list recent emails so that I can see what's in my inbox.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function `list_emails(limit: int = 10, folder: str = "Inbox")`
- [ ] Call `bridge.list_emails(limit=limit, folder=folder)`
- [ ] Convert list of dicts to list of `EmailSummary` models
- [ ] Handle missing `received_time` (set to None)
- [ ] Return list of `EmailSummary` objects
- [ ] Add test verifying list output and limit parameter
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-017: Implement send_email tool
**Description:** As a Claude Code user, I want to send emails so that I can communicate with others.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with params: to, subject, body, cc, bcc, html_body, file_paths, save_draft
- [ ] Call `bridge.send_email()` with all parameters
- [ ] Handle three return types: False (failed), True (sent), str (draft EntryID)
- [ ] Return `SendEmailResult` with appropriate success/message
- [ ] Add tests for sent email and draft saved
- [ ] Manual test with Claude Code (send draft)
- [ ] Typecheck passes

### US-018: Implement reply_email tool
**Description:** As a Claude Code user, I want to reply to emails so that I can respond to messages.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with params: entry_id, body, reply_all, html_body
- [ ] Call `bridge.reply_email(entry_id, body, reply_all=reply_all, html_body=html_body)`
- [ ] Return `OperationResult(success=True, message="Reply sent successfully")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add tests for reply and reply_all
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-019: Implement forward_email tool
**Description:** As a Claude Code user, I want to forward emails so that I can share messages with others.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with params: entry_id, to, body, html_body
- [ ] Call `bridge.forward_email(entry_id, to, body, html_body=html_body)`
- [ ] Return `OperationResult(success=True, message="Email forwarded successfully")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add test verifying forward
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-020: Implement move_email tool
**Description:** As a Claude Code user, I want to move emails to folders so that I can organize my inbox.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with params: entry_id, folder
- [ ] Call `bridge.move_email(entry_id, folder)`
- [ ] Return `OperationResult(success=True, message="Email moved successfully")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add test verifying move
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-021: Implement search_emails tool
**Description:** As a Claude Code user, I want to search emails so that I can find specific messages.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with params: filter_query, limit
- [ ] Call `bridge.search_emails(filter_query=filter_query, limit=limit)`
- [ ] Convert results to list of `EmailSummary` models
- [ ] Add examples in docstring (e.g., "[Subject] = 'invoice'", "urn:schemas:httpmail:subject LIKE '%test%'")
- [ ] Add test verifying search
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

---

### Email Resources

### US-022: Implement inbox://emails resource
**Description:** As a Claude Code user, I want read-only access to my inbox emails so that I can get context without calling tools.

**Acceptance Criteria:**
- [ ] Create `@mcp.resource("inbox://emails")` function
- [ ] Return JSON string of 50 most recent emails
- [ ] Use `json.dumps(emails, indent=2)` for formatting
- [ ] Add test verifying resource access via MCP client
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-023: Implement inbox://unread resource
**Description:** As a Claude Code user, I want read-only access to unread emails so that I can focus on new messages.

**Acceptance Criteria:**
- [ ] Create `@mcp.resource("inbox://unread")` function
- [ ] Fetch 1000 emails and filter to unread only
- [ ] Return JSON string of unread emails
- [ ] Add test verifying resource access
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-024: Implement email://{entry_id} resource
**Description:** As a Claude Code user, I want read-only access to specific emails so that I can reference them by EntryID.

**Acceptance Criteria:**
- [ ] Create `@mcp.resource("email://{entry_id}")` function with entry_id parameter
- [ ] Call `bridge.get_email_body(entry_id)`
- [ ] Return JSON string of email or error if not found
- [ ] Add test for both found and not found cases
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

---

### Calendar Tools & Resources (7 tools total)

### US-025: Implement list_calendar_events tool
**Description:** As a Claude Code user, I want to list calendar events so that I can see my schedule.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with params: days (default 7), all (default False)
- [ ] Call `bridge.list_calendar_events(days=days, all_events=all)`
- [ ] Convert results to list of `AppointmentSummary` models
- [ ] Handle None values for start, end, organizer
- [ ] Add test for both days and all parameters
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-026: Implement create_appointment tool
**Description:** As a Claude Code user, I want to create calendar appointments so that I can schedule meetings.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with params: subject, start, end, location, body, all_day, required_attendees, optional_attendees
- [ ] Call `bridge.create_appointment()` with all parameters
- [ ] Return `CreateAppointmentResult` with entry_id if successful
- [ ] Return failure result if bridge returns None
- [ ] Add test verifying appointment creation
- [ ] Manual test with Claude Code (create real appointment)
- [ ] Typecheck passes

### US-027: Implement edit_appointment tool
**Description:** As a Claude Code user, I want to edit existing appointments so that I can update meeting details.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with all appointment fields as optional params
- [ ] Call `bridge.edit_appointment(entry_id, **kwargs)`
- [ ] Return `OperationResult(success=True, message="Appointment updated successfully")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add test verifying edit
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-028: Implement respond_to_meeting tool
**Description:** As a Claude Code user, I want to respond to meeting invitations so that I can accept/decline meetings.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with params: entry_id, response (accept/decline/tentative)
- [ ] Call `bridge.respond_to_meeting(entry_id, response=response)`
- [ ] Return `OperationResult(success=True, message="Meeting response sent successfully")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add tests for accept, decline, tentative
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-029: Implement get_free_busy tool
**Description:** As a Claude Code user, I want to check availability so that I can schedule meetings.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with params: email_address, start_date, end_date, duration
- [ ] Call `bridge.get_free_busy()` with all parameters
- [ ] Convert result to `FreeBusyInfo` model
- [ ] Handle failure case where bridge returns {email, error, resolved} without free_busy
- [ ] Add test for successful and failed lookups
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

---

### Calendar Resources

### US-030: Implement calendar://today resource
**Description:** As a Claude Code user, I want read-only access to today's events so that I can see my schedule.

**Acceptance Criteria:**
- [ ] Create `@mcp.resource("calendar://today")` function
- [ ] Call `bridge.list_calendar_events(days=1)`
- [ ] Return JSON string of events
- [ ] Add test verifying resource access
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-031: Implement calendar://week resource
**Description:** As a Claude Code user, I want read-only access to this week's events so that I can plan ahead.

**Acceptance Criteria:**
- [ ] Create `@mcp.resource("calendar://week")` function
- [ ] Call `bridge.list_calendar_events(days=7)`
- [ ] Return JSON string of events
- [ ] Add test verifying resource access
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

---

### Task Tools & Resources (7 tools total)

### US-032: Implement list_tasks tool
**Description:** As a Claude Code user, I want to list incomplete tasks so that I can see what I need to do.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function `list_tasks()`
- [ ] Call `bridge.list_tasks(include_completed=False)`
- [ ] Convert results to list of `TaskSummary` models
- [ ] Handle None status and priority
- [ ] Add test verifying incomplete-only behavior
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-033: Implement list_all_tasks tool
**Description:** As a Claude Code user, I want to list all tasks so that I can see my complete task list.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function `list_all_tasks()`
- [ ] Call `bridge.list_tasks(include_completed=True)`
- [ ] Convert results to list of `TaskSummary` models
- [ ] Add test verifying all tasks returned
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-034: Implement create_task tool
**Description:** As a Claude Code user, I want to create tasks so that I can track my to-dos.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with params: subject, body, due_date, priority, status
- [ ] Call `bridge.create_task()` with all parameters
- [ ] Return `CreateTaskResult` with entry_id if successful
- [ ] Return failure result if bridge returns None
- [ ] Add test verifying task creation
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-035: Implement edit_task tool
**Description:** As a Claude Code user, I want to edit tasks so that I can update task details.

**Acceptance Criteria:**
- [ ] Create `@mcp.tool()` function with all task fields as optional params
- [ ] Call `bridge.edit_task(entry_id, **kwargs)`
- [ ] Return `OperationResult(success=True, message="Task updated successfully")`
- [ ] Raise `McpError` if bridge returns False
- [ ] Add test verifying edit
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

---

### Task Resources

### US-036: Implement tasks://active resource
**Description:** As a Claude Code user, I want read-only access to active tasks so that I can see what's pending.

**Acceptance Criteria:**
- [ ] Create `@mcp.resource("tasks://active")` function
- [ ] Call `bridge.list_tasks(include_completed=False)`
- [ ] Return JSON string of tasks
- [ ] Add test verifying resource access
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

### US-037: Implement tasks://all resource
**Description:** As a Claude Code user, I want read-only access to all tasks so that I can see my complete task list.

**Acceptance Criteria:**
- [ ] Create `@mcp.resource("tasks://all")` function
- [ ] Call `bridge.list_tasks(include_completed=True)`
- [ ] Return JSON string of tasks
- [ ] Add test verifying resource access
- [ ] Manual test with Claude Code
- [ ] Typecheck passes

---

### Error Handling & Polish

### US-038: Create custom exception types
**Description:** As a developer, I need custom exception types so that users get clear, actionable error messages.

**Acceptance Criteria:**
- [ ] Create `OutlookNotFoundError(McpError)` for missing items
- [ ] Create `OutlookComError(McpError)` for COM failures
- [ ] Create `OutlookValidationError(McpError)` for invalid input
- [ ] Add `src/mailtool/mcp/exceptions.py` module
- [ ] Update all tools to use custom exceptions
- [ ] Add tests for each exception type
- [ ] Typecheck passes

### US-039: Add comprehensive logging
**Description:** As a developer, I need logging so that I can debug issues in production.

**Acceptance Criteria:**
- [ ] Add `logging` import to server module
- [ ] Configure logger with appropriate level
- [ ] Log lifespan events (startup, shutdown)
- [ ] Log tool calls with parameters
- [ ] Log errors with stack traces
- [ ] Test logging output during tool execution
- [ ] Typecheck passes

---

### Testing Infrastructure

### US-040: Create MCP test fixtures
**Description:** As a developer, I need test fixtures so that I can test MCP tools efficiently.

**Acceptance Criteria:**
- [ ] Create `tests/mcp/conftest.py` with fixtures:
  - `mcp_server()` - provides FastMCP instance
  - `outlook_bridge()` - provides OutlookBridge instance
  - `sample_email()` - creates test email with cleanup
  - `sample_appointment()` - creates test appointment with cleanup
  - `sample_task()` - creates test task with cleanup
- [ ] All fixtures include proper cleanup
- [ ] Fixtures use [TEST] prefix for easy identification
- [ ] Run `pytest --fixtures` to verify fixtures are registered

### US-041: Create model validation tests
**Description:** As a developer, I need to test Pydantic models so that I know they validate data correctly.

**Acceptance Criteria:**
- [ ] Create `tests/mcp/test_models.py`
- [ ] Test all email models with valid data
- [ ] Test all calendar models with valid data
- [ ] Test all task models with valid data
- [ ] Test missing optional fields don't cause errors
- [ ] Test invalid data raises ValidationError
- [ ] All tests pass: `uv run pytest tests/mcp/test_models.py -v`

### US-042: Create tool integration tests
**Description:** As a developer, I need to test all 23 tools via MCP client so that I know they work end-to-end.

**Acceptance Criteria:**
- [ ] Create `tests/mcp/test_tools.py`
- [ ] Test all 23 tools using MCP ClientSession
- [ ] Test structured output format for each tool
- [ ] Test error handling for invalid inputs
- [ ] Test tools return expected Pydantic models
- [ ] Mark slow tests with `@pytest.mark.slow`
- [ ] All tests pass: `uv run pytest tests/mcp/test_tools.py -v`

### US-043: Create resource integration tests
**Description:** As a developer, I need to test MCP resources so that I know they return data correctly.

**Acceptance Criteria:**
- [ ] Create `tests/mcp/test_resources.py`
- [ ] Test all 7 resources via MCP ClientSession
- [ ] Test resources return valid JSON
- [ ] Test resources handle not-found cases
- [ ] All tests pass: `uv run pytest tests/mcp/test_resources.py -v`

### US-044: Create end-to-end workflow tests
**Description:** As a developer, I need to test complete workflows so that I know tools work together.

**Acceptance Criteria:**
- [ ] Create `tests/mcp/test_integration.py`
- [ ] Test email workflow (list, create draft, get, delete)
- [ ] Test calendar workflow (list, create, edit, respond, delete)
- [ ] Test task workflow (list, create, edit, complete, delete)
- [ ] Mark with `@pytest.mark.integration` and `@pytest.mark.slow`
- [ ] All tests pass: `uv run pytest -m integration -v`

---

### Documentation & Deployment

### US-045: Update README.md
**Description:** As a user, I need updated documentation so that I can understand the new MCP server architecture.

**Acceptance Criteria:**
- [ ] Add "MCP Server (v2 with FastMCP)" section
- [ ] Document benefits: structured output, resources, type safety
- [ ] Link to MCP_SDK_V2_MIGRATION_PLAN.md
- [ ] Update installation instructions (still uses `uv run --with pywin32`)
- [ ] Update examples to show structured output
- [ ] Verify README renders correctly

### US-046: Update CLAUDE.md
**Description:** As an AI assistant, I need updated architecture docs so that I can understand the new codebase.

**Acceptance Criteria:**
- [ ] Update MCP Server section with FastMCP details
- [ ] Document new file structure (src/mailtool/mcp/)
- [ ] Document Pydantic models and their purpose
- [ ] Document MCP resources and URI patterns
- [ ] Update tool list with return types
- [ ] Add migration notes section

### US-047: Update MCP_INTEGRATION.md
**Description:** As a developer, I need updated MCP documentation so that I can integrate with the new server.

**Acceptance Criteria:**
- [ ] Update architecture section with FastMCP details
- [ ] Document structured output format
- [ ] Document all MCP resources with examples
- [ ] Add troubleshooting section for SDK-specific issues
- [ ] Update examples to show new patterns
- [ ] Verify all links work

### US-048: Update plugin.json configuration
**Description:** As a user, I need the plugin configuration updated so that Claude Code loads the new server.

**Acceptance Criteria:**
- [ ] Update `.claude-plugin/plugin.json` command to: `uv run --with pywin32 -m mailtool.mcp.server`
- [ ] Bump version to 2.3.0
- [ ] Update description to mention MCP SDK v2
- [ ] Add PYTHONUNBUFFERED=1 to env
- [ ] Test plugin loads in Claude Code
- [ ] Verify all 23 tools are discoverable

### US-049: Create migration guide for users
**Description:** As a user, I need migration instructions so that I can update my plugin configuration.

**Acceptance Criteria:**
- [ ] Create `docs/MIGRATION_V2.3.md`
- [ ] Explain breaking change (plugin.json command change)
- [ ] Provide step-by-step update instructions
- [ ] Show before/after plugin.json examples
- [ ] Document new features (structured output, resources)
- [ ] Link from README.md

### US-050: Verify old server still works (rollback)
**Description:** As a developer, I need to verify the old server still works so that I can rollback if needed.

**Acceptance Criteria:**
- [ ] Keep `mcp_server.py` (old version) unchanged
- [ ] Test old server starts: `uv run --with pywin32 mcp_server.py`
- [ ] Test old server with MCP inspector
- [ ] Verify all 23 tools work in old server
- [ ] Document rollback procedure in MIGRATION_V2.3.md
- [ ] Commit old server backup before cutover

---

## Functional Requirements

### SDK Infrastructure
- FR-1: Add `mcp>=0.9.0` dependency to pyproject.toml
- FR-2: Create `src/mailtool/mcp/` package with server, models, resources, lifespan, exceptions modules
- FR-3: Implement `outlook_lifespan()` async context manager for COM lifecycle
- FR-4: Create FastMCP server instance with name "mailtool-outlook-bridge"
- FR-5: Server must start via `uv run --with pywin32 -m mailtool.mcp.server`

### Data Models
- FR-6: Define `EmailSummary` model with entry_id, subject, sender, sender_name, received_time, unread, has_attachments
- FR-7: Define `EmailDetails` model extending EmailSummary with body, html_body
- FR-8: Define `SendEmailResult` model with success, entry_id, message
- FR-9: Define `AppointmentSummary` model with all calendar fields
- FR-10: Define `AppointmentDetails` model extending AppointmentSummary with body
- FR-11: Define `CreateAppointmentResult` model with success, entry_id, message
- FR-12: Define `FreeBusyInfo` model with email, start_date, end_date, free_busy, resolved, error
- FR-13: Define `TaskSummary` model with entry_id, subject, body, due_date, status, priority, complete, percent_complete
- FR-14: Define `CreateTaskResult` model with success, entry_id, message
- FR-15: Define `OperationResult` model with success, message
- FR-16: All Pydantic fields must have descriptive Field() descriptions for LLM understanding

### Email Tools (9 tools)
- FR-17: `list_emails(limit, folder)` returns list[EmailSummary]
- FR-18: `get_email(entry_id)` returns EmailDetails, raises McpError if not found
- FR-19: `send_email(to, subject, body, cc, bcc, html_body, file_paths, save_draft)` returns SendEmailResult
- FR-20: `reply_email(entry_id, body, reply_all, html_body)` returns OperationResult
- FR-21: `forward_email(entry_id, to, body, html_body)` returns OperationResult
- FR-22: `mark_email(entry_id, unread)` returns OperationResult
- FR-23: `move_email(entry_id, folder)` returns OperationResult
- FR-24: `delete_email(entry_id)` returns OperationResult
- FR-25: `search_emails(filter_query, limit)` returns list[EmailSummary]

### Calendar Tools (7 tools)
- FR-26: `list_calendar_events(days, all)` returns list[AppointmentSummary]
- FR-27: `create_appointment(subject, start, end, location, body, all_day, required_attendees, optional_attendees)` returns CreateAppointmentResult
- FR-28: `get_appointment(entry_id)` returns AppointmentDetails, raises McpError if not found
- FR-29: `edit_appointment(entry_id, **kwargs)` returns OperationResult
- FR-30: `respond_to_meeting(entry_id, response)` returns OperationResult
- FR-31: `delete_appointment(entry_id)` returns OperationResult
- FR-32: `get_free_busy(email_address, start_date, end_date, duration)` returns FreeBusyInfo

### Task Tools (7 tools)
- FR-33: `list_tasks()` returns list[TaskSummary] (incomplete only)
- FR-34: `list_all_tasks()` returns list[TaskSummary] (all tasks)
- FR-35: `create_task(subject, body, due_date, priority, status)` returns CreateTaskResult
- FR-36: `get_task(entry_id)` returns TaskSummary, raises McpError if not found
- FR-37: `edit_task(entry_id, **kwargs)` returns OperationResult
- FR-38: `complete_task(entry_id)` returns OperationResult
- FR-39: `delete_task(entry_id)` returns OperationResult

### MCP Resources (7 resources)
- FR-40: `inbox://emails` resource returns 50 recent emails as JSON
- FR-41: `inbox://unread` resource returns unread emails as JSON
- FR-42: `email://{entry_id}` resource returns specific email as JSON
- FR-43: `calendar://today` resource returns today's events as JSON
- FR-44: `calendar://week` resource returns this week's events as JSON
- FR-45: `tasks://active` resource returns active tasks as JSON
- FR-46: `tasks://all` resource returns all tasks as JSON

### Error Handling
- FR-47: Create `OutlookNotFoundError` for missing items
- FR-48: Create `OutlookComError` for COM failures
- FR-49: Create `OutlookValidationError` for invalid inputs
- FR-50: All tools must raise appropriate McpError exceptions on failure

### Testing
- FR-51: All Pydantic models must have validation tests
- FR-52: All 23 tools must have integration tests via MCP ClientSession
- FR-53: All 7 resources must have integration tests
- FR-54: End-to-end workflow tests for email, calendar, and tasks
- FR-55: Test suite must pass: `uv run pytest tests/ -v`

### Documentation
- FR-56: Update README.md with FastMCP architecture details
- FR-57: Update CLAUDE.md with new file structure and patterns
- FR-58: Update MCP_INTEGRATION.md with resource documentation
- FR-59: Create MIGRATION_V2.3.md with user migration guide
- FR-60: Update .claude-plugin/plugin.json to point to new server entry point

---

## Non-Goals (Out of Scope)

- ❌ Adding new MCP tools beyond the existing 23
- ❌ Implementing MCP Prompts (defer to future release)
- ❌ Adding authentication/authorization (defer to future release)
- ❌ Implementing monitoring/metrics (defer to future release)
- ❌ Performance optimization beyond maintaining current speeds
- ❌ Supporting multiple Outlook profiles simultaneously
- ❌ Backwards compatibility with old plugin.json format (manual update required)
- ❌ Automated migration scripts for user plugin configurations
- ❌ Removing or changing the bridge.py implementation
- ❌ Modifying existing test structure (bridge tests stay as-is)

---

## Design Considerations

### File Structure
- Keep `mcp_server.py` (old version) for rollback
- New code goes in `src/mailtool/mcp/` package
- Clear separation: models, resources, server, lifespan, exceptions

### Type Safety
- All tools use type hints for automatic schema generation
- Pydantic models provide runtime validation
- Field descriptions critical for LLM understanding

### COM Threading
- All COM calls must stay on same thread
- Lifespan context ensures single bridge instance
- No use of thread pools or parallel execution in tools

### Error Messages
- Custom exceptions provide clear, actionable errors
- User-friendly messages for common failures
- Stack traces logged for debugging

---

## Technical Considerations

### Dependencies
- `mcp>=0.9.0` for SDK runtime
- Python 3.13+ (existing requirement)
- `pywin32>=306` on Windows only (existing pattern)

### COM Lifecycle
- Bridge created once during lifespan startup
- Warmup ensures Outlook is responsive before serving requests
- Garbage collection on shutdown prevents COM leaks

### Testing Strategy
- Use MCP ClientSession for integration tests
- Stdio transport to test real server behavior
- Fixtures with auto-cleanup for test data
- Mark slow tests to allow quick CI runs

### Deployment
- Update plugin.json command to new entry point
- Users must manually update configuration
- Keep old server as backup for rollback
- Document rollback procedure clearly

---

## Success Metrics

- ✅ All 23 tools return structured Pydantic output
- ✅ All 7 resources accessible via MCP
- ✅ 100% test coverage for all tools and resources
- ✅ Zero regression in existing functionality
- ✅ Code reduction: ~70% less than manual implementation
- ✅ All tests pass: `uv run pytest tests/ -v`
- ✅ Manual testing with Claude Code successful
- ✅ Documentation complete and accurate
- ✅ Rollback plan tested and documented

---

## Open Questions

1. Should we add `tasks://overdue` resource? (Plan mentions it but doesn't specify implementation)
2. Should we add `calendar://{date}` resource for specific dates? (Plan mentions it but doesn't specify implementation)
3. Should `send_email` expose `file_paths` parameter? (Bridge supports it, current MCP tool doesn't)
4. What timeout values should we use for COM warmup retries? (Currently hardcoded 0.5s)
5. Should we add structured logging format (JSON) for production monitoring?

---

## Implementation Order

**Week 1: Foundation**
1. US-001 to US-007: SDK setup, models, lifespan, server skeleton
2. US-038 to US-039: Custom exceptions, logging
3. US-040 to US-041: Test fixtures, model tests

**Week 2: Core Email**
1. US-008 to US-015: Simple tools (get, mark, delete for email, calendar, tasks)
2. US-016 to US-021: Email tools
3. US-022 to US-024: Email resources
4. US-042: Email tool tests

**Week 3: Calendar & Tasks**
1. US-025 to US-029: Calendar tools
2. US-030 to US-031: Calendar resources
3. US-032 to US-037: Task tools and resources
4. US-043 to US-044: Resource and integration tests

**Week 4: Polish & Deploy**
1. US-045 to US-049: Documentation updates
2. US-050: Rollback verification
3. Final testing and bug fixes
4. Production cutover

---

**Document Version**: 1.0
**Created**: 2025-01-19
**Author**: Generated from MCP_SDK_V2_MIGRATION_PLAN.md
**Status**: Ready for Implementation
