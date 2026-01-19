# Mailtool - Outlook Automation Bridge

A Python library and CLI tool for accessing Outlook email, calendar, and tasks from WSL2 via Windows COM automation.

**Uses [uv](https://github.com/astral-sh/uv) for dependency management - no global Python needed!**

## 🚀 Installation

### PyPI Installation (Recommended)

```bash
# Install via pip
pip install mailtool

# Or via uv
uv pip install mailtool
```

### Claude Code Integration

For Claude Code integration, install the [mailtool-plugin](https://github.com/utsmok/mailtool-plugin) package:

```bash
/plugin marketplace add utsmok/mailtool-plugin
/plugin install mailtool
```

This will configure the MCP server with 24 tools for email, calendar, and task management.

## Prerequisites

- Windows with Outlook (classic) installed and running
- WSL2 with `uv` installed (`pip install uv` or `curl -LsSf https://astral.sh/uv/install.sh | sh`)
- `uv.exe` accessible from WSL2 (automatically available if installed on Windows)

## Setup

### 1. Start Outlook

Make sure Outlook is running and logged into your `s.mok@utwente.nl` account.

### 2. That's it!

Dependencies are managed automatically by `uv`. No manual pip installs needed.

## Usage

### As a Python Library

```python
from mailtool.bridge import OutlookBridge

# Create bridge instance
bridge = OutlookBridge()

# List emails
emails = bridge.list_emails(limit=5)
for email in emails:
    print(f"{email['subject']}: {email['sender']}")

# Create appointment
entry_id = bridge.create_appointment(
    subject="Team Meeting",
    start="2025-01-20 14:00:00",
    end="2025-01-20 15:00:00",
    location="Room 101"
)
```

### As a CLI Tool

```bash
# List recent emails
mailtool emails --limit 5

# List calendar events for next 7 days
mailtool calendar --days 7

# Get specific email body (use entry_id from emails command)
mailtool email --id <entry_id>
```

### As an MCP Server (for Claude Code)

See the [mailtool-plugin](https://github.com/utsmok/mailtool-plugin) repository for Claude Code integration instructions.

## How It Works

The library uses Windows COM automation to communicate with Outlook:

1. Python creates a COM object to access the running Outlook instance
2. Uses O(1) direct lookups via `GetItemFromID()` for performance
3. Returns structured data (emails, calendar events, tasks) as Python dictionaries
4. MCP server mode exposes this functionality via JSON-RPC for AI agents

## Project Structure

```
mailtool/
├── pyproject.toml          # uv project config
├── src/
│   └── mailtool/
│       ├── __init__.py
│       ├── bridge.py       # Core COM automation (~1100 lines)
│       ├── cli.py          # CLI interface
│       └── mcp/            # MCP Server (SDK v2 + FastMCP)
│           ├── __init__.py
│           ├── server.py   # FastMCP server with 24 tools
│           ├── models.py   # Pydantic models
│           ├── lifespan.py # Async COM bridge lifecycle
│           ├── resources.py # 7 resources
│           └── exceptions.py # Custom exceptions
├── tests/
│   ├── conftest.py         # Test fixtures
│   ├── test_bridge.py      # Core connectivity tests
│   ├── test_emails.py      # Email operation tests
│   ├── test_calendar.py    # Calendar operation tests
│   ├── test_tasks.py       # Task operation tests
│   └── mcp/                # MCP server tests
│       ├── test_models.py      # Pydantic model tests
│       ├── test_tools.py       # Tool implementation tests
│       ├── test_resources.py   # Resource tests
│       ├── test_integration.py # End-to-end workflow tests
│       └── test_exceptions.py  # Exception class tests
├── docs/                   # Documentation
└── .github/
    └── workflows/
        ├── ci.yml          # Continuous Integration
        └── publish.yml     # PyPI publishing
```

## Advantages

- ✅ **uv for dependencies** - No global Python pollution
- ✅ **Official MCP SDK v2** - Type-safe, well-documented, maintainable
- ✅ **Structured output** - Pydantic models for all tool results
- ✅ **7 Resources** - Quick data access without tool calls
- ✅ **No API registration** - Uses existing Outlook auth
- ✅ **Works with any Outlook account**
- ✅ **Full access** to email, calendar, and tasks
- ✅ **Stable** - Doesn't break on UI changes
- ✅ **Cross-shell** - Works from WSL2, PowerShell, etc.

## Limitations

- ⚠️ Outlook must be running on Windows
- ⚠️ Windows-specific (COM automation)
- ⚠️ MCP server requires Windows with Outlook (works from WSL2/Linux clients)

## Claude Code Integration (MCP)

**NEW: v2.3.0 - Now powered by MCP Python SDK v2 with FastMCP framework!**

This includes a Model Context Protocol (MCP) server for Claude Code integration using the official MCP Python SDK v2 and FastMCP framework.

### Key Features

- **24 Tools** for email, calendar, and task management
- **7 Resources** for quick data access (inbox, calendar, tasks)
- **Structured Output** - All tools return typed Pydantic models
- **Type Safety** - Full type annotations for better IDE support
- **Error Handling** - Custom exception classes with detailed error messages
- **Logging** - Comprehensive logging for debugging and monitoring
- **Zero-Config** - Uses `uv run --with pywin32` for dependency-free execution

### Manual Installation

If you prefer manual installation or want to contribute:

```bash
# Clone the repository
git clone https://github.com/utsmok/mailtool.git
cd mailtool

# Install in editable mode
uv pip install -e .
```

Then Claude Code can:
- 📧 Read, send, reply to, forward, move, search, and manage emails
- 📅 View, create, edit, respond to meetings, check free/busy, and manage appointments
- ✅ Create, edit, complete, delete, and manage tasks

### MCP Server Architecture

**Version 2.3.0** uses the official MCP Python SDK v2 with FastMCP framework:

```
Claude Code (WSL2/Linux)
    ↓ (JSON-RPC via stdio)
FastMCP Server (mailtool.mcp.server)
    ↓ (async context manager)
Outlook COM Bridge (thread pool executor)
    ↓ (COM)
Outlook Application
```

**Key improvements from v2.2:**
- ✅ Official MCP SDK v2 (mcp>=0.9.0) with FastMCP framework
- ✅ Structured Pydantic models for all tool outputs (EmailDetails, AppointmentDetails, TaskSummary, etc.)
- ✅ 7 resources for quick data access (inbox://emails, calendar://today, tasks://active, etc.)
- ✅ Custom exception classes (OutlookNotFoundError, OutlookComError, OutlookValidationError)
- ✅ Comprehensive logging for debugging and monitoring
- ✅ Type-safe tool definitions with @mcp.tool() decorator
- ✅ Async lifespan management for COM bridge lifecycle

See [MCP_INTEGRATION.md](MCP_INTEGRATION.md) for full documentation.

## Future Directions

This could become:
- **CLI Tool**: Full-featured email/calendar CLI
- **Web App**: Backend for a web interface
- **Library**: Importable Python module

## Troubleshooting

### "Could not connect to Outlook"
- Make sure Outlook is running
- Check that you're logged into your account

### "uv.exe not found"
- Install uv on Windows: `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
- Make sure uv is in your Windows PATH

### UNC path warnings (harmless)
- These appear because of WSL2 → Windows path translation
- Safe to ignore, everything still works

### MCP tools not available after plugin installation
- This may be due to [Claude Code Bug #16143](https://github.com/anthropics/claude-code/issues/16143)
- The MCP server is configured in `.mcp.json` (not inline in `plugin.json`) to work around this issue
- Try restarting Claude Code after plugin installation
- Verify the plugin installed correctly: `/plugin list`
- Check MCP server status: `/mcp`

## Development

```bash
# Add new dependencies
uv add <package>

# Run on Linux/WSL2 (for tooling)
uv run python <script>

# Run on Windows (for COM automation)
./outlook.sh <command>

# Run tests
./run_tests.sh

# Run linter and formatter
uv run ruff check .
uv run ruff format .
```

### MCP Server Development

The MCP server is implemented in `src/mailtool/mcp/` using the official MCP Python SDK v2:

- **server.py** - FastMCP server with 23 tools
- **models.py** - Pydantic models for structured output
- **lifespan.py** - Async context manager for COM bridge lifecycle
- **resources.py** - 7 resources for quick data access
- **exceptions.py** - Custom exception classes

See [CLAUDE.md](CLAUDE.md) for development patterns and architecture.

### Performance Benchmarks

Performance benchmarks are available in `scripts/benchmarks/` for validating MCP server performance:

```bash
# Run performance benchmarks (requires Windows with Outlook running)
uv run --with pytest --with pywin32 python -m scripts.benchmarks.performance_benchmark
```

**Note:** Benchmarks require Windows with Outlook running and pywin32 installed. They cannot run in WSL2 or CI/CD environments without Outlook access.

See [scripts/benchmarks/README.md](scripts/benchmarks/README.md) for benchmark documentation and [scripts/benchmarks/EXPECTED_RESULTS.md](scripts/benchmarks/EXPECTED_RESULTS.md) for expected output format and success criteria.
