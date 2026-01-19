# Mailtool - Outlook Automation

Access your Office 365 email and calendar from WSL2 via Windows Outlook COM automation.

**Uses [uv](https://github.com/astral-sh/uv) for dependency management - no global Python needed!**

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

```bash
# List recent emails
./outlook.sh emails --limit 5

# List calendar events for next 7 days
./outlook.sh calendar --days 7

# Get specific email body (use entry_id from emails command)
./outlook.sh email --id <entry_id>
```

## How It Works

1. WSL2 calls wrapper script (`outlook.sh`)
2. Wrapper calls Windows batch file (`outlook.bat`)
3. Batch file uses `uv run --with pywin32` to execute the Python script
4. Python script uses COM to talk to running Outlook instance
5. Data returned as JSON

## Project Structure

```
mailtool/
├── pyproject.toml          # uv project config
├── outlook.bat             # Windows entry point (uses uv)
├── outlook.sh              # WSL2 wrapper
├── src/
│   └── mailtool/
│       ├── __init__.py
│       ├── bridge.py       # Core COM automation (~1100 lines)
│       ├── cli.py          # CLI interface
│       └── mcp/            # MCP Server (SDK v2 + FastMCP)
│           ├── __init__.py
│           ├── server.py   # FastMCP server with 23 tools
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
├── .claude-plugin/
│   └── plugin.json         # Claude Code plugin manifest
└── .venv/                  # Linux virtualenv (for tooling)
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

- **23 Tools** for email, calendar, and task management
- **7 Resources** for quick data access (inbox, calendar, tasks)
- **Structured Output** - All tools return typed Pydantic models
- **Type Safety** - Full type annotations for better IDE support
- **Error Handling** - Custom exception classes with detailed error messages
- **Logging** - Comprehensive logging for debugging and monitoring
- **Zero-Config** - Uses `uv run --with pywin32` for dependency-free execution

### Installation

```bash
# Clone to your Claude Code plugins directory
git clone <repo> ~/.claude-code/plugins/mailtool
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

Performance benchmarks are available in `scripts/benchmarks/` to compare the legacy MCP server (v2.2) against the new SDK v2 implementation (v2.3):

```bash
# Run performance benchmarks (requires Windows with Outlook running)
uv run --with pytest --with pywin32 python -m scripts.benchmarks.performance_benchmark
```

**Note:** Benchmarks require Windows with Outlook running and pywin32 installed. They cannot run in WSL2 or CI/CD environments without Outlook access.

See [scripts/benchmarks/README.md](scripts/benchmarks/README.md) for benchmark documentation and [scripts/benchmarks/EXPECTED_RESULTS.md](scripts/benchmarks/EXPECTED_RESULTS.md) for expected output format and success criteria.
