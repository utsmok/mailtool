#!/usr/bin/env bash
#
# Outlook Bridge Wrapper for WSL2
# Calls the Windows batch script which uses uv for dependency management
#

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Convert WSL path to Windows path for the batch file
WINDOWS_BATCH=$(wslpath -w "$SCRIPT_DIR/outlook.bat")

# Execute the Windows batch file
cmd.exe /c "$WINDOWS_BATCH" "$@"
