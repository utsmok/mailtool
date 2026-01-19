#!/bin/bash
# Ralph Wiggum - Long-running AI agent loop (Claude Code)
# Usage: ./ralph.sh [max_iterations]
#
# Requirements:
#   - Claude Code CLI (claude) installed and configured
#   - ANTHROPIC_API_KEY environment variable set
#   - prd.json and prompt.md in the same directory

set -e

MAX_ITERATIONS=${1:-10}
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PRD_FILE="$SCRIPT_DIR/prd.json"
PROGRESS_FILE="$SCRIPT_DIR/progress.txt"
ARCHIVE_DIR="$SCRIPT_DIR/archive"
LAST_BRANCH_FILE="$SCRIPT_DIR/.last-branch"
STOP_REQUEST_FILE="$SCRIPT_DIR/.stop-requested"

# Clean up any stale stop request file
rm -f "$STOP_REQUEST_FILE"

# Clean up any leftover listener processes from previous runs
pkill -f "listen_for_stop" 2>/dev/null || true

# Clean up any leftover stop request files
find "$SCRIPT_DIR" -name ".stop-requested" -type f -delete 2>/dev/null || true

# Cleanup function for exit traps
cleanup() {
  # Kill the listener process if it's running
  if [ -n "$LISTENER_PID" ]; then
    kill $LISTENER_PID 2>/dev/null || true
  fi
  # Clean up stop request file
  rm -f "$STOP_REQUEST_FILE"
}

# Set trap for cleanup on exit
trap cleanup EXIT INT TERM

# Function to listen for stop signal in background
listen_for_stop() {
  while true; do
    # Read a single character with timeout (non-blocking)
    read -s -n 1 -t 0.5 key 2>/dev/null || true
    if [ "$key" = "s" ] || [ "$key" = "S" ]; then
      touch "$STOP_REQUEST_FILE"
      echo ""
      echo ">>> Stop requested! This iteration will complete, then exit."
      break
    fi
  done
}

# Archive previous run if branch changed
if [ -f "$PRD_FILE" ] && [ -f "$LAST_BRANCH_FILE" ]; then
  CURRENT_BRANCH=$(jq -r '.branchName // empty' "$PRD_FILE" 2>/dev/null || echo "")
  LAST_BRANCH=$(cat "$LAST_BRANCH_FILE" 2>/dev/null || echo "")

  if [ -n "$CURRENT_BRANCH" ] && [ -n "$LAST_BRANCH" ] && [ "$CURRENT_BRANCH" != "$LAST_BRANCH" ]; then
    # Archive the previous run
    DATE=$(date +%Y-%m-%d)
    # Strip "ralph/" prefix from branch name for folder
    FOLDER_NAME=$(echo "$LAST_BRANCH" | sed 's|^ralph/||')
    ARCHIVE_FOLDER="$ARCHIVE_DIR/$DATE-$FOLDER_NAME"

    echo "Archiving previous run: $LAST_BRANCH"
    mkdir -p "$ARCHIVE_FOLDER"
    [ -f "$PRD_FILE" ] && cp "$PRD_FILE" "$ARCHIVE_FOLDER/"
    [ -f "$PROGRESS_FILE" ] && cp "$PROGRESS_FILE" "$ARCHIVE_FOLDER/"
    echo "   Archived to: $ARCHIVE_FOLDER"

    # Reset progress file for new run
    echo "# Ralph Progress Log" > "$PROGRESS_FILE"
    echo "Started: $(date)" >> "$PROGRESS_FILE"
    echo "---" >> "$PROGRESS_FILE"
  fi
fi

# Track current branch
if [ -f "$PRD_FILE" ]; then
  CURRENT_BRANCH=$(jq -r '.branchName // empty' "$PRD_FILE" 2>/dev/null || echo "")
  if [ -n "$CURRENT_BRANCH" ]; then
    echo "$CURRENT_BRANCH" > "$LAST_BRANCH_FILE"
  fi
fi

# Initialize progress file if it doesn't exist
if [ ! -f "$PROGRESS_FILE" ]; then
  echo "# Ralph Progress Log" > "$PROGRESS_FILE"
  echo "Started: $(date)" >> "$PROGRESS_FILE"
  echo "---" >> "$PROGRESS_FILE"
fi

echo "Starting Ralph - Max iterations: $MAX_ITERATIONS"
echo "Press 's' at any time to stop after current iteration (will ask for confirmation)"

for i in $(seq 1 $MAX_ITERATIONS); do
  # Start stop listener in background
  listen_for_stop &
  LISTENER_PID=$!

  echo ""
  echo "═══════════════════════════════════════════════════════"
  echo "  Ralph Iteration $i of $MAX_ITERATIONS"
  echo "═══════════════════════════════════════════════════════"

  # Run Claude Code with the ralph prompt
  OUTPUT=$(claude -p "$(cat "$SCRIPT_DIR/prompt.md")" --allowedTools "Read,Write,Edit,Bash,Glob,Grep,Task,AskUserQuestion" 2>&1 | tee /dev/stderr) || true

  # Stop the listener
  kill $LISTENER_PID 2>/dev/null || true
  wait $LISTENER_PID 2>/dev/null || true

  # Check for completion signal
  if echo "$OUTPUT" | grep -q "<promise>COMPLETE</promise>"; then
    echo ""
    echo "Ralph completed all tasks!"
    echo "Completed at iteration $i of $MAX_ITERATIONS"
    rm -f "$STOP_REQUEST_FILE"
    exit 0
  fi

  # Check if user requested stop
  if [ -f "$STOP_REQUEST_FILE" ]; then
    echo ""
    echo "═══════════════════════════════════════════════════════"
    echo "  STOP REQUESTED"
    echo "═══════════════════════════════════════════════════════"
    echo -n "Confirm stop and exit? (y/n): "
    read -n 1 confirm
    echo ""
    if [ "$confirm" = "y" ] || [ "$confirm" = "Y" ]; then
      echo ""
      echo "Stopping gracefully after iteration $i..."
      echo "Progress saved to: $PROGRESS_FILE"
      echo "You can resume later by running this script again."
      rm -f "$STOP_REQUEST_FILE"
      exit 0
    else
      echo "Continuing..."
      rm -f "$STOP_REQUEST_FILE"
    fi
  fi

  echo "Iteration $i complete. Continuing..."
  sleep 2
done

echo ""
echo "Ralph reached max iterations ($MAX_ITERATIONS) without completing all tasks."
echo "Check $PROGRESS_FILE for status."
exit 1
