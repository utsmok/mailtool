#!/usr/bin/env pwsh
# Ralph Wiggum - Long-running AI agent loop (Claude Code)
# Usage: .\ralph.ps1 [-MaxIterations <int>]
#
# Requirements:
#   - Claude Code CLI (claude) installed and configured
#   - ANTHROPIC_API_KEY environment variable set
#   - prd.json and prompt.md in the same directory
#   - PowerShell 7+ (recommended) or Windows PowerShell 5.1+

param(
    [int]$MaxIterations = 10
)

# strict mode for safety
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Script directory paths
$ScriptDir = $PSScriptRoot
$PrdFile = Join-Path $ScriptDir "prd.json"
$ProgressFile = Join-Path $ScriptDir "progress.txt"
$ArchiveDir = Join-Path $ScriptDir "archive"
$LastBranchFile = Join-Path $ScriptDir ".last-branch"
$StopRequestFile = Join-Path $ScriptDir ".stop-requested"

# Clean up any stale stop request file
if (Test-Path $StopRequestFile) {
    Remove-Item $StopRequestFile -Force
}

# Clean up any leftover listener processes from previous runs
Get-Process | Where-Object {
    $_.ProcessName -like "*listen*" -and
    $_.CommandLine -like "*ralph*"
} | Stop-Process -Force -ErrorAction SilentlyContinue

# Cleanup function for exit traps
$cleanup = {
    # Kill the listener job if it's running
    if ($null -ne $script:ListenerJob) {
        Stop-Job -Job $script:ListenerJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:ListenerJob -ErrorAction SilentlyContinue
    }
    # Clean up stop request file
    if (Test-Path $StopRequestFile) {
        Remove-Item $StopRequestFile -Force
    }
}

# Register cleanup on exit
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action $cleanup

# Function to listen for stop signal in background
function Start-StopListener {
    $listenerScript = {
        $host.ui.RawUI.FlushInputBuffer()
        while ($true) {
            if ($host.ui.RawUI.KeyAvailable) {
                $key = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                if ($key.Character -eq 's' -or $key.Character -eq 'S') {
                    "" | Out-File -FilePath $using:StopRequestFile -Encoding ASCII
                    Write-Host ""
                    Write-Host ">>> Stop requested! This iteration will complete, then exit."
                    break
                }
            }
            Start-Sleep -Milliseconds 500
        }
    }.GetNewClosure()

    Start-Job -ScriptBlock $listenerScript -Name "RalphStopListener"
}

# Archive previous run if branch changed
if ((Test-Path $PrdFile) -and (Test-Path $LastBranchFile)) {
    $prdJson = Get-Content $PrdFile -Raw | ConvertFrom-Json
    $currentBranch = if ($prdJson.branchName) { $prdJson.branchName } else { "" }
    $lastBranch = if (Test-Path $LastBranchFile) {
        Get-Content $LastBranchFile -Raw
    } else {
        ""
    }

    if ($currentBranch -and $lastBranch -and $currentBranch -ne $lastBranch) {
        # Archive the previous run
        $date = Get-Date -Format "yyyy-MM-dd"
        # Strip "ralph/" prefix from branch name for folder
        $folderName = $lastBranch -replace "^ralph/", ""
        $archiveFolder = Join-Path $ArchiveDir "$date-$folderName"

        Write-Host "Archiving previous run: $lastBranch"
        New-Item -ItemType Directory -Path $archiveFolder -Force | Out-Null
        if (Test-Path $PrdFile) {
            Copy-Item $PrdFile $archiveFolder -Force
        }
        if (Test-Path $ProgressFile) {
            Copy-Item $ProgressFile $archiveFolder -Force
        }
        Write-Host "   Archived to: $archiveFolder"

        # Reset progress file for new run
        $progressHeader = @"
# Ralph Progress Log
Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
---
"@
        $progressHeader | Out-File -FilePath $ProgressFile -Encoding ASCII
    }
}

# Track current branch
if (Test-Path $PrdFile) {
    $prdJson = Get-Content $PrdFile -Raw | ConvertFrom-Json
    $currentBranch = if ($prdJson.branchName) { $prdJson.branchName } else { "" }
    if ($currentBranch) {
        $currentBranch | Out-File -FilePath $LastBranchFile -Encoding ASCII
    }
}

# Initialize progress file if it doesn't exist
if (-not (Test-Path $ProgressFile)) {
    $progressHeader = @"
# Ralph Progress Log
Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
---
"@
    $progressHeader | Out-File -FilePath $ProgressFile -Encoding ASCII
}

Write-Host "Starting Ralph - Max iterations: $MaxIterations"
Write-Host "Press 's' at any time to stop after current iteration (will ask for confirmation)"

# Main iteration loop
for ($i = 1; $i -le $MaxIterations; $i++) {
    # Start stop listener in background
    $script:ListenerJob = Start-StopListener

    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════"
    Write-Host "  Ralph Iteration $i of $MaxIterations"
    Write-Host "═══════════════════════════════════════════════════════"

    # Read prompt
    $promptText = Get-Content (Join-Path $ScriptDir "prompt.md") -Raw

    # Run Claude Code with the ralph prompt
    $output = & claude -p $promptText --allowedTools "Read,Write,Edit,Bash,Glob,Grep,Task,AskUserQuestion" 2>&1
    $lastExitCode = $LASTEXITCODE

    # Stop the listener
    Stop-Job -Job $script:ListenerJob -ErrorAction SilentlyContinue
    Remove-Job -Job $script:ListenerJob -ErrorAction SilentlyContinue
    $script:ListenerJob = $null

    # Check for completion signal
    if ($output -match "<promise>COMPLETE</promise>") {
        Write-Host ""
        Write-Host "Ralph completed all tasks!"
        Write-Host "Completed at iteration $i of $MaxIterations"
        if (Test-Path $StopRequestFile) {
            Remove-Item $StopRequestFile -Force
        }
        exit 0
    }

    # Check if user requested stop
    if (Test-Path $StopRequestFile) {
        Write-Host ""
        Write-Host "═══════════════════════════════════════════════════════"
        Write-Host "  STOP REQUESTED"
        Write-Host "═══════════════════════════════════════════════════════"
        $confirm = Read-Host "Confirm stop and exit? (y/n)"
        if ($confirm -eq "y" -or $confirm -eq "Y") {
            Write-Host ""
            Write-Host "Stopping gracefully after iteration $i..."
            Write-Host "Progress saved to: $ProgressFile"
            Write-Host "You can resume later by running this script again."
            Remove-Item $StopRequestFile -Force
            exit 0
        } else {
            Write-Host "Continuing..."
            Remove-Item $StopRequestFile -Force
        }
    }

    Write-Host "Iteration $i complete. Continuing..."
    Start-Sleep -Seconds 2
}

Write-Host ""
Write-Host "Ralph reached max iterations ($MaxIterations) without completing all tasks."
Write-Host "Check $ProgressFile for status."
exit 1
