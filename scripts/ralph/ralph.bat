@echo off
REM Ralph Wiggum - Windows batch wrapper
REM Usage: ralph.bat [max_iterations]

setlocal

REM Get script directory
set SCRIPT_DIR=%~dp0

REM Default to 10 iterations if not specified
set MAX_ITERATIONS=%1
if "%MAX_ITERATIONS%"=="" set MAX_ITERATIONS=10

REM Run PowerShell script
powershell.exe -ExecutionPolicy Bypass -File "%SCRIPT_DIR%ralph.ps1" -MaxIterations %MAX_ITERATIONS%
