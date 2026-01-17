@echo off
REM Outlook Bridge - Windows entry point using uv
REM This runs on Windows and uses uv to manage dependencies

setlocal
set SCRIPT_DIR=%~dp0
set PYTHON_SCRIPT=%SCRIPT_DIR%src\mailtool_outlook_bridge.py

cd /d "%SCRIPT_DIR%"

REM Use uv to run with pywin32 dependency
uv run --with pywin32 python "%PYTHON_SCRIPT%" %*
