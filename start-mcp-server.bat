@echo off
REM Get the plugin root directory
set PLUGIN_ROOT=%~dp0
REM Change to plugin root
cd /d "%PLUGIN_ROOT%"
REM Add src to PYTHONPATH so Python can find mailtool module
set "PYTHONPATH=%PLUGIN_ROOT%src;%PYTHONPATH%"
REM Use uv to run Python with pywin32, passing args after --
uv run --with pywin32 -- python -m mailtool.mcp.server
