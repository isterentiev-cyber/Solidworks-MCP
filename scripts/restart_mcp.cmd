@echo off
REM Double-clickable wrapper for restart_mcp.ps1 -- kills the SolidWorks MCP
REM server process so the Claude desktop app respawns it with fresh tool
REM schemas. SolidWorks itself is left alone.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0restart_mcp.ps1"
echo.
pause
