@echo off
REM Double-clickable restart for the SolidWorks MCP server.
REM Checks the code on disk, kills the running server process(es), then waits
REM for the Claude desktop app to bring it back. SolidWorks is left alone.
REM
REM Pass through any switch, e.g.:  restart_mcp.cmd -List
REM                                 restart_mcp.cmd -Id 23516
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0restart_mcp.ps1" %*
echo.
pause
