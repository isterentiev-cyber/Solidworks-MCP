<#
.SYNOPSIS
    Restart the SolidWorks MCP server: kill it, check the code, wait for the
    client to bring it back.

.DESCRIPTION
    The server is a stdio CHILD PROCESS of the Claude desktop app, so nothing
    here can start it directly -- a python.exe launched by hand would just sit
    on stdin with no client attached and only confuse the process list. What a
    restart actually means is: kill it, and let the client respawn it.

    Measured behaviour (2026-09-22): the client respawns the server LAZILY, on
    the first tool call, not on its own. So this script kills, verifies the
    code on disk is servable, then waits -- and if nothing comes back, tells
    you to make one tool call in Claude.

    The app starts one server PER SESSION, so several are normally running at
    once, all talking to the same SolidWorks over COM.

    SolidWorks is NOT touched: the COM connection is re-established by the next
    connect_solidworks, and open documents stay open.

    NOTE: a restart does NOT make a newly added tool appear in a chat that is
    already open -- Claude Desktop fixes the tool list when the conversation
    starts. New tools need a NEW chat. See NOTES.md.

.PARAMETER Id
    Kill only this PID. Without it, every matching process is killed -- other
    Claude sessions respawn theirs on their next tool call, but anything
    mid-operation there is interrupted.

.PARAMETER List
    Only list the processes, kill nothing.

.PARAMETER SkipCheck
    Do not run scripts/selftest.py before killing.

.PARAMETER WaitSeconds
    How long to wait for the client to respawn the server (default 20).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts\restart_mcp.ps1
    powershell -ExecutionPolicy Bypass -File scripts\restart_mcp.ps1 -List
    powershell -ExecutionPolicy Bypass -File scripts\restart_mcp.ps1 -Id 23516
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [int]$Id,
    [switch]$List,
    [switch]$SkipCheck,
    [int]$WaitSeconds = 20
)

$ErrorActionPreference = "Stop"
$marker = "solidworks_mcp_server.py"
$repo = Split-Path -Parent $PSScriptRoot
$venvPy = Join-Path $repo ".venv\Scripts\python.exe"

function Get-McpServers {
    # Match on the command line, not the image name: plenty of unrelated
    # python.exe processes are usually running (other MCP servers, editors).
    @(Get-CimInstance Win32_Process -Filter "Name like '%python%'" |
        Where-Object { $_.CommandLine -and $_.CommandLine -like "*$marker*" })
}

function Show-Servers($procs) {
    foreach ($p in $procs | Sort-Object CreationDate) {
        $age = [int]((Get-Date) - $p.CreationDate).TotalMinutes
        Write-Host ("  PID {0,-7} started {1:HH:mm:ss} ({2} min ago)" -f `
            $p.ProcessId, $p.CreationDate, $age)
    }
}

# ---------------------------------------------------------------- inventory
$procs = Get-McpServers
if ($procs.Count -eq 0) {
    Write-Host "No SolidWorks MCP server is running." -ForegroundColor Yellow
}
else {
    Write-Host ("Running server process(es): {0}" -f $procs.Count) -ForegroundColor Cyan
    Show-Servers $procs
}
Write-Host ""

if ($List) {
    Write-Host "-List given, nothing killed." -ForegroundColor Yellow
    Write-Host "Kill one with:  restart_mcp.ps1 -Id <PID>"
    exit 0
}

# ------------------------------------------------------------- check first
# Killing a server whose code does not import just swaps a working server for
# a dead one, and the client reports that as "no such tool" with no traceback
# anywhere. Check before killing, not after.
if (-not $SkipCheck) {
    if (-not (Test-Path $venvPy)) {
        Write-Host "venv python not found at $venvPy -- skipping the code check." -ForegroundColor Yellow
    }
    else {
        Write-Host "Checking the code on disk..." -ForegroundColor Cyan
        & $venvPy (Join-Path $PSScriptRoot "selftest.py")
        if ($LASTEXITCODE -ne 0) {
            Write-Host ""
            Write-Host "Code on disk is NOT servable -- nothing killed." -ForegroundColor Red
            Write-Host "Fix the problem above, or re-run with -SkipCheck to kill anyway."
            exit 1
        }
        Write-Host ""
    }
}

# -------------------------------------------------------------------- kill
if ($procs.Count -gt 0) {
    if ($Id) {
        $targets = @($procs | Where-Object { $_.ProcessId -eq $Id })
        if ($targets.Count -eq 0) {
            Write-Host ("PID {0} is not a SolidWorks MCP server." -f $Id) -ForegroundColor Red
            exit 1
        }
    }
    else {
        $targets = $procs
        if ($procs.Count -gt 1) {
            Write-Host ("Killing all {0}. Other Claude sessions respawn theirs on" -f $procs.Count) -ForegroundColor Yellow
            Write-Host "their next tool call, but anything running there now is interrupted."
            Write-Host "Use -Id <PID> to kill just one."
            Write-Host ""
        }
    }

    foreach ($p in $targets) {
        if ($PSCmdlet.ShouldProcess("PID $($p.ProcessId)", "Stop-Process")) {
            try {
                Stop-Process -Id $p.ProcessId -Force
                Write-Host ("  killed PID {0}" -f $p.ProcessId) -ForegroundColor Green
            }
            catch {
                Write-Host ("  could not kill PID {0}: {1}" -f $p.ProcessId, $_.Exception.Message) `
                    -ForegroundColor Red
            }
        }
    }
    Write-Host ""
}

if ($WhatIfPreference) { exit 0 }

# ------------------------------------------------------------ wait for it
$killed = @($targets | ForEach-Object { $_.ProcessId })
Write-Host ("Waiting up to {0}s for the client to bring it back..." -f $WaitSeconds) -ForegroundColor Cyan

$deadline = (Get-Date).AddSeconds($WaitSeconds)
$fresh = @()
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 700
    $fresh = @(Get-McpServers | Where-Object { $killed -notcontains $_.ProcessId })
    if ($fresh.Count -gt 0) { break }
}

Write-Host ""
if ($fresh.Count -gt 0) {
    Write-Host "Back up:" -ForegroundColor Green
    Show-Servers $fresh
}
else {
    Write-Host "Not back yet -- that is normal." -ForegroundColor Yellow
    Write-Host "The client starts the server LAZILY, on the first tool call."
    Write-Host "Ask Claude to run any SolidWorks tool and it will respawn."
}

Write-Host ""
Write-Host "Reminder: a restart does not reveal NEWLY ADDED tools in a chat" -ForegroundColor DarkGray
Write-Host "that is already open -- the tool list is fixed when the chat starts." -ForegroundColor DarkGray
Write-Host "New tools need a new chat. SolidWorks itself was not touched." -ForegroundColor DarkGray
