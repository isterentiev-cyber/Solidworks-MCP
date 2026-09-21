<#
.SYNOPSIS
    Kill SolidWorks MCP server process(es) so the client respawns them.

.DESCRIPTION
    The server is a stdio child process of the Claude desktop app, and the app
    starts ONE PER SESSION -- so several are normally running at once, all
    talking to the same SolidWorks over COM. Tool SCHEMAS are fixed when a
    process starts, so a brand-new tool only becomes visible after its process
    restarts.

    In most cases you do NOT need this: reload_api sends a tools/list_changed
    notification, which picks up new tools live. Reach for this script when
    that is not enough -- notably after editing server.py itself, which
    reload_api deliberately does not reload.

    SolidWorks is NOT touched: the COM connection is re-established by the next
    connect_solidworks, and open documents stay open.

.PARAMETER Id
    Kill only this PID. Without it, every matching process is killed -- other
    Claude sessions will respawn theirs on their next tool call, but anything
    mid-operation there is interrupted.

.PARAMETER List
    Only list the processes, kill nothing.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts\restart_mcp.ps1 -List
    powershell -ExecutionPolicy Bypass -File scripts\restart_mcp.ps1 -Id 23516
    powershell -ExecutionPolicy Bypass -File scripts\restart_mcp.ps1
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [int]$Id,
    [switch]$List
)

$ErrorActionPreference = "Stop"
$marker = "solidworks_mcp_server.py"

# Match on the command line, not the image name: plenty of unrelated python.exe
# processes are usually running (other MCP servers, editors, the venv itself).
$procs = @(Get-CimInstance Win32_Process -Filter "Name like '%python%'" |
    Where-Object { $_.CommandLine -and $_.CommandLine -like "*$marker*" })

if ($procs.Count -eq 0) {
    Write-Host "SolidWorks MCP server is not running." -ForegroundColor Yellow
    Write-Host "Nothing to do - the client starts it on the next tool call."
    exit 0
}

Write-Host ("Found {0} SolidWorks MCP server process(es):" -f $procs.Count) -ForegroundColor Cyan
foreach ($p in $procs | Sort-Object CreationDate) {
    $started = $p.CreationDate
    $age = [int]((Get-Date) - $started).TotalMinutes
    Write-Host ("  PID {0,-7} started {1:HH:mm:ss} ({2} min ago)" -f $p.ProcessId, $started, $age)
}
Write-Host ""

if ($List) {
    Write-Host "-List given, nothing killed." -ForegroundColor Yellow
    Write-Host "Kill one with:  restart_mcp.ps1 -Id <PID>"
    exit 0
}

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
        Write-Host ("Killing all {0}. Other Claude sessions respawn theirs on their" -f $procs.Count) -ForegroundColor Yellow
        Write-Host "next tool call, but anything running there right now is interrupted."
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
Write-Host "Done. If the client does not respawn the server by itself," -ForegroundColor Yellow
Write-Host "reconnect it from the app's MCP servers list (or restart the app)."
Write-Host "SolidWorks itself was not touched - your documents are still open."
