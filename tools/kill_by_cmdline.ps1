# Kill processes whose command line matches a keyword AND belongs to this
# project. Called by tools/app_ctl.py (stop). Kept in a separate file to avoid
# the classic for /f + PowerShell quoting mess in batch scripts.
#
# IMPORTANT: keep this file pure ASCII. Windows PowerShell 5.1 reads .ps1
# files without a BOM as ANSI, so non-ASCII comments would be garbled and
# can break parsing. $Root arrives as a real argument (UTF-16 via
# CreateProcess), so a non-ASCII project path still compares correctly.
#
# Usage:
#   kill_by_cmdline.ps1 -Pattern 'app\.py' -Root '<project root>' [-Quiet] [-DryRun]
#     -Root    mandatory path scope: only processes whose command line contains
#              this path are considered (F48: a bare substring match used to
#              force-kill unrelated projects' python processes).
#     -DryRun  report matching pids ("match <pid>") without killing anything.
param(
    [Parameter(Mandatory = $true)][string]$Pattern,
    [Parameter(Mandatory = $true)][string]$Root,
    [switch]$Quiet,
    [switch]$DryRun
)

$rootPath = $Root.TrimEnd('\')
$procs = @(Get-CimInstance Win32_Process |
    Where-Object {
        $_.Name -match '^(python|pythonw)\.exe$' -and $_.CommandLine -and
        $_.CommandLine -match $Pattern -and
        $_.CommandLine.IndexOf($rootPath, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
    })

# No match: exit 1 with no output. The caller (app_ctl.stop) must report
# "not found" honestly instead of treating empty output as success.
if ($procs.Count -eq 0) { exit 1 }

foreach ($p in $procs) {
    if ($DryRun) { Write-Output ("match " + $p.ProcessId); continue }
    try {
        Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
        if (-not $Quiet) { Write-Output ("killed " + $p.ProcessId) }
    } catch {
        if (-not $Quiet) { Write-Output ("failed " + $p.ProcessId + ": " + $_.Exception.Message) }
    }
}
exit 0
