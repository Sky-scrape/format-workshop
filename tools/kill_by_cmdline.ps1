# Kill processes whose command line matches a keyword.
# Called by 停止.bat. Kept in a separate file to avoid the classic
# for /f + PowerShell quoting mess in batch scripts.
#
# IMPORTANT: keep this file pure ASCII. Windows PowerShell 5.1 reads .ps1
# files without a BOM as ANSI, so non-ASCII comments would be garbled and
# can break parsing.
param(
    [Parameter(Mandatory = $true)][string]$Pattern,
    [switch]$Quiet
)

$procs = @(Get-CimInstance Win32_Process |
    Where-Object { $_.Name -match '^(python|pythonw)\.exe$' -and $_.CommandLine -match $Pattern })

if ($procs.Count -eq 0) { exit 1 }

foreach ($p in $procs) {
    try {
        Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
        if (-not $Quiet) { Write-Output ("killed " + $p.ProcessId) }
    } catch {
        if (-not $Quiet) { Write-Output ("failed " + $p.ProcessId + ": " + $_.Exception.Message) }
    }
}
exit 0
