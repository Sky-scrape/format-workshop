# Emit machine-readable process info, one line per process: pid|role|parentName
# Pure ASCII on purpose:
#   - Windows PowerShell 5.1 reads .ps1 without a BOM as ANSI, so non-ASCII
#     content here can break parsing.
#   - Chinese wording belongs to the caller (tools/app_ctl.py). Keeping this
#     file data-only avoids encoding mismatches between PowerShell stdout and
#     Python.

param(
    [Parameter(Mandatory = $true)][string]$Root
)

# -Root scopes the report to this project, same rule as kill_by_cmdline.ps1
# (F48): another project's python processes carrying the same keywords
# (app.py / keepalive / desktop.py) used to be reported as ours.
$rootPath = $Root.TrimEnd('\')
$procs = @(Get-CimInstance Win32_Process | Where-Object {
    $_.Name -match '^(python|pythonw)\.exe$' -and $_.CommandLine -and
    $_.CommandLine -match 'app\.py|keepalive|desktop\.py' -and
    $_.CommandLine.IndexOf($rootPath, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
})

foreach ($p in $procs) {
    $parName = "-"
    $par = Get-CimInstance Win32_Process -Filter "ProcessId=$($p.ParentProcessId)" -ErrorAction SilentlyContinue
    if ($par) { $parName = $par.Name }

    $role = "app"
    if ($p.CommandLine -match 'keepalive') { $role = "keepalive" }
    if ($p.CommandLine -match 'desktop\.py') { $role = "shell" }

    Write-Output ("{0}|{1}|{2}" -f $p.ProcessId, $role, $parName)
}
exit 0
