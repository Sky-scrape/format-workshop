# Report python/pythonw processes running app.py or keepalive.py, with parent name.
# Kept pure ASCII: Windows PowerShell 5.1 reads .ps1 without a BOM as ANSI,
# so non-ASCII content can break parsing.

$procs = @(Get-CimInstance Win32_Process | Where-Object {
    $_.Name -match '^(python|pythonw)\.exe$' -and $_.CommandLine -match 'app\.py|keepalive'
})

if ($procs.Count -eq 0) {
    Write-Output "     (none)"
    exit 0
}

foreach ($p in $procs) {
    $par = Get-CimInstance Win32_Process -Filter "ProcessId=$($p.ParentProcessId)" -ErrorAction SilentlyContinue
    $parName = if ($par) { $par.Name } else { "?" }
    $role = if ($p.CommandLine -match 'keepalive') { "keepalive" } else { "app" }
    $flag = if ($parName -eq "explorer.exe") { "OK 可长期驻留" } else { "注意: 父进程不是 explorer" }
    Write-Output ("     pid=$($p.ProcessId)  $role  parent=$parName  [$flag]")
}
exit 0
