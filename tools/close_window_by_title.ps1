# Ask a top-level window to close itself (WM_CLOSE) and wait until it is
# really gone. Called by tools/app_ctl.py (stop) before force-killing the
# desktop shell: WM_CLOSE takes the shell's own close path (native confirm
# while a conversion is running, geometry saved on closing - D-05/D-01),
# where Stop-Process -Force would cut a running conversion in half with no
# confirmation at all.
#
# IMPORTANT: keep this file pure ASCII (same rule as kill_by_cmdline.ps1).
# The window title is Chinese and arrives as a real argument (UTF-16 via
# CreateProcess), so it compares correctly regardless of codepage.
#
# Usage:
#   close_window_by_title.ps1 -Title '<window title>' [-TimeoutSec 15]
#
# Exit codes:
#   0  the window is gone (WM_CLOSE accepted, or tray mode hid it)
#   2  no window with that title exists (nothing to close; no output)
#   3  the window is still there after the timeout (the user dismissed the
#      confirm box, etc.) - the caller falls back to a force kill
param(
    [Parameter(Mandatory = $true)][string]$Title,
    [int]$TimeoutSec = 15
)

Add-Type -Namespace PaperShell -Name Win32Close -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
public static extern System.IntPtr FindWindow(string className, string windowName);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool PostMessage(System.IntPtr hWnd, uint msg, System.IntPtr wParam, System.IntPtr lParam);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool IsWindowVisible(System.IntPtr hWnd);
'@

# FindWindowW by exact title: the same matching rule the shell itself uses
# for its single-instance check and its reap backstop (desktop.py), so the
# two editions ("..." and "... dev suffix") never hit each other's windows.
$hwnd = [PaperShell.Win32Close]::FindWindow($null, $Title)
if ($hwnd -eq [System.IntPtr]::Zero) { exit 2 }

# Post, not send: the shell may answer with a modal confirm box (D-05);
# SendMessage would block this script until the user answers, Post leaves
# the deadline in the polling loop below.
[void][PaperShell.Win32Close]::PostMessage($hwnd, 0x0010, [System.IntPtr]::Zero, [System.IntPtr]::Zero)

$deadline = (Get-Date).AddSeconds($TimeoutSec)
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 300
    $hwnd = [PaperShell.Win32Close]::FindWindow($null, $Title)
    if ($hwnd -eq [System.IntPtr]::Zero) { exit 0 }
    if (-not [PaperShell.Win32Close]::IsWindowVisible($hwnd)) { exit 0 }
}
exit 3
