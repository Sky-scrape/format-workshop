# Convert a legacy .doc to .docx with whichever Office suite this machine has.
# Called by app.py (doc_to_docx). Kept in PowerShell instead of Python so the
# portable bundle needs no pywin32, no COM apartment handling, and no reference
# counting inside the Flask process -- same arrangement as kill_by_cmdline.ps1.
#
# IMPORTANT: keep this file pure ASCII. Windows PowerShell 5.1 reads .ps1
# without a BOM as ANSI; non-ASCII comments come out garbled and can break
# parsing. Paths arrive as real arguments (UTF-16 via CreateProcess), so a
# non-ASCII working directory still converts correctly.
#
# Usage:
#   doc2docx.ps1 -In <in.doc> -Out <out.docx> -ProgId Word.Application
#
# Exit codes -- the caller uses them to decide what to try next:
#   0  converted, output file exists
#   2  suite present but this file failed (corrupt, password, weird macro)
#   3  suite not installed / COM object not registered
param(
    [Parameter(Mandatory=$true)][string]$In,
    [Parameter(Mandatory=$true)][string]$Out,
    [string]$ProgId = "Word.Application"
)

$ErrorActionPreference = "Stop"

# Instantiate first: a missing suite throws here and we report 3, which is a
# different fact from "the file is broken" and must not be conflated.
$app = $null
$ownInstance = $false
try {
    $app = New-Object -ComObject $ProgId
    # CoCreateInstance can hand back the instance the author already has open.
    # Quitting that one would close their unsaved manuscript, so only quit when
    # we can see the instance is empty (i.e. ours).
    try { $ownInstance = ($app.Documents.Count -eq 0) } catch { $ownInstance = $true }
} catch {
    exit 3
}

$doc = $null
try {
    # Only touch visibility on an instance we created. Hiding the author's
    # running Word would look like their manuscript just vanished, which is a
    # far worse experience than a document flashing open and closed.
    if ($ownInstance) {
        try { $app.Visible = $false } catch { }
    }
    try { $app.DisplayAlerts = 0 } catch { }

    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
    # ReadOnly is deliberate: the author's original must come back untouched.
    $doc = $app.Documents.Open($In, $false, $true, $false)

    # 16 = wdFormatXMLDocument (.docx). SaveAs2 exists on Word 2010+ and on
    # WPS; older builds only have SaveAs, so fall back once.
    try { $doc.SaveAs2($Out, 16) } catch { $doc.SaveAs($Out, 16) }

    if (-not (Test-Path -LiteralPath $Out)) { exit 2 }
    exit 0
}
catch {
    exit 2
}
finally {
    # Close()/Quit() take VARIANT by-ref arguments, and `Close(0)` raises
    # "parameter 1 should be PSReference" under PowerShell 5.1 -- swallowed by
    # the catch below, which used to leave WINWORD.EXE running for every
    # conversion. Mark the document clean and call the no-argument overloads.
    if ($doc) {
        try { $doc.Saved = $true } catch { }
        try { $doc.Close() } catch { }
    }
    if ($app) {
        if ($ownInstance) { try { $app.Quit() } catch { } }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
    }
}
