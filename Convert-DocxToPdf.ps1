param(
    [Parameter(Mandatory = $true)]
    [string]$FolderPath
)

Add-Type -AssemblyName System.Windows.Forms

function Show-Msg($text, $title = 'DOCX to PDF', $icon = 'Information') {
    [System.Windows.Forms.MessageBox]::Show($text, $title, 'OK', $icon) | Out-Null
}

# Normalize path
try {
    $FolderPath = (Resolve-Path -LiteralPath $FolderPath -ErrorAction Stop).Path
} catch {
    Show-Msg "Folder not found: $FolderPath" 'Error' 'Error'; exit 1
}
if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
    Show-Msg "Not a folder: $FolderPath" 'Error' 'Error'; exit 1
}

# Collect .docx files, excluding Word's lock files (~$name.docx)
$docxFiles = Get-ChildItem -LiteralPath $FolderPath -Filter *.docx -File |
             Where-Object { -not $_.Name.StartsWith('~$') }

if ($docxFiles.Count -eq 0) {
    Show-Msg "No .docx files found in:`n$FolderPath"; exit 0
}

# Start Word
try {
    $word = New-Object -ComObject Word.Application
} catch {
    Show-Msg "Microsoft Word is not installed or its COM server is unavailable.`n$($_.Exception.Message)" 'Error' 'Error'
    exit 1
}
$word.Visible = $false
$word.DisplayAlerts = 0   # wdAlertsNone

$wdFormatPDF = 17
$converted = 0
$skipped   = 0
$failed    = @()

foreach ($file in $docxFiles) {
    $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, '.pdf')
    if (Test-Path -LiteralPath $pdfPath) { $skipped++; continue }  # don't overwrite

    $doc = $null
    try {
        $doc = $word.Documents.Open(
            $file.FullName,
            $false,   # ConfirmConversions
            $true     # ReadOnly
        )
        $doc.SaveAs2($pdfPath, $wdFormatPDF)
        $doc.Close($false)
        $converted++
    } catch {
        $failed += "$($file.Name): $($_.Exception.Message)"
        if ($doc) { try { $doc.Close($false) } catch {} }
    }
}

$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

$summary = "Converted: $converted`nSkipped (pdf already exists): $skipped`nTotal .docx: $($docxFiles.Count)"
if ($failed.Count -gt 0) { $summary += "`n`nFailures:`n" + ($failed -join "`n") }
Show-Msg $summary 'DOCX to PDF' ($(if ($failed) { 'Warning' } else { 'Information' }))
