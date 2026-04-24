# Copies the worker to a stable location and registers two HKCU context-menu entries.
# Run as your normal user (no admin needed).

$installDir = Join-Path $env:LOCALAPPDATA 'DocxToPdf'
New-Item -ItemType Directory -Path $installDir -Force | Out-Null

$source = Join-Path $PSScriptRoot 'Convert-DocxToPdf.ps1'
$dest   = Join-Path $installDir   'Convert-DocxToPdf.ps1'
Copy-Item -LiteralPath $source -Destination $dest -Force

# Launch hidden so no console flashes up
$cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden ' +
       "-File `"$dest`" -FolderPath `"%V`""

$menuName = 'ExportDocxToPdf'
$label    = 'Export DOCX to PDF'
$icon     = 'shell32.dll,1'   # generic document icon; change index if you prefer

# (a) right-click ON a folder  -> Directory\shell
# (b) right-click INSIDE a folder (empty space) -> Directory\Background\shell
$roots = @(
    "HKCU:\Software\Classes\Directory\shell\$menuName",
    "HKCU:\Software\Classes\Directory\Background\shell\$menuName"
)

foreach ($root in $roots) {
    New-Item -Path $root -Force | Out-Null
    Set-ItemProperty -Path $root -Name '(Default)' -Value $label
    Set-ItemProperty -Path $root -Name 'Icon'      -Value $icon
    New-Item -Path "$root\command" -Force | Out-Null
    Set-ItemProperty -Path "$root\command" -Name '(Default)' -Value $cmd
}

Write-Host "Installed. Right-click a folder -> 'Show more options' -> $label"
