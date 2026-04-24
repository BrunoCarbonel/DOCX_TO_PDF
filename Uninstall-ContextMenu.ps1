Remove-Item 'HKCU:\Software\Classes\Directory\shell\ExportDocxToPdf'            -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item 'HKCU:\Software\Classes\Directory\Background\shell\ExportDocxToPdf' -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $env:LOCALAPPDATA 'DocxToPdf')                           -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Uninstalled."
