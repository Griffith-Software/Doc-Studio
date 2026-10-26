# Asks Excel to open a workbook in repair mode and prints Excel's repair log (what it had to remove/fix).
# Usage: powershell -File tests\excel-repair-log.ps1 <file.xlsx>
param([string]$File)
$full = (Resolve-Path $File).Path
$before = Get-Date
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $m = [Type]::Missing
  try {
    $wb = $xl.Workbooks.Open($full, $m, $true, $m, $m, $m, $m, $m, $m, $m, $m, $m, $m, $m, 1)
    "opened in repair mode: $($wb.Name)"
    $wb.Close($false)
  } catch { "repair open failed: $($_.Exception.Message)" }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $before } | ForEach-Object { "--- $($_.Name)"; Get-Content $_.FullName -Raw }
