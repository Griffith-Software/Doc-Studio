# Opens files in the installed Microsoft Office (read-only, no repair) to prove Doc Studio's output is valid.
# Usage: powershell -File tests\office-validate.ps1 <file> [<file> ...]
# For workbooks it also prints recalculated values of the first sheet's used range (first 12 rows x 8 columns).
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Files)
$ErrorActionPreference = 'Stop'
$results = @()
$word = $null; $excel = $null; $ppt = $null
try {
  foreach ($f in $Files) {
    $full = (Resolve-Path $f).Path
    $ext = [IO.Path]::GetExtension($full).ToLower()
    try {
      if ($ext -in '.docx', '.docm', '.doc', '.rtf', '.odt') {
        if (-not $word) { $word = New-Object -ComObject Word.Application; $word.Visible = $false; $word.DisplayAlerts = 0 }
        # OpenAndRepair = $false: a damaged file throws instead of being silently fixed.
        $d = $word.Documents.Open($full, $false, $true, $false, '', '', $true, '', '', 0, 0, $false, $false)
        $hf = @()
        foreach ($sec in $d.Sections) { $hf += ('H' + $sec.Index + '=' + ($sec.Headers.Item(1).Range.Text -replace '[

]+', ' ').Trim()); $hf += ('F' + $sec.Index + '=' + ($sec.Footers.Item(1).Range.Text -replace '[

]+', ' ').Trim()) }
        $notes = @(); foreach ($fn in $d.Footnotes) { $notes += ($fn.Range.Text -replace '[

]+', ' ').Trim() }
        $results += [pscustomobject]@{ File = (Split-Path $full -Leaf); OK = $true; Info = "paragraphs=$($d.Paragraphs.Count) tables=$($d.Tables.Count) shapes=$($d.InlineShapes.Count) pages=$($d.ComputeStatistics(2)) lists=$($d.Lists.Count)`n    " + ($hf -join ' | ') + "`n    notes: " + ($notes -join ' | ') + "`n    first: " + (($d.Paragraphs.Item(1).Range.Text + ' / ' + $(if ($d.Paragraphs.Count -gt 1) { $d.Paragraphs.Item(2).Range.Text } else { '' })) -replace '[

]+', ' ') }
        $d.Close(0)
      } elseif ($ext -in '.xlsx', '.xlsm', '.xls', '.xlsb', '.ods') {
        if (-not $excel) { $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false }
        # Excel repairs damaged files silently when alerts are off; a repaired workbook opens with "[Repaired]"
        # in its title, so check for that.
        $wb = $excel.Workbooks.Open($full, 0, $true)
        if ($excel.ActiveWindow -and ($excel.ActiveWindow.Caption -match 'Repaired')) { throw "Excel had to repair the file" }
        $excel.CalculateFull()
        $ws = $wb.Worksheets.Item(1)
        $rows = @()
        for ($r = 1; $r -le 12; $r++) {
          $vals = @()
          for ($c = 1; $c -le 8; $c++) { $vals += [string]$ws.Cells.Item($r, $c).Text }
          $rows += ($vals -join ' | ')
        }
        $names = @(); foreach ($s in $wb.Worksheets) { $names += $s.Name }
        $hiddenRows = @(); for ($r = 1; $r -le 20; $r++) { if ($ws.Rows.Item($r).Hidden) { $hiddenRows += $r } }
        $extra = " filter=$($ws.AutoFilterMode) hiddenRows=$($hiddenRows -join ',') tables=$($ws.ListObjects.Count) cf=$($ws.Cells.FormatConditions.Count)"
        $names[-1] = $names[-1] + $extra
        $results += [pscustomobject]@{ File = (Split-Path $full -Leaf); OK = $true; Info = "sheets=$($names -join ',') charts=$($ws.ChartObjects().Count)`n    " + ($rows -join "`n    ") }
        $wb.Close($false)
      } elseif ($ext -in '.pptx', '.ppt', '.odp') {
        if (-not $ppt) { $ppt = New-Object -ComObject PowerPoint.Application }
        $p = $ppt.Presentations.Open($full, -1, 0, 0)
        $texts = @(); foreach ($s in $p.Slides) { foreach ($sh in $s.Shapes) { if ($sh.HasTextFrame) { $texts += $sh.TextFrame.TextRange.Text } } }
        $results += [pscustomobject]@{ File = (Split-Path $full -Leaf); OK = $true; Info = "slides=$($p.Slides.Count) text=" + (($texts -join ' / ') -replace '\s+', ' ').Substring(0, [Math]::Min(200, (($texts -join ' / ') -replace '\s+', ' ').Length)) }
        $p.Close()
      } else {
        $results += [pscustomobject]@{ File = (Split-Path $full -Leaf); OK = $false; Info = 'unsupported type' }
      }
    } catch {
      $results += [pscustomobject]@{ File = (Split-Path $full -Leaf); OK = $false; Info = $_.Exception.Message }
    }
  }
} finally {
  if ($word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
  if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
  if ($ppt) { $ppt.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) }
}
foreach ($r in $results) { "{0} {1}`n    {2}" -f ($(if ($r.OK) { 'PASS' } else { 'FAIL' })), $r.File, $r.Info }
