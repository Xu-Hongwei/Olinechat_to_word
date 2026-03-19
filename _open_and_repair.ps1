$ErrorActionPreference='Stop'
$src = (Resolve-Path '.\native_selfcheck.docx').Path
$dst = Join-Path (Get-Location) 'native_selfcheck_repaired.docx'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($src, $false, $true, $false, '', '', $true, '', '', 0, 0, $false, $true)
  $doc.SaveAs2($dst, 16)
  $doc.Close()
  'REPAIRED=' + $dst
}
finally {
  $word.Quit()
}
