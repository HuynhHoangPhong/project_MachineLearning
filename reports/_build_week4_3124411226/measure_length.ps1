param([string]$InputDocx,[string]$OutputJson)
$ErrorActionPreference='Stop'
$word=$null
$document=$null
try {
  $word=New-Object -ComObject Word.Application
  $word.Visible=$false
  $word.DisplayAlerts=0
  $word.AutomationSecurity=3
  $document=$word.Documents.Open($InputDocx,$false,$true,$false)
  $document.Repaginate()
  $locations=@()
  foreach ($par in $document.Paragraphs) {
    $t=$par.Range.Text.Trim([char]13,[char]7,[char]12)
    if ($t -match '^[567]\. ' -or $t -match '^6\.3\. ' -or $t -match '^6\.4\. ') {
      $locations+=@{text=$t;page=$par.Range.Information(3);y_pt=$par.Range.Information(6)}
    }
  }
  $result=@{pages=$document.ComputeStatistics(2);word_count=$document.ComputeStatistics(0);headings=$locations}
  $result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $OutputJson -Encoding utf8
  $result | ConvertTo-Json -Depth 5
} finally {
  if ($null -ne $document) {$document.Close(0);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)}
  if ($null -ne $word) {$word.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}
}
