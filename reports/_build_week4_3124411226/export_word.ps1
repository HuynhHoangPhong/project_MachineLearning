param([string]$InputDocx, [string]$OutputPdf)
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $document = $word.Documents.Open($InputDocx, $false, $true, $false)
    $document.Repaginate()
    $document.ExportAsFixedFormat($OutputPdf, 17)
    Write-Output ('Exported pages: ' + $document.ComputeStatistics(2))
} finally {
    if ($null -ne $document) {
        $document.Close(0)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    if ($null -ne $word) {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}
