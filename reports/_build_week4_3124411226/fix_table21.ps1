param([string]$OperationsPath)
$ErrorActionPreference='Stop'
$word=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
$doc=$word.Documents.Item('3124411226_Huong_dan_bo_sung_bao_cao_week4.docx')
$ops=Get-Content -LiteralPath $OperationsPath -Raw -Encoding utf8 | ConvertFrom-Json
$table=$doc.Tables.Item(7)
$word.UndoRecord.StartCustomRecord('Correct table punctuation')
try {
    for ($i=0;$i -lt 5;$i++) {
        for ($j=0;$j -lt 3;$j++) {
            $range=$table.Cell($i+2,$j+1).Range.Duplicate
            $range.End=$range.End-1
            $range.Text=[string]$ops.table_21[$i][$j]
            $range.Font.Size=11
        }
    }
    $doc.Save()
    Write-Output 'Corrected table 21 and saved.'
} finally {$word.UndoRecord.EndCustomRecord()}
