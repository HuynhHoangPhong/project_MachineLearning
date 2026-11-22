param([string]$DocumentPath,[string]$OperationsPath)
$ErrorActionPreference='Stop'
$word=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
$document=$null
foreach ($item in $word.Documents) {if ($item.FullName -eq $DocumentPath) {$document=$item;break}}
if ($null -eq $document) {throw 'The intended document is not open in the active Word instance.'}
if ($document.ReadOnly) {throw 'The document is read-only.'}
$ops=Get-Content -LiteralPath $OperationsPath -Raw -Encoding utf8 | ConvertFrom-Json
$targets=@()
foreach ($op in $ops.paragraphs) {
    $matches=@()
    foreach ($p in $document.Paragraphs) {
        $text=$p.Range.Text.TrimEnd([char]13,[char]7)
        if ($text -eq $op.old) {$matches+=,$p.Range.Duplicate}
    }
    if ($matches.Count -ne 1) {throw ('Preflight paragraph mismatch: '+$op.old.Substring(0,[Math]::Min(80,$op.old.Length)))}
    $targets+=,[pscustomobject]@{Range=$matches[0];New=$op.new;Start=$matches[0].Start}
}
$table=$document.Tables.Item(7)
for ($i=0;$i -lt 5;$i++) {
    for ($j=0;$j -lt 3;$j++) {
        $actual=$table.Cell($i+2,$j+1).Range.Text.TrimEnd([char]13,[char]7)
        if ($actual -ne $ops.old_table_21[$i][$j]) {throw ('Preflight table mismatch at '+$i+','+$j)}
    }
}
$anchor=$null
foreach ($p in $document.Paragraphs) {
    if ($p.Range.Text.TrimEnd([char]13,[char]7) -eq $ops.insert_before) {$anchor=$p.Range.Duplicate;break}
}
if ($null -eq $anchor) {throw 'Insertion heading not found.'}
$word.UndoRecord.StartCustomRecord('Review week 4 report content and length')
try {
    foreach ($target in ($targets | Sort-Object Start -Descending)) {
        if ($null -eq $target.New) {
            [void]$target.Range.Delete()
        } else {
            $range=$target.Range.Duplicate
            $range.End=$range.End-1
            $range.Text=[string]$target.New
        }
    }
    for ($i=0;$i -lt 5;$i++) {
        for ($j=0;$j -lt 3;$j++) {
            $range=$table.Cell($i+2,$j+1).Range.Duplicate
            $range.End=$range.End-2
            $range.Text=[string]$ops.table_21[$i][$j]
            $range.Font.Name='Times New Roman'
            $range.Font.Size=11
        }
    }
    $start=$anchor.Start
    $newText=($ops.insertions -join "`r")+"`r"
    $insertion=$document.Range($start,$start)
    $insertion.InsertBefore($newText)
    $inserted=$document.Range($start,$start+$newText.Length)
    $inserted.Style='Normal'
    $inserted.Font.Name='Times New Roman'
    $inserted.Font.Size=11
    $inserted.Font.Bold=0
    $inserted.ParagraphFormat.LineSpacingRule=5
    $inserted.ParagraphFormat.LineSpacing=13.8
    $inserted.ParagraphFormat.SpaceBefore=0
    $inserted.ParagraphFormat.SpaceAfter=6
    $inserted.ParagraphFormat.KeepWithNext=0
    $inserted.Paragraphs.Item(1).Range.Font.Bold=-1
    $inserted.Paragraphs.Item(1).KeepWithNext=-1
    $document.Repaginate()
    $document.Save()
    Write-Output ('Saved in place: '+$document.FullName)
    Write-Output ('Pages: '+$document.ComputeStatistics(2))
} finally {
    $word.UndoRecord.EndCustomRecord()
}
