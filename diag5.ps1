$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$book = $excel.Workbooks.Open("C:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm", 0, $false)
$wsIn = $book.Sheets("シミュレーションイン")

$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row
$lastCol = $wsIn.Cells(2, $wsIn.Columns.Count).End(-4159).Column

$rng = $wsIn.Range($wsIn.Cells(1,1), $wsIn.Cells($lastRow, $lastCol))
$d = $rng.Value2

# 月一覧（row3以降のC列）
$months = [System.Collections.Generic.SortedSet[object]]::new()
for ($r = 3; $r -le $d.GetUpperBound(0); $r++) {
    $m = $d[$r, 3]
    if ($m) { $months.Add($m) | Out-Null }
}
Write-Host "Unique months in data: $($months -join ', ')"
Write-Host "Total data rows: $($d.GetUpperBound(0) - 2)"

$book.Close($false)
$excel.Quit()