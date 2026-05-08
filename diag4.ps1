$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$book = $excel.Workbooks.Open("C:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm", 0, $false)
$wsIn = $book.Sheets("シミュレーションイン")

$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row
$lastCol = $wsIn.Cells(2, $wsIn.Columns.Count).End(-4159).Column
Write-Host "lastRow=$lastRow lastCol=$lastCol"

$rng = $wsIn.Range($wsIn.Cells(1,1), $wsIn.Cells([Math]::Min(5, $lastRow), $lastCol))
$d = $rng.Value2
for ($r = 1; $r -le $d.GetUpperBound(0); $r++) {
    $line = ""
    for ($c = 1; $c -le $d.GetUpperBound(1); $c++) {
        $line += "[$r,$c]=$($d[$r,$c]) | "
    }
    Write-Host $line
}

$book.Close($false)
$excel.Quit()