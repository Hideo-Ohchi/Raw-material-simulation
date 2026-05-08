$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$book = $excel.Workbooks.Open("C:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm", 0, $false)
$wsIn = $book.Sheets("シミュレーションイン")

# 行数と列数を個別に取得
$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row
$lastCol = $wsIn.Cells(2, $wsIn.Columns.Count).End(-4159).Column
Write-Host "lastRow=$lastRow, lastCol=$lastCol"

# 明示的な範囲で Value2 を取得
$rng = $wsIn.Range($wsIn.Cells(1,1), $wsIn.Cells($lastRow, $lastCol))
$rawData = $rng.Value2
Write-Host "rawData type: $($rawData.GetType().Name)"
Write-Host "rawData dims: $($rawData.GetLength(0)) x $($rawData.GetLength(1))"
Write-Host "A1=$($rawData[0,0]) B1=$($rawData[0,1]) C1=$($rawData[0,2])"
Write-Host "A2=$($rawData[1,0]) C2=$($rawData[1,2]) E2=$($rawData[1,4])"

$book.Close($false)
$excel.Quit()