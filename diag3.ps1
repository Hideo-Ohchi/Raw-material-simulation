$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$book = $excel.Workbooks.Open("C:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm", 0, $false)
$wsIn = $book.Sheets("シミュレーションイン")

$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row
$lastCol = $wsIn.Cells(2, $wsIn.Columns.Count).End(-4159).Column

$rng = $wsIn.Range($wsIn.Cells(1,1), $wsIn.Cells($lastRow, $lastCol))
$rawData = $rng.Value2

# 配列のLowerBound確認
Write-Host "LB0=$($rawData.GetLowerBound(0)) UB0=$($rawData.GetUpperBound(0))"
Write-Host "LB1=$($rawData.GetLowerBound(1)) UB1=$($rawData.GetUpperBound(1))"
Write-Host "A1(1,1)=$($rawData[1,1])"
Write-Host "A2(2,1)=$($rawData[2,1])"
Write-Host "B2(2,2)=$($rawData[2,2])"
Write-Host "C2(2,3)=$($rawData[2,3])"
Write-Host "D2(2,4)=$($rawData[2,4])"
Write-Host "E2(2,5)=$($rawData[2,5])"

$book.Close($false)
$excel.Quit()