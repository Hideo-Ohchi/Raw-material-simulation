$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$book = $excel.Workbooks.Open("C:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm", 0, $false)
$wsIn = $book.Sheets("シミュレーションイン")

$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
$lastCol = $wsIn.Cells(1, $wsIn.Columns.Count).End(-4159).Column  # xlToLeft = -4159
Write-Host "Rows: $lastRow, Cols: $lastCol"
Write-Host "A1: $($wsIn.Cells(1,1).Value)"
Write-Host "A2: $($wsIn.Cells(2,1).Value)"
Write-Host "C2: $($wsIn.Cells(2,3).Value)"
Write-Host "E2: $($wsIn.Cells(2,5).Value)"

$book.Close($false)
$excel.Quit()