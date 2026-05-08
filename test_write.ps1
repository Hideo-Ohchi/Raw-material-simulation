$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$book = $excel.Workbooks.Open("C:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm", 0, $false)
$ws = $book.Sheets("シミュレーションイン")

# 直接セルに書き込みテスト
Write-Host "Before: A1=$($ws.Cells(1,1).Value2)"
$ws.Cells(1,1).Value2 = "TEST_VALUE"
Write-Host "After write: A1=$($ws.Cells(1,1).Value2)"
$book.Save()
Write-Host "Saved"

# 再確認
Write-Host "After save: A1=$($ws.Cells(1,1).Value2)"
$book.Close($false)
$excel.Quit()