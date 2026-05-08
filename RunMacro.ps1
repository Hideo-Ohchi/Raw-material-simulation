$excel = New-Object -ComObject Excel.Application
$excel.AutomationSecurity = 1  # msoAutomationSecurityLow

$wb = $excel.Workbooks.Open("$(Get-Location)\BOM copy 1.xlsm", $false, $false)

# シート確認
Write-Host "Sheets: $($wb.Sheets.Count)"
foreach ($sheet in $wb.Sheets) {
    Write-Host "  - $($sheet.Name)"
}

Start-Sleep -Seconds 1

# 可視化してマクロ実行
$excel.Visible = $true
$excel.DisplayAlerts = $false

Write-Host "Executing macro..."
try {
    $result = $excel.Run("シミュレーション.ConvertToHorizontalFormatV2")
    Write-Host "Result: $result"
} catch {
    Write-Host "Error: $_"
}

Start-Sleep -Seconds 2
$wb.Save()
$wb.Close($false)
$excel.Quit()
Write-Host "Done"
