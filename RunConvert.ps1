$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open("$(Get-Location)\BOM copy 1.xlsm", $false, $false)

# マクロ実行
Write-Host "Running ConvertToHorizontalFormatV2..."
$excel.Run("シミュレーション.ConvertToHorizontalFormatV2")
Write-Host "Complete"

Start-Sleep -Seconds 2
$wb.Save()
$wb.Close($false)
$excel.Quit()
