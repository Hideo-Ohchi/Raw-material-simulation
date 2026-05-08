Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open((Get-Location).Path + "\BOM copy 1.xlsm", $false, $false)

Write-Host "Workbook opened"
Write-Host "Sheets count: $($wb.Sheets.Count)"

# Get sheet count
$wsIn = $wb.Sheets(17)  # Try sheet 17 for simulation input
Write-Host "Sheet 17 name: $($wsIn.Name)"

$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row
Write-Host "Last row with data: $lastRow"

if ($lastRow -gt 0) {
    Write-Host "Data exists"
    for ($i = 1; $i -le [Math]::Min(3, $lastRow); $i++) {
        Write-Host "Row $i : $($wsIn.Cells($i, 1).Value) | $($wsIn.Cells($i, 2).Value) | $($wsIn.Cells($i, 3).Value)"
    }
} else {
    Write-Host "No data found"
}

$wb.Close($false)
$excel.Quit()
Write-Host "Done"
