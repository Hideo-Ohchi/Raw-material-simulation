$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open((Get-Location).Path + "\BOM copy 1.xlsm", $false, $false)
$wsIn = $wb.Sheets(17)

$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row
$lastCol = $wsIn.Cells(1, $wsIn.Columns.Count).End(-4159).Column

Write-Host "=== After Transformation ==="
Write-Host "Total rows: $lastRow"
Write-Host "Total cols: $lastCol"
Write-Host ""
Write-Host "Headers (Row 1):"
for ($i = 1; $i -le [Math]::Min(6, $lastCol); $i++) {
    $val = $wsIn.Cells(1, $i).Value2
    Write-Host "  Col $i : [$val]"
}

Write-Host ""
Write-Host "First 3 data rows:"
for ($r = 2; $r -le [Math]::Min(4, $lastRow); $r++) {
    $row_output = "Row $r : "
    for ($i = 1; $i -le [Math]::Min(6, $lastCol); $i++) {
        $val = $wsIn.Cells($r, $i).Value2
        $row_output += "[$val] "
    }
    Write-Host $row_output
}

$wb.Close($false)
$excel.Quit()
Write-Host ""
Write-Host "Done"
