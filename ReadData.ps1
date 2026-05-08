$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open((Get-Location).Path + "\BOM copy 1.xlsm", $false, $false)
$wsIn = $wb.Sheets(17)

$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row

Write-Host "Total rows: $lastRow"
Write-Host ""
Write-Host "First 5 rows:"

for ($i = 1; $i -le [Math]::Min(5, $lastRow); $i++) {
    $c1 = $wsIn.Cells($i, 1).Value2
    $c2 = $wsIn.Cells($i, 2).Value2
    $c3 = $wsIn.Cells($i, 3).Value2
    $c4 = $wsIn.Cells($i, 4).Value2
    $c5 = $wsIn.Cells($i, 5).Value2
    Write-Host "Row $i : [$c1] [$c2] [$c3] [$c4] [$c5]"
}

$wb.Close($false)
$excel.Quit()
Write-Host ""
Write-Host "Done"
