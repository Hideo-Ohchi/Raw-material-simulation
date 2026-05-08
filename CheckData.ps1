$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$wb = $excel.Workbooks.Open((Get-Location).Path + "\BOM copy 1.xlsm", $false, $false)

$wsIn = $wb.Sheets("シミュレーションイン")

Write-Host "=== Current Data Structure ==="
Write-Host "Row 1:"
For ($i = 1; $i -le 10; $i++) {
    $val = $wsIn.Cells(1, $i).Value
    if ($val -ne $null) {
        Write-Host "  Col $i : $val"
    }
}

Write-Host "`nRow 2:"
For ($i = 1; $i -le 10; $i++) {
    $val = $wsIn.Cells(2, $i).Value
    if ($val -ne $null) {
        Write-Host "  Col $i : $val"
    }
}

Write-Host "`nRow 3:"
For ($i = 1; $i -le 10; $i++) {
    $val = $wsIn.Cells(3, $i).Value
    if ($val -ne $null) {
        Write-Host "  Col $i : $val"
    }
}

Write-Host "`nTotal rows: $(($wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row))"

$wb.Close($false)
$excel.Quit()
Write-Host "`nDone"
