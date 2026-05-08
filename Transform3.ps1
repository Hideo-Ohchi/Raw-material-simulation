Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open((Get-Location).Path + "\BOM copy 1.xlsm", $false, $false)
$wsIn = $wb.Sheets(17)

$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row

Write-Host "Loading data..."
$data = @()
for ($r = 3; $r -le $lastRow; $r++) {
    $data += @{
        Code = $wsIn.Cells($r, 1).Value2
        Name = $wsIn.Cells($r, 2).Value2
        Month = $wsIn.Cells($r, 3).Value2
        Qty = $wsIn.Cells($r, 4).Value2
    }
}

$months = @()
$items = @()
foreach ($d in $data) {
    $m = [string]$d.Month
    $c = [string]$d.Code
    if (-not ($months -contains $m) -and $m -ne "") { $months += $m }
    if (-not ($items -contains $c) -and $c -ne "") { $items += $c }
}

$months = $months | Sort-Object
Write-Host "Months: $($months.Count), Items: $($items.Count)"

$wsIn.Cells.Clear()
$wsIn.Cells(1, 1).Value2 = "ItemCode"
$wsIn.Cells(1, 2).Value2 = "ItemName"
for ($m = 0; $m -lt $months.Count; $m++) {
    $wsIn.Cells(1, 3 + $m).Value2 = $months[$m]
}

for ($i = 0; $i -lt $items.Count; $i++) {
    $itemCode = $items[$i]
    $wsIn.Cells(2 + $i, 1).Value2 = $itemCode
    $itemName = ($data | Where-Object {[string]$_.Code -eq $itemCode} | Select-Object -First 1).Name
    $wsIn.Cells(2 + $i, 2).Value2 = $itemName
    
    for ($m = 0; $m -lt $months.Count; $m++) {
        $qty = ($data | Where-Object {[string]$_.Code -eq $itemCode -and [string]$_.Month -eq $months[$m]} | Select-Object -First 1).Qty
        $wsIn.Cells(2 + $i, 3 + $m).Value2 = if($qty) {$qty} else {0}
    }
}

Write-Host "Saving..."
$wb.Save()
$wb.Close($false)
$excel.Quit()
Write-Host "Complete"