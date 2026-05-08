Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open((Get-Location).Path + "\BOM copy 1.xlsm", $false, $false)
$wsIn = $wb.Sheets(17)

$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row

Write-Host "Reading data..."
# Load data starting from row 3
$data = @()
for ($r = 3; $r -le $lastRow; $r++) {
    $item = @{
        Code = $wsIn.Cells($r, 1).Value2
        Name = $wsIn.Cells($r, 2).Value2
        Month = $wsIn.Cells($r, 3).Value2
        Qty1 = $wsIn.Cells($r, 4).Value2
        Qty2 = $wsIn.Cells($r, 5).Value2
    }
    $data += $item
}

Write-Host "Loaded $($data.Count) rows"

# Collect unique months and items
$months = @()
$items = @()

foreach ($d in $data) {
    $m = [string]$d.Month
    $c = [string]$d.Code
    
    if (-not ($months -contains $m) -and $m -ne "") { $months += $m }
    if (-not ($items -contains $c) -and $c -ne "") { $items += $c }
}

# Sort months
$months = $months | Sort-Object
Write-Host "Months: $($months -join ',')"
Write-Host "Items: $($items.Count)"

# Clear sheet
$wsIn.Cells.Clear()

# Write headers
$wsIn.Cells(1, 1).Value2 = "品目コード"
$wsIn.Cells(1, 2).Value2 = "品目名"
for ($m = 0; $m -lt $months.Count; $m++) {
    $wsIn.Cells(1, 3 + $m).Value2 = $months[$m]
}

Write-Host "Headers written"

# Write data
for ($i = 0; $i -lt $items.Count; $i++) {
    $itemCode = $items[$i]
    $wsIn.Cells(2 + $i, 1).Value2 = $itemCode
    
    # Find item name
    $itemName = ""
    foreach ($d in $data) {
        if ([string]$d.Code -eq $itemCode) {
            $itemName = $d.Name
            break
        }
    }
    $wsIn.Cells(2 + $i, 2).Value2 = $itemName
    
    # Write quantities
    for ($m = 0; $m -lt $months.Count; $m++) {
        $qty = 0
        foreach ($d in $data) {
            if ([string]$d.Code -eq $itemCode -and [string]$d.Month -eq $months[$m]) {
                $qty = $d.Qty1
                break
            }
        }
        $wsIn.Cells(2 + $i, 3 + $m).Value2 = $qty
    }
}

Write-Host "Data written: $($items.Count) items x $($months.Count) months"

$wb.Save()
$wb.Close($false)
$excel.Quit()

Write-Host "Complete"
