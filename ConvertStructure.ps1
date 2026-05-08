$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

$workbookPath = "$(Get-Location)\BOM copy 1.xlsm"
$book = $excel.Workbooks.Open($workbookPath)
$wsIn = $book.Sheets("シミュレーションイン")

$lastRow = $wsIn.UsedRange.Rows.Count

if ($lastRow -le 1) {
    Write-Host "No data to convert"
    $book.Close($false)
    $excel.Quit()
    exit
}

# Create temp sheet
try { $book.Sheets("_Temp").Delete() } catch {}
$wsTemp = $book.Sheets.Add()
$wsTemp.Name = "_Temp"

$wsTemp.Cells(1, 1) = "品目コード"
$wsTemp.Cells(1, 2) = "品目名"

# Collect months
$months = @()
for ($i = 2; $i -le $lastRow; $i++) {
    $month = $wsIn.Cells($i, 3).Value
    if ($month -and $month -ne "") {
        if ($months -notcontains $month) {
            $months += $month
        }
    }
}

# Sort months
$months = $months | Sort-Object

# Set month headers
for ($j = 0; $j -lt $months.Count; $j++) {
    $wsTemp.Cells(1, $j + 3) = $months[$j]
}

# Fill data
$tempRow = 2
$currentItem = ""

for ($i = 2; $i -le $lastRow; $i++) {
    $itemCode = $wsIn.Cells($i, 1).Value
    $itemName = $wsIn.Cells($i, 2).Value
    $yearMonth = $wsIn.Cells($i, 3).Value
    $simQty = $wsIn.Cells($i, 5).Value
    
    if ($itemCode -and $itemCode -ne "") {
        if ($itemCode -ne $currentItem) {
            $currentItem = $itemCode
            $tempRow++
            $wsTemp.Cells($tempRow, 1) = $itemCode
            $wsTemp.Cells($tempRow, 2) = $itemName
        }
        
        # Find month column
        for ($j = 0; $j -lt $months.Count; $j++) {
            if ($months[$j] -eq $yearMonth) {
                $wsTemp.Cells($tempRow, $j + 3) = $simQty
                break
            }
        }
    }
}

# Copy back to original sheet
$wsIn.Cells.Clear()
$lastRow = $wsTemp.UsedRange.Rows.Count
$colCount = $months.Count + 2

$source = $wsTemp.Range($wsTemp.Cells(1, 1), $wsTemp.Cells($lastRow, $colCount))
$source.Copy()
$wsIn.Range("A1").PasteSpecial(-4104)

# Delete temp sheet
$wsTemp.Delete()

# Save
$book.Save()

Write-Host "シミュレーションイン: Data structure converted"

# Also update output sheet header
$wsOut = $book.Sheets("シミュレーションアウト")
$wsOut.Cells(1, 1) = "品目コード"
$wsOut.Cells(1, 2) = "品目名"
for ($j = 0; $j -lt $months.Count; $j++) {
    $wsOut.Cells(1, $j + 3) = $months[$j]
}

$book.Save()
$book.Close($false)
$excel.Quit()

Write-Host "Conversion completed successfully"
