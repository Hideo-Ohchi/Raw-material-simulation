$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$book = $excel.Workbooks.Open("C:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm", 0, $false)
$wsIn  = $book.Sheets("シミュレーションイン")
$wsOut = $book.Sheets("シミュレーションアウト")

$lastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row
$rng = $wsIn.Range($wsIn.Cells(1,1), $wsIn.Cells($lastRow, 5))
$d = $rng.Value2
Write-Host "Read $($d.GetUpperBound(0)) rows x $($d.GetUpperBound(1)) cols"

# データ収集（行3以降, 1-based）
$monthSet  = New-Object System.Collections.Generic.SortedSet[string]
$itemOrder = New-Object System.Collections.Generic.List[string]
$itemNames = @{}
$itemQty   = @{}

for ($r = 3; $r -le $d.GetUpperBound(0); $r++) {
    $code  = [string]($d[$r, 1])
    $name  = [string]($d[$r, 2])
    $month = [string]($d[$r, 3])
    $qty   = $d[$r, 5]
    if ($code -ne "" -and $month -ne "") {
        [void]$monthSet.Add($month)
        if (-not $itemNames.ContainsKey($code)) {
            [void]$itemOrder.Add($code)
            $itemNames[$code] = $name
        }
        $itemQty["${code}|${month}"] = $qty
    }
}

$months = @($monthSet)
Write-Host "Items=$($itemOrder.Count), Months=$($months -join ',')"

# 出力配列（SetValueで2D配列に書き込み）
$nRows = $itemOrder.Count + 2
$nCols = $months.Count + 2
$out   = [System.Array]::CreateInstance([object], $nRows, $nCols)

# 行0: ラベル行
$out.SetValue("品目コード", 0, 0)
$out.SetValue("品目名",   0, 1)
for ($j = 0; $j -lt $months.Count; $j++) {
    $out.SetValue([double]$months[$j], 0, $j + 2)
}

# 行1+: データ
for ($i = 0; $i -lt $itemOrder.Count; $i++) {
    $code = $itemOrder[$i]
    $out.SetValue($code,            $i + 1, 0)
    $out.SetValue($itemNames[$code], $i + 1, 1)
    for ($j = 0; $j -lt $months.Count; $j++) {
        $key = "${code}|$($months[$j])"
        if ($itemQty.ContainsKey($key)) {
            $out.SetValue($itemQty[$key], $i + 1, $j + 2)
        }
    }
}

# シミュレーションイン に書き込み
$wsIn.Cells.Clear()
$wsIn.Range($wsIn.Cells(1,1), $wsIn.Cells($nRows - 1, $nCols)).Value2 = $out
Write-Host "シミュレーションイン: Done"

# シミュレーションアウト のヘッダー
$hdr = [System.Array]::CreateInstance([object], 1, $nCols)
$hdr.SetValue("品目コード", 0, 0)
$hdr.SetValue("品目名",    0, 1)
for ($j = 0; $j -lt $months.Count; $j++) {
    $hdr.SetValue([double]$months[$j], 0, $j + 2)
}
$wsOut.Cells.Clear()
$wsOut.Range($wsOut.Cells(1,1), $wsOut.Cells(1, $nCols)).Value2 = $hdr
Write-Host "シミュレーションアウト: Done"

$book.Save()
$book.Close($false)
$excel.Quit()
Write-Host "Completed"