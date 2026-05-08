$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AutomationSecurity = 1

$workbookPath = "C:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm"
$book = $excel.Workbooks.Open($workbookPath, 0, $false)
Write-Host "File opened"

$wsIn = $book.Sheets("シミュレーションイン")

# 配列一括読み込み（高速）
$usedRange = $wsIn.UsedRange
$rawData = $usedRange.Value2

if ($rawData -eq $null) {
    Write-Host "No data"
    $book.Close($false); $excel.Quit(); exit
}
Write-Host "Data loaded"

$rows = $rawData.GetLength(0)

# 月一覧収集（C列=index 2, 0-based）
$months = [System.Collections.Generic.List[object]]::new()
for ($i = 1; $i -lt $rows; $i++) {
    $month = $rawData[$i, 2]
    if ($month -and -not $months.Contains($month)) { $months.Add($month) }
}
$months = [System.Linq.Enumerable]::OrderBy($months, [Func[object,object]]{ param($x) $x })
Write-Host "Months: $($months.Count)"

# 品目辞書作成
$itemKeys  = [System.Collections.Generic.List[string]]::new()
$itemNames = [System.Collections.Generic.Dictionary[string,object]]::new()
$itemData  = [System.Collections.Generic.Dictionary[string,System.Collections.Generic.Dictionary[object,object]]]::new()

for ($i = 1; $i -lt $rows; $i++) {
    $code  = "$($rawData[$i, 0])"
    $name  = $rawData[$i, 1]
    $month = $rawData[$i, 2]
    $qty   = $rawData[$i, 4]
    if ($code -and $code -ne "") {
        if (-not $itemKeys.Contains($code)) {
            $itemKeys.Add($code)
            $itemNames[$code] = $name
            $itemData[$code]  = [System.Collections.Generic.Dictionary[object,object]]::new()
        }
        if ($month) { $itemData[$code][$month] = $qty }
    }
}
Write-Host "Items: $($itemKeys.Count)"

# 出力配列作成
$outRows = $itemKeys.Count + 1
$outCols = $months.Count + 2
$outArr  = New-Object 'object[,]' $outRows, $outCols

$outArr[0, 0] = "品目コード"
$outArr[0, 1] = "品目名"
$mList = @($months)
for ($j = 0; $j -lt $mList.Length; $j++) { $outArr[0, $j + 2] = $mList[$j] }

for ($r = 0; $r -lt $itemKeys.Count; $r++) {
    $code = $itemKeys[$r]
    $outArr[$r + 1, 0] = $code
    $outArr[$r + 1, 1] = $itemNames[$code]
    for ($j = 0; $j -lt $mList.Length; $j++) {
        $m = $mList[$j]
        if ($itemData[$code].ContainsKey($m)) { $outArr[$r + 1, $j + 2] = $itemData[$code][$m] }
    }
}

# シートをクリアして一括書き込み
$wsIn.Cells.Clear()
$wsIn.Range($wsIn.Cells(1, 1), $wsIn.Cells($outRows, $outCols)).Value2 = $outArr
Write-Host "シミュレーションイン: Done"

# シミュレーションアウト ヘッダー更新
$wsOut = $book.Sheets("シミュレーションアウト")
$wsOut.Cells.Clear()
$hdr = New-Object 'object[,]' 1, $outCols
$hdr[0, 0] = "品目コード"; $hdr[0, 1] = "品目名"
for ($j = 0; $j -lt $mList.Length; $j++) { $hdr[0, $j + 2] = $mList[$j] }
$wsOut.Range($wsOut.Cells(1, 1), $wsOut.Cells(1, $outCols)).Value2 = $hdr
Write-Host "シミュレーションアウト: Done"

$book.Save(); $book.Close($false); $excel.Quit()
Write-Host "Completed"