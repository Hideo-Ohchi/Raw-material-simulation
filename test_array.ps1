$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$book = $excel.Workbooks.Open("C:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm", 0, $false)
$ws = $book.Sheets("シミュレーションイン")

# New-Object + SetValue を使った2D配列テスト
$arr = New-Object 'object[,]' 3,5
$arr.SetValue("品目コード", 0, 0)
$arr.SetValue("品目名",   0, 1)
$arr.SetValue(202605,      0, 2)
$arr.SetValue(202606,      0, 3)
$arr.SetValue(202607,      0, 4)
$arr.SetValue("TEST001",  1, 0)
$arr.SetValue("テスト品目", 1, 1)
$arr.SetValue(100,         1, 2)
$arr.SetValue(200,         1, 3)
$arr.SetValue(300,         1, 4)

# 既存データをクリアして書き込み
$ws.Cells.Clear()
$ws.Range($ws.Cells(1,1), $ws.Cells(3,5)).Value2 = $arr

# 確認
$d2 = $ws.Range($ws.Cells(1,1), $ws.Cells(3,5)).Value2
for ($r=1;$r-le$d2.GetUpperBound(0);$r++){
    $line=""
    for($c=1;$c-le$d2.GetUpperBound(1);$c++){$line+="[$r,$c]=$($d2[$r,$c]) | "}
    Write-Host $line
}

$book.Close($false)
$excel.Quit()