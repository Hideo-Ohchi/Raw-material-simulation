$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1
$book = $excel.Workbooks.Open("C:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm", 0, $false)
$ws = $book.Sheets("シミュレーションイン")
$lr = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row
$lc = $ws.Cells(2, $ws.Columns.Count).End(-4159).Column
Write-Host "Rows=$lr Cols=$lc"
$d = $ws.Range($ws.Cells(1,1), $ws.Cells([Math]::Min(4,$lr), $lc)).Value2
for ($r=1;$r-le$d.GetUpperBound(0);$r++){
    $line=""
    for($c=1;$c-le$d.GetUpperBound(1);$c++){$line+="[$r,$c]=$($d[$r,$c]) | "}
    Write-Host $line
}
$book.Close($false)
$excel.Quit()