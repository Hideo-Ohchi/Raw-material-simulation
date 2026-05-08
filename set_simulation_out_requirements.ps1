Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open("$(Get-Location)\BOM copy 1.xlsm", $false, $false)
$wsIn = $wb.Sheets('シミュレーションイン')
$wsOut = $wb.Sheets('シミュレーションアウト')
$wsB = $wb.Sheets('BOM')
$wsM = $wb.Sheets('品目マスタ')

$lastM = $wsM.Cells($wsM.Rows.Count,1).End(-4162).Row
$nameMap = @{}
$kindMap = @{}
for($r=2;$r -le $lastM;$r++){
    $code = ([string]$wsM.Cells($r,1).Value2).Trim()
    if(-not $code){ continue }
    $nameMap[$code] = [string]$wsM.Cells($r,2).Value2
    $kindMap[$code] = [string]$wsM.Cells($r,3).Value2
}

$lastB = $wsB.Cells($wsB.Rows.Count,1).End(-4162).Row
$bomMap = @{}
for($r=2;$r -le $lastB;$r++){
    $p = ([string]$wsB.Cells($r,1).Value2).Trim()
    $c = ([string]$wsB.Cells($r,2).Value2).Trim()
    if(-not $p -or -not $c){ continue }
    $q = 0.0
    try { $q = [double]$wsB.Cells($r,3).Value2 } catch { $q = 0.0 }
    if($q -eq 0.0){ continue }
    if(-not $bomMap.ContainsKey($p)){ $bomMap[$p] = New-Object System.Collections.ArrayList }
    [void]$bomMap[$p].Add(@($c,$q))
}

$lastRowIn = $wsIn.Cells($wsIn.Rows.Count,1).End(-4162).Row
$lastColIn = $wsIn.Cells(1,$wsIn.Columns.Count).End(-4159).Column
$months = @()
for($c=3;$c -le $lastColIn;$c++){
    $m = [string]$wsIn.Cells(1,$c).Value2
    if($m){ $months += $m }
}
$monthCount = $months.Count

$rawReq = @{}
$maxDepth = 20

function Add-RawQty {
    param([string]$code,[int]$mIdx,[double]$qty)
    if(-not $rawReq.ContainsKey($code)){
        $rawReq[$code] = New-Object 'double[]' $monthCount
    }
    $rawReq[$code][$mIdx] += $qty
}

function Expand-Req {
    param([string]$code,[double]$factor,[int]$mIdx,[hashtable]$visited,[int]$depth)
    if($factor -eq 0.0 -or $depth -gt $maxDepth){ return }
    if($visited.ContainsKey($code)){ return }

    $kind = ''
    if($kindMap.ContainsKey($code)){ $kind = $kindMap[$code] }
    if($kind -eq '原料'){
        Add-RawQty -code $code -mIdx $mIdx -qty $factor
        return
    }

    if(-not $bomMap.ContainsKey($code)){ return }

    $visited[$code] = $true
    foreach($pair in $bomMap[$code]){
        Expand-Req -code ([string]$pair[0]) -factor ($factor * [double]$pair[1]) -mIdx $mIdx -visited $visited -depth ($depth + 1)
    }
    $visited.Remove($code) | Out-Null
}

for($r=2;$r -le $lastRowIn;$r++){
    $prod = ([string]$wsIn.Cells($r,1).Value2).Trim()
    if(-not $prod){ continue }
    for($mi=0;$mi -lt $monthCount;$mi++){
        $q = 0.0
        try { $q = [double]$wsIn.Cells($r,3+$mi).Value2 } catch { $q = 0.0 }
        if($q -ne 0.0){
            $visited = @{}
            Expand-Req -code $prod -factor $q -mIdx $mi -visited $visited -depth 0
        }
    }
}

$codes = @($rawReq.Keys) | Sort-Object
$validCodes = New-Object System.Collections.ArrayList
foreach($code in $codes){
    $arr = $rawReq[$code]
    $has = $false
    for($i=0;$i -lt $monthCount;$i++){ if([math]::Abs($arr[$i]) -gt 1e-12){ $has = $true; break } }
    if($has){ [void]$validCodes.Add($code) }
}

$rowCount = 1 + $validCodes.Count
$colCount = 2 + $monthCount
$out = New-Object 'object[,]' $rowCount, $colCount

$out[0,0] = '品目コード'
$out[0,1] = '品目名'
for($mi=0;$mi -lt $monthCount;$mi++){ $out[0,(2+$mi)] = $months[$mi] }

for($ri=0;$ri -lt $validCodes.Count;$ri++){
    $code = [string]$validCodes[$ri]
    $out[(1+$ri),0] = $code
    $out[(1+$ri),1] = $(if($nameMap.ContainsKey($code)){$nameMap[$code]}else{''})
    $arr = $rawReq[$code]
    for($mi=0;$mi -lt $monthCount;$mi++){
        $out[(1+$ri),(2+$mi)] = [math]::Round([double]$arr[$mi],6)
    }
}

$wsOut.Cells.Clear()
$wsOut.Range('A1').Resize($rowCount,$colCount).Value2 = $out
$wb.Save()

Write-Host ("SimulationOut rows={0}, rawItems={1}, months={2}" -f $rowCount, $validCodes.Count, $monthCount)
for($r=1;$r -le [Math]::Min(6,$rowCount);$r++){
    $line = "R${r}: "
    for($c=1;$c -le [Math]::Min(6,$colCount);$c++){
        $line += "[$($wsOut.Cells($r,$c).Value2)] "
    }
    Write-Host $line
}

$wb.Close($false)
$excel.Quit()


