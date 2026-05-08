# Check-SimulationData.ps1
# シミュレーション入のデータと生産計画のデータを確認

param(
    [string]$ExcelFile = "原料シミュレーション.xlsm"
)

# Stop Excel
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# Create Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Open workbook
    $wb = $excel.Workbooks.Open((Resolve-Path $ExcelFile).Path, $false, $false)
    
    # Get sheets
    $wsConfig = $wb.Sheets("設定")
    $wsPlan = $wb.Sheets("生産計画")
    $wsIn = $wb.Sheets("シミュレーション入")
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "【1】設定シートの値" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    $b1 = $wsConfig.Range("B1").Value
    $b2 = $wsConfig.Range("B2").Value
    
    Write-Host "B1（開始年月）: $b1"
    Write-Host "B2（計画月数）: $b2"
    Write-Host ""
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "【2】シミュレーション入の月列" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    $inMonths = @()
    for ($c = 3; $c -le 10; $c++) {
        $v = $wsIn.Cells(1, $c).Value
        if ($v) {
            $inMonths += $v
            Write-Host "  C列（$([char](64+$c))）: $v"
        }
    }
    
    Write-Host "月列数: $($inMonths.Count)"
    Write-Host ""
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "【3】シミュレーション入のデータ行数" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    $inLastRow = $wsIn.Cells($wsIn.Rows.Count, 1).End(-4162).Row
    Write-Host "最終行: $inLastRow"
    Write-Host "ヘッダー: 1行"
    Write-Host "データ: $($inLastRow - 1)行"
    Write-Host ""
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "【4】生産計画のデータ件数（対象期間別）" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    $planLastRow = $wsPlan.Cells($wsPlan.Rows.Count, 1).End(-4162).Row
    
    $counts = @{}
    for ($r = 2; $r -le $planLastRow; $r++) {
        $code = $wsPlan.Cells($r, 1).Value
        $ym = $wsPlan.Cells($r, 2).Value
        $qty = $wsPlan.Cells($r, 3).Value
        
        if ($code -and $ym) {
            $ymStr = [string]$ym
            if (-not $counts.ContainsKey($ymStr)) {
                $counts[$ymStr] = 0
            }
            $counts[$ymStr] += 1
        }
    }
    
    Write-Host "生産計画に含まれる年月："
    $sorted = $counts.Keys | Sort-Object
    foreach ($ym in $sorted) {
        Write-Host "  $ym : $($counts[$ym])件"
    }
    Write-Host ""
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "【5】シミュレーション入の最初の5行" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    for ($r = 1; $r -le [Math]::Min(5, $inLastRow); $r++) {
        $line = "Row $r : "
        for ($c = 1; $c -le [Math]::Min(6, 3 + $inMonths.Count); $c++) {
            $v = $wsIn.Cells($r, $c).Value
            $line += "[$v] "
        }
        Write-Host $line
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "【分析】" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    if ($b1 -and $b2) {
        Write-Host "✓ 設定値が存在します"
        Write-Host "  開始年月: $b1"
        Write-Host "  計画月数: $b2"
        
        # Check if data matches
        $expectedMonths = 0
        if ($inMonths.Count -ne $b2) {
            Write-Host "⚠ 注意: シミュレーション入の月列数（$($inMonths.Count)）が計画月数（$b2）と一致しません"
        } else {
            Write-Host "✓ 月列数が計画月数と一致しています"
        }
        
        # Check if first month matches
        if ($inMonths[0] -ne $b1) {
            Write-Host "⚠ 注意: シミュレーション入の最初の月（$($inMonths[0])）が開始年月（$b1）と一致しません"
        } else {
            Write-Host "✓ 最初の月が開始年月と一致しています"
        }
    } else {
        Write-Host "⚠ 警告: 設定値が入力されていません"
    }
    
    Write-Host ""
    
} catch {
    Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
}

$wb.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
