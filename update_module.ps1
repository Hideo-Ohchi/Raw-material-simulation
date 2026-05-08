$excelFile = "c:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm"
$basFile = "c:\Users\119351\Desktop\原料シミュレーション\Simulation_Module.bas"

# ファイル存在確認
if (-not (Test-Path $excelFile)) {
    Write-Host "Error: File not found - $excelFile"
    exit
}

if (-not (Test-Path $basFile)) {
    Write-Host "Error: Module file not found - $basFile"
    exit
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

$excelFileFullPath = (Resolve-Path $excelFile).Path
Write-Host "Opening: $excelFileFullPath"

$wb = $excel.Workbooks.Open($excelFileFullPath)

# 既存モジュール削除
foreach ($c in $wb.VBProject.VBComponents) {
    if ($c.Name -eq "シミュレーション") {
        $wb.VBProject.VBComponents.Remove($c)
        Write-Host "Old module removed"
    }
}

# 新規モジュールインポート
$wb.VBProject.VBComponents.Import($basFile)
Write-Host "New module imported"

# 保存
$wb.Save()

# クローズ
$wb.Close($false)
$excel.Quit()

Write-Host "Complete"
