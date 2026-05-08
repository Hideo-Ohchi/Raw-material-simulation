$excelFile = 'c:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm'
$basFile = 'c:\Users\119351\Desktop\原料シミュレーション\Simulation_Module.bas'

if (-not (Test-Path $excelFile)) { Write-Host "Excel not found"; exit }
if (-not (Test-Path $basFile)) { Write-Host "Module not found"; exit }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$wb = $excel.Workbooks.Open((Resolve-Path $excelFile).Path)

foreach ($c in $wb.VBProject.VBComponents) {
    if ($c.Name -eq "シミュレーション") {
        $wb.VBProject.VBComponents.Remove($c)
        Write-Host "Old module removed"
    }
}

$wb.VBProject.VBComponents.Import((Resolve-Path $basFile).Path)
Write-Host "Module imported"

$wb.Save()
$wb.Close($false)
$excel.Quit()
Write-Host "Done"
