cd 'c:\Users\119351\Desktop\原料シミュレーション'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

$wb = $excel.Workbooks.Open("$(Get-Location)\BOM copy 1.xlsm")
foreach ($c in $wb.VBProject.VBComponents) {
    if ($c.Name -eq "シミュレーション") {
        $wb.VBProject.VBComponents.Remove($c)
        Write-Host "Old module removed"
    }
}

$wb.VBProject.VBComponents.Import("$(Get-Location)\Simulation_Module.bas")
Write-Host "Module imported"

$wb.Save()
$wb.Close($false)
$excel.Quit()
Write-Host "Done"
