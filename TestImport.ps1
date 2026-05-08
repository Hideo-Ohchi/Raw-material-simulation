$excel = New-Object -ComObject Excel.Application
$excel.AutomationSecurity = 1
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open((Get-Location).Path + "\BOM copy 1.xlsm", $false, $false)

# Remove old module
foreach ($c in $wb.VBProject.VBComponents) {
    if ($c.Name -eq "DataModule") {
        $wb.VBProject.VBComponents.Remove($c)
        Write-Host "Module removed"
    }
}

# Import
$wb.VBProject.VBComponents.Import((Get-Location).Path + "\DataModule.bas")
Write-Host "Module imported"

$excel.Visible = $true
Start-Sleep -Seconds 1

# Run
$excel.Run("DataModule.ConvertData")
Write-Host "Macro executed"

$wb.Save()
$wb.Close($false)
$excel.Quit()
Write-Host "Done"
