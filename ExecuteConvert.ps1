Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$excel = New-Object -ComObject Excel.Application
$excel.AutomationSecurity = 1
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open((Get-Location).Path + "\BOM copy 1.xlsm", $false, $false)

# Remove old module
foreach ($c in $wb.VBProject.VBComponents) {
    if ($c.Name -eq "ConvertModule") {
        $wb.VBProject.VBComponents.Remove($c)
    }
}

# Import conversion module
$wb.VBProject.VBComponents.Import((Get-Location).Path + "\ConvertModule.bas")
Write-Host "Module imported"

$excel.Visible = $true
Start-Sleep -Seconds 1

# Execute conversion
$result = $excel.Run("ConvertModule.ConvertToHorizontalFormat")
Write-Host "Conversion executed"

Start-Sleep -Seconds 2
$wb.Save()
Write-Host "Saved"

$wb.Close($false)
$excel.Quit()
Write-Host "Complete"
