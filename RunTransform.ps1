Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$excel = New-Object -ComObject Excel.Application
$excel.AutomationSecurity = 1
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open((Get-Location).Path + "\BOM copy 1.xlsm", $false, $false)

# Remove old transform module
foreach ($c in $wb.VBProject.VBComponents) {
    if ($c.Name -eq "TransformModule") {
        $wb.VBProject.VBComponents.Remove($c)
        Write-Host "Old module removed"
    }
}

# Import
$wb.VBProject.VBComponents.Import((Get-Location).Path + "\TransformModule.bas")
Write-Host "Module imported"

$excel.Visible = $true
Start-Sleep -Seconds 1

# Execute
$excel.Run("TransformModule.TransformToHorizontal")
Write-Host "Transformation executed"

Start-Sleep -Seconds 2
$wb.Save()
Write-Host "File saved"

$wb.Close($false)
$excel.Quit()
Write-Host "Complete"
