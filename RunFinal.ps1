$ErrorActionPreference = "Continue"

Write-Host "Clearing Excel processes..."
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

Write-Host "Opening workbook..."
$excelFile = "$(Get-Location)\BOM copy 1.xlsm"
$modulePath = "$(Get-Location)\SimulationModule.bas"

Write-Host "File: $excelFile"
Write-Host "Module: $modulePath"

$excel = New-Object -ComObject Excel.Application
$excel.AutomationSecurity = 1
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open($excelFile, $false, $false)
Write-Host "Workbook opened"

# Remove old modules
foreach ($c in $wb.VBProject.VBComponents) {
    if ($c.Name -eq "SimulationModule" -or $c.Name -eq "シミュレーション") {
        $wb.VBProject.VBComponents.Remove($c)
        Write-Host "Removed: $($c.Name)"
    }
}

# Import new module
Write-Host "Importing module..."
$wb.VBProject.VBComponents.Import($modulePath)
Write-Host "Module imported successfully"

$excel.Visible = $true

# Run macro with error handling
Write-Host "Executing macro..."
$result = $excel.Run("SimulationModule.ConvertToHorizontalFormatV2")
Write-Host "Macro result: $result"

Start-Sleep -Seconds 2
$wb.Save()
Write-Host "Saved"

$wb.Close($false)
$excel.Quit()
Write-Host "Done"
