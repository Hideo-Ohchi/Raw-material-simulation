# Reset and import new VBA module
$excel = New-Object -ComObject Excel.Application
$excel.AutomationSecurity = 1

# Close previous instance
Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$wb = $excel.Workbooks.Open("$(Get-Location)\BOM copy 1.xlsm", $false, $false)

# Remove old module
foreach ($c in $wb.VBProject.VBComponents) {
    if ($c.Name -eq "SimulationModule" -or $c.Name -eq "シミュレーション") {
        $wb.VBProject.VBComponents.Remove($c)
        Write-Host "Old module removed"
    }
}

# Import new module
$wb.VBProject.VBComponents.Import("$(Get-Location)\SimulationModule.bas")
Write-Host "New module imported"

$excel.Visible = $true

# Run macro
Write-Host "Running ConvertToHorizontalFormatV2..."
try {
    $excel.Run("SimulationModule.ConvertToHorizontalFormatV2")
    Write-Host "Macro completed"
} catch {
    Write-Host "Error: $_"
}

Start-Sleep -Seconds 3
$wb.Save()
$wb.Close($false)
$excel.Quit()
Write-Host "Done"
