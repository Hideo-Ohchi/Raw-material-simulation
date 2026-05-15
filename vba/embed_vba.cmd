@'
# VBA Installation Script
$ExcelFile = "$PSScriptRoot\BOM copy 1.xlsm"
$VBAFile = "$PSScriptRoot\BuildSimulationInput.vba"

# Stop Excel
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# Create Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Open workbook
$wb = $excel.Workbooks.Open($ExcelFile, $false, $false)

# Get VBA project
$vbProject = $wb.VBProject

# Read VBA code from file
$vbaCode = Get-Content -Path $VBAFile -Raw -Encoding Default

# Check and remove existing module
foreach ($comp in $vbProject.VBComponents) {
    if ($comp.Name -eq "BuildSimulationInput") {
        $vbProject.VBComponents.Remove($comp)
        break
    }
}

# Add new module (1 = StandardModule)
$newModule = $vbProject.VBComponents.Add(1)
$newModule.Name = "BuildSimulationInput"

# Add VBA code
$newModule.CodeModule.AddFromString($vbaCode)

# Save workbook
$wb.Save()

# Close
$wb.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()

Write-Host "OK: VBA embedded successfully"
'@ | Set-Content -Path "c:\Users\119351\Desktop\原料シミュレーション\embed_vba_simple.ps1" -Encoding Default
