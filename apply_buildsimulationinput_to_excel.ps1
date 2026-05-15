$ErrorActionPreference = 'Stop'

$wbPath = (Get-ChildItem -File -Filter *.xlsm |
    Where-Object { $_.Name -notlike '~$*' } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1 -ExpandProperty FullName)
$vbaPath = Join-Path (Get-Location) 'BuildSimulationInput.vba'

if ([string]::IsNullOrWhiteSpace($wbPath) -or -not (Test-Path $wbPath)) { throw 'Workbook not found (*.xlsm).' }
if (-not (Test-Path $vbaPath)) { throw "VBA file not found: $vbaPath" }

$excel = $null
$createdExcel = $false
$openedHere = $false
$wb = $null

try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    }
    catch {
        $excel = New-Object -ComObject Excel.Application
        $createdExcel = $true
    }

    $excel.DisplayAlerts = $false

    foreach ($w in $excel.Workbooks) {
        if ($w.FullName -ieq $wbPath) {
            $wb = $w
            break
        }
    }

    if ($null -eq $wb) {
        $wb = $excel.Workbooks.Open($wbPath, $false, $false)
        $openedHere = $true
    }

    $vbProject = $wb.VBProject
    $vbaCode = Get-Content -Path $vbaPath -Raw -Encoding UTF8
    $vbaCode = ($vbaCode -split "`r?`n" | Where-Object { $_ -notmatch '^Attribute VB_Name' }) -join "`r`n"

    $targetComp = $null
    foreach ($comp in $vbProject.VBComponents) {
        if ($comp.Type -eq 1) {
            $cm = $comp.CodeModule
            if ($cm.CountOfLines -gt 0 -and $cm.Find('Sub BuildSimulationInput(', 1, 1, $cm.CountOfLines, 255)) {
                $targetComp = $comp
                break
            }
        }
    }

    if ($null -eq $targetComp) {
        $targetComp = $vbProject.VBComponents.Add(1)
        $targetComp.Name = 'Module1'
    }
    else {
        $cm = $targetComp.CodeModule
        if ($cm.CountOfLines -gt 0) {
            $cm.DeleteLines(1, $cm.CountOfLines)
        }
    }

    $targetComp.CodeModule.AddFromString($vbaCode)
    $wb.Save() | Out-Null

    Write-Host 'OK: BuildSimulationInput was applied to 原料シミュレーション.xlsm and saved.' -ForegroundColor Green
}
finally {
    if ($openedHere -and $wb) { $wb.Close($true) | Out-Null }
    if ($createdExcel -and $excel) { $excel.Quit() }
    if ($excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
