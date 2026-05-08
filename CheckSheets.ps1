$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$book = $excel.Workbooks.Open("$(Get-Location)\BOM copy 1.xlsm")
Write-Host "Sheet names:"
foreach ($sheet in $book.Sheets) {
    Write-Host " - $($sheet.Name)"
}
$book.Close($false)
$excel.Quit()
