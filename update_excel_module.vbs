' VBScript for updating Excel VBA module
' Save as: update_excel_module.vbs

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Dim excelPath, basPath
excelPath = "c:\Users\119351\Desktop\原料シミュレーション\BOM copy 1.xlsm"
basPath = "c:\Users\119351\Desktop\原料シミュレーション\Simulation_Module.bas"

On Error Resume Next

' Open workbook
Set objWorkbook = objExcel.Workbooks.Open(excelPath)
If Err.Number <> 0 Then
    MsgBox "Error opening file: " & excelPath
    objExcel.Quit
    WScript.Quit
End If

' Remove old module
For Each objComponent In objWorkbook.VBProject.VBComponents
    If objComponent.Name = "シミュレーション" Then
        objWorkbook.VBProject.VBComponents.Remove objComponent
        WScript.Echo "Old module removed"
    End If
Next

' Import new module
objWorkbook.VBProject.VBComponents.Import basPath
WScript.Echo "Module imported"

' Save
objWorkbook.Save
WScript.Echo "File saved"

' Close
objWorkbook.Close False
objExcel.Quit

WScript.Echo "Complete"
