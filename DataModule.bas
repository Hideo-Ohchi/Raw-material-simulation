Attribute VB_Name = "DataModule"
Option Explicit

Public Sub ConvertData()
    Dim wsIn As Worksheet
    Dim wsOut As Worksheet
    
    On Error GoTo Err1
    
    Set wsIn = ThisWorkbook.Sheets("シミュレーションイン")
    Set wsOut = ThisWorkbook.Sheets("シミュレーションアウト")
    
    Dim lastRow As Long
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
    
    MsgBox "Rows: " & lastRow, vbInformation
    
    Exit Sub
Err1:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
