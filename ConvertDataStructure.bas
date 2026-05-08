Attribute VB_Name = "データ構造変更"

Option Explicit

''' 現在のデータ構造を新しい構造に変換
''' 現在：A=品目コード, B=品目名, C=対象年月, D=元生産量, E=シミュレーション生産量
''' 新規：A=品目コード, B=品目名, C以降=月別生産量

Public Sub ConvertSimulationInputFormat()
    
    Dim wsIn As Worksheet
    Dim wsTemp As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim itemCode As String, itemName As String, yearMonth As String
    Dim simQty As Double
    Dim dictKey As String
    Dim dict As Object
    Dim colIndex As Long
    
    On Error GoTo ErrorHandler
    
    Set wsIn = ThisWorkbook.Sheets("シミュレーション入力")
    
    ' Get current data dimensions
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then
        MsgBox "No data to convert", vbExclamation
        Exit Sub
    End If
    
    ' Create temporary sheet for conversion
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("_Temp").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    Set wsTemp = ThisWorkbook.Sheets.Add(, ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTemp.Name = "_Temp"
    
    ' Dictionary to store unique months and items
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' First pass: collect all items and months
    For i = 2 To lastRow
        itemCode = wsIn.Cells(i, 1).Value
        yearMonth = wsIn.Cells(i, 3).Value
        
        If itemCode <> "" And yearMonth <> "" Then
            dictKey = itemCode
            If Not dict.exists(dictKey) Then
                dict.Add dictKey, itemCode
            End If
        End If
    Next i
    
    ' Build new headers starting from month columns
    wsTemp.Cells(1, 1).Value = "品目コード"
    wsTemp.Cells(1, 2).Value = "品目名"
    
    colIndex = 3
    Dim monthList As Collection
    Set monthList = New Collection
    Dim monthAdded As Boolean
    
    For i = 2 To lastRow
        yearMonth = wsIn.Cells(i, 3).Value
        If yearMonth <> "" Then
            monthAdded = False
            On Error Resume Next
            For j = 1 To monthList.Count
                If monthList(j) = yearMonth Then
                    monthAdded = True
                    Exit For
                End If
            Next j
            On Error GoTo ErrorHandler
            
            If Not monthAdded Then
                monthList.Add yearMonth
                wsTemp.Cells(1, colIndex).Value = yearMonth
                colIndex = colIndex + 1
            End If
        End If
    Next i
    
    ' Second pass: fill in data
    colIndex = 2
    Dim currentItem As String, tempRow As Long
    currentItem = ""
    tempRow = 2
    
    For i = 2 To lastRow
        itemCode = wsIn.Cells(i, 1).Value
        itemName = wsIn.Cells(i, 2).Value
        yearMonth = wsIn.Cells(i, 3).Value
        simQty = wsIn.Cells(i, 5).Value
        
        If itemCode <> "" Then
            If itemCode <> currentItem Then
                currentItem = itemCode
                tempRow = tempRow + 1
                wsTemp.Cells(tempRow, 1).Value = itemCode
                wsTemp.Cells(tempRow, 2).Value = itemName
            End If
            
            ' Find month column
            For j = 3 To colIndex - 1 + monthList.Count
                If wsTemp.Cells(1, j).Value = yearMonth Then
                    wsTemp.Cells(tempRow, j).Value = simQty
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' Copy converted data back to original sheet
    wsIn.Cells.Clear
    lastRow = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
    lastCol = wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column
    
    wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(lastRow, lastCol)).Copy
    wsIn.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Delete temp sheet
    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True
    
    MsgBox "Data structure converted successfully", vbInformation
    
    Exit Sub
ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

''' シミュレーション出力も同じ構造に変換
Public Sub ConvertSimulationOutputFormat()
    
    Dim wsOut As Worksheet
    Dim lastRow As Long, lastCol As Long
    
    On Error GoTo ErrorHandler
    
    Set wsOut = ThisWorkbook.Sheets("シミュレーション出力")
    
    lastRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    lastCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
    
    ' Verify current structure and convert if needed
    If wsOut.Cells(1, 1).Value <> "品目コード" Then
        MsgBox "Output sheet header not found", vbExclamation
        Exit Sub
    End If
    
    MsgBox "Output sheet structure is already correct", vbInformation
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub