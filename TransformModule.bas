Attribute VB_Name = "TransformModule"
Option Explicit

Public Sub TransformToHorizontal()
    Dim wsIn As Worksheet
    Dim lastRow As Long
    Dim r As Long, c As Long
    Dim dataStart As Long
    Dim sourceData() As Variant
    Dim newData() As Variant
    Dim newRow As Long
    Dim monthStr As String
    Dim itemCode As String
    Dim itemName As String
    Dim qty As Long
    Dim i As Long, j As Long, k As Long
    Dim monthArray() As String
    Dim itemArray() As String
    Dim monthIdx As Long
    Dim itemIdx As Long
    Dim found As Boolean
    Dim srcRow As Long
    
    On Error GoTo ErrH
    
    Set wsIn = ThisWorkbook.Sheets("シミュレーションイン")
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(-4162).Row
    
    ' Find data start (Row 3 based on the format)
    dataStart = 3
    
    ' Load all source data
    ReDim sourceData(1 To lastRow - dataStart + 1, 1 To 5)
    For r = dataStart To lastRow
        For c = 1 To 5
            sourceData(r - dataStart + 1, c) = wsIn.Cells(r, c).Value2
        Next c
    Next r
    
    ' Collect unique months and items
    ReDim monthArray(1 To 100)
    ReDim itemArray(1 To 500)
    monthIdx = 0
    itemIdx = 0
    
    For i = 1 To UBound(sourceData, 1)
        monthStr = CStr(sourceData(i, 3))
        itemCode = CStr(sourceData(i, 1))
        
        ' Check if month exists
        found = False
        For j = 1 To monthIdx
            If monthArray(j) = monthStr Then
                found = True
                Exit For
            End If
        Next j
        If Not found And monthStr <> "" Then
            monthIdx = monthIdx + 1
            monthArray(monthIdx) = monthStr
        End If
        
        ' Check if item exists
        found = False
        For j = 1 To itemIdx
            If itemArray(j) = itemCode Then
                found = True
                Exit For
            End If
        Next j
        If Not found And itemCode <> "" Then
            itemIdx = itemIdx + 1
            itemArray(itemIdx) = itemCode
        End If
    Next i
    
    ' Sort months ascending (bubble sort)
    For i = 1 To monthIdx - 1
        For j = i + 1 To monthIdx
            If monthArray(i) > monthArray(j) Then
                monthStr = monthArray(i)
                monthArray(i) = monthArray(j)
                monthArray(j) = monthStr
            End If
        Next j
    Next i
    
    ' Build new data array
    ReDim newData(1 To itemIdx + 1, 1 To monthIdx + 2)
    
    ' Headers
    newData(1, 1) = "品目コード"
    newData(1, 2) = "品目名"
    For i = 1 To monthIdx
        newData(1, 2 + i) = monthArray(i)
    Next i
    
    ' Data rows
    For i = 1 To itemIdx
        itemCode = itemArray(i)
        newData(i + 1, 1) = itemCode
        
        ' Find item name
        For srcRow = 1 To UBound(sourceData, 1)
            If CStr(sourceData(srcRow, 1)) = itemCode Then
                newData(i + 1, 2) = sourceData(srcRow, 2)
                Exit For
            End If
        Next srcRow
        
        ' Fill quantities for each month
        For j = 1 To monthIdx
            monthStr = monthArray(j)
            qty = 0
            
            For srcRow = 1 To UBound(sourceData, 1)
                If CStr(sourceData(srcRow, 1)) = itemCode And CStr(sourceData(srcRow, 3)) = monthStr Then
                    qty = sourceData(srcRow, 4)
                    Exit For
                End If
            Next srcRow
            
            newData(i + 1, 2 + j) = qty
        Next j
    Next i
    
    ' Clear and write new format
    wsIn.Cells.Clear
    wsIn.Range("A1").Resize(itemIdx + 1, monthIdx + 2).Value2 = newData
    
    MsgBox "変換完了: " & itemIdx & " 品目 x " & monthIdx & " ヶ月", vbInformation
    
    Exit Sub
ErrH:
    MsgBox "エラー: " & Err.Description & " (" & Err.Number & ")", vbCritical
End Sub
