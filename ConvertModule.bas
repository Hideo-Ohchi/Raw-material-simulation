Attribute VB_Name = "ConvertModule"
Option Explicit

Public Sub ConvertToHorizontalFormat()
    Dim wsIn As Worksheet
    Dim lastRow As Long
    Dim dataStart As Long
    Dim i As Long
    
    On Error GoTo ErrH
    
    Set wsIn = ThisWorkbook.Sheets("シミュレーションイン")
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
    
    ' Find where data starts (first row with numeric item code in column A)
    dataStart = 2
    For i = 2 To lastRow
        If wsIn.Cells(i, 1).Value <> "" Then
            dataStart = i
            Exit For
        End If
    Next i
    
    ' Collect unique months from column C (current format)
    Dim monthList As Collection
    Dim itemList As Collection
    Dim dataMap As Object
    Dim r As Long, c As Long
    Dim month As String
    Dim itemCode As String
    Dim quantity As Long
    Dim foundMonth As Boolean
    
    Set monthList = New Collection
    Set itemList = New Collection
    Set dataMap = CreateObject("Scripting.Dictionary")
    
    ' Scan data to collect months and items
    For r = dataStart To lastRow
        itemCode = CStr(wsIn.Cells(r, 1).Value)
        month = CStr(wsIn.Cells(r, 3).Value)
        quantity = wsIn.Cells(r, 4).Value
        
        If itemCode <> "" And month <> "" Then
            ' Add month if new
            foundMonth = False
            On Error Resume Next
            monthList(CStr(month))
            If Err.Number <> 0 Then
                Err.Clear
                foundMonth = False
            Else
                foundMonth = True
            End If
            On Error GoTo ErrH
            
            If Not foundMonth Then
                monthList.Add month, CStr(month)
            End If
            
            ' Add item if new
            If dataMap.Exists(itemCode) = False Then
                itemList.Add itemCode
                dataMap.Add itemCode, CreateObject("Scripting.Dictionary")
            End If
            
            ' Store quantity
            dataMap(itemCode)(month) = quantity
        End If
    Next r
    
    ' Write new format
    wsIn.Cells.Clear
    
    ' Headers
    wsIn.Cells(1, 1).Value = "品目コード"
    wsIn.Cells(1, 2).Value = "品目名"
    For i = 1 To monthList.Count
        wsIn.Cells(1, 2 + i).Value = monthList(i)
    Next i
    
    ' Data rows
    Dim itemName As String
    For i = 1 To itemList.Count
        itemCode = itemList(i)
        wsIn.Cells(1 + i, 1).Value = itemCode
        
        ' Get item name from original data
        For r = dataStart To lastRow
            If CStr(wsIn.Cells(r, 1).Value) = itemCode Then
                itemName = wsIn.Cells(r, 2).Value
                Exit For
            End If
        Next r
        
        wsIn.Cells(1 + i, 2).Value = itemName
        
        ' Write quantities
        For c = 1 To monthList.Count
            If dataMap(itemCode).Exists(monthList(c)) Then
                wsIn.Cells(1 + i, 2 + c).Value = dataMap(itemCode)(monthList(c))
            End If
        Next c
    Next i
    
    MsgBox "完了: " & itemList.Count & " 品目, " & monthList.Count & " ヶ月", vbInformation
    
    Exit Sub
ErrH:
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub
