Attribute VB_Name = "SimulationModule"
Option Explicit

Public Sub ConvertToHorizontalFormatV2()
    Dim wsIn As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long
    Dim dataStart As Long
    Dim r As Long, c As Long
    Dim srcData() As Variant
    Dim monthList() As String
    Dim itemCodes() As String
    Dim itemNames() As String
    Dim monthCount As Long
    Dim itemCount As Long
    Dim tmpMonth As String
    Dim tmpItem As String
    Dim i As Long, j As Long
    
    On Error GoTo ErrH
    
    Set wsIn = ThisWorkbook.Sheets("シミュレーションイン")
    Set wsOut = ThisWorkbook.Sheets("シミュレーションアウチE)
    
    ' Find data start row (first row with 10+ digit item code)
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
    dataStart = 0
    For r = 1 To lastRow
        tmpItem = CStr(wsIn.Cells(r, 1).Value)
        If tmpItem <> "" And Len(tmpItem) >= 10 Then
            On Error Resume Next
            If CLng(tmpItem) > 0 Then
                dataStart = r
                Exit For
            End If
            On Error GoTo ErrH
        End If
    Next r
    
    If dataStart = 0 Then
        MsgBox "チE?Eタが見つかりません", vbExclamation
        Exit Sub
    End If
    
    ' Load source data
    ReDim srcData(1 To lastRow - dataStart + 1, 1 To 5)
    For r = dataStart To lastRow
        For c = 1 To 5
            srcData(r - dataStart + 1, c) = wsIn.Cells(r, c).Value
        Next c
    Next r
    
    ' Collect unique months and items
    ReDim monthList(0)
    ReDim itemCodes(0)
    ReDim itemNames(0)
    monthCount = 0
    itemCount = 0
    
    For i = 1 To UBound(srcData, 1)
        tmpMonth = CStr(srcData(i, 3))
        tmpItem = CStr(srcData(i, 1))
        
        ' Add month if new
        Dim found As Boolean
        found = False
        For j = 1 To monthCount
            If monthList(j) = tmpMonth Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            monthCount = monthCount + 1
            ReDim Preserve monthList(monthCount)
            monthList(monthCount) = tmpMonth
        End If
        
        ' Add item if new
        found = False
        For j = 1 To itemCount
            If itemCodes(j) = tmpItem Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            itemCount = itemCount + 1
            ReDim Preserve itemCodes(itemCount)
            ReDim Preserve itemNames(itemCount)
            itemCodes(itemCount) = tmpItem
            itemNames(itemCount) = srcData(i, 2)
        End If
    Next i
    
    ' Sort months ascending
    For i = 1 To monthCount - 1
        For j = i + 1 To monthCount
            If monthList(i) > monthList(j) Then
                tmpMonth = monthList(i)
                monthList(i) = monthList(j)
                monthList(j) = tmpMonth
            End If
        Next j
    Next i
    
    ' Clear and rebuild シミュレーションイン
    wsIn.Cells.Clear
    
    ' Write headers
    wsIn.Cells(1, 1).Value = "品目コーチE
    wsIn.Cells(1, 2).Value = "品目吁E
    For i = 1 To monthCount
        wsIn.Cells(1, 2 + i).Value = monthList(i)
    Next i
    
    ' Write data rows
    For i = 1 To itemCount
        wsIn.Cells(1 + i, 1).Value = itemCodes(i)
        wsIn.Cells(1 + i, 2).Value = itemNames(i)
        
        For j = 1 To monthCount
            Dim qty As Long
            qty = 0
            ' Find quantity for this item and month
            For r = 1 To UBound(srcData, 1)
                If srcData(r, 1) = itemCodes(i) And srcData(r, 3) = monthList(j) Then
                    qty = srcData(r, 4)
                    Exit For
                End If
            Next r
            wsIn.Cells(1 + i, 2 + j).Value = qty
        Next j
    Next i
    
    ' Update シミュレーションアウチEheaders
    wsOut.Cells(1, 1).Value = "品目コーチE
    wsOut.Cells(1, 2).Value = "品目吁E
    For i = 1 To monthCount
        wsOut.Cells(1, 2 + i).Value = monthList(i)
    Next i
    
    MsgBox "変換完?E " & itemCount & " 品目 x " & monthCount & " ヶ?E, vbInformation
    
    Exit Sub
ErrH:
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Public Sub InitializeFromProductionPlan()
    MsgBox "実裁E???E, vbInformation
End Sub

Public Sub ExecuteSimulation()
    MsgBox "シミュレーション実?E, vbInformation
End Sub
