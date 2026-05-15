Sub BuildSimulationInput()
    Dim wsConfig As Worksheet
    Dim wsPlan As Worksheet
    Dim wsIn As Worksheet
    Dim wsMaster As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set wsConfig = ThisWorkbook.Sheets(1)
    Set wsPlan = ThisWorkbook.Sheets(2)
    Set wsIn = ThisWorkbook.Sheets(16)
    Set wsMaster = ThisWorkbook.Sheets(6)
    
    Application.ScreenUpdating = False
    
    Dim startYM As Long
    Dim monthCount As Integer
    Dim b1Value As Variant
    Dim b2Value As Variant
    
    b1Value = wsConfig.Range("B1").Value
    b2Value = wsConfig.Range("B2").Value
    
    On Error Resume Next
    startYM = CLng(b1Value)
    monthCount = CInt(b2Value)
    On Error GoTo ErrorHandler
    
    If IsEmpty(startYM) Or startYM <= 0 Or monthCount <= 0 Then
        MsgBox "Error: B1 or B2 is not set correctly." & vbCrLf & _
               "B1 (Start Month): " & b1Value & vbCrLf & _
               "B2 (Month Count): " & b2Value, vbCritical, "Config Error"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim planDict As Object
    Set planDict = CreateObject("Scripting.Dictionary")
    
    Dim planLastRow As Long
    planLastRow = wsPlan.Cells(wsPlan.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    Dim itemCode As String
    Dim ym As Long
    Dim qty As Double
    
    For r = 2 To planLastRow
        itemCode = CStr(wsPlan.Cells(r, 1).Value)
        itemCode = Trim(itemCode)
        
        If Len(itemCode) > 0 And Not IsError(wsPlan.Cells(r, 2).Value) Then
            ym = CLng(wsPlan.Cells(r, 2).Value)
            
            If ym >= startYM And ym < AddMonths(startYM, monthCount) Then
                qty = 0
                On Error Resume Next
                qty = CDbl(wsPlan.Cells(r, 3).Value)
                On Error GoTo ErrorHandler
                
                If Not planDict.Exists(itemCode) Then
                    planDict.Add itemCode, CreateObject("Scripting.Dictionary")
                End If
                
                planDict(itemCode)(CStr(ym)) = qty
            End If
        End If
    Next r
    
    Dim months() As Long
    ReDim months(0 To monthCount - 1)
    Dim m As Integer
    
    For m = 0 To monthCount - 1
        months(m) = AddMonths(startYM, m)
    Next m
    
    Dim nameMap As Object
    Set nameMap = CreateObject("Scripting.Dictionary")
    
    Dim masterLast As Long
    masterLast = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    Dim mCode As String
    Dim mName As String
    
    For i = 2 To masterLast
        mCode = CStr(wsMaster.Cells(i, 1).Value)
        mCode = Trim(mCode)
        
        If Len(mCode) > 0 Then
            mName = CStr(wsMaster.Cells(i, 2).Value)
            If Not nameMap.Exists(mCode) Then
                nameMap.Add mCode, mName
            End If
        End If
    Next i
    
    wsIn.Cells.Clear()
    
    wsIn.Cells(1, 1).Value = "ItemCode"
    wsIn.Cells(1, 2).Value = "ItemName"
    
    For m = 0 To monthCount - 1
        wsIn.Cells(1, 3 + m).Value = CStr(months(m))
    Next m
    
    Dim outRow As Long
    outRow = 2
    
    Dim sortedKeys() As String
    ReDim sortedKeys(0 To planDict.Count - 1)
    
    Dim keyIdx As Integer
    Dim key As Variant
    keyIdx = 0
    
    For Each key In planDict.Keys
        sortedKeys(keyIdx) = CStr(key)
        keyIdx = keyIdx + 1
    Next key
    
    If planDict.Count > 1 Then
        Call QuickSortString(sortedKeys, LBound(sortedKeys), UBound(sortedKeys))
    End If
    
    Dim sortIdx As Integer
    Dim itemKey As String
    Dim monthYM As Long
    
    For sortIdx = LBound(sortedKeys) To UBound(sortedKeys)
        itemKey = sortedKeys(sortIdx)
        
        wsIn.Cells(outRow, 1).Value = itemKey
        
        If nameMap.Exists(itemKey) Then
            wsIn.Cells(outRow, 2).Value = nameMap(itemKey)
        Else
            wsIn.Cells(outRow, 2).Value = ""
        End If
        
        For m = 0 To monthCount - 1
            monthYM = months(m)
            
            If planDict(itemKey).Exists(CStr(monthYM)) Then
                wsIn.Cells(outRow, 3 + m).Value = planDict(itemKey)(CStr(monthYM))
            Else
                wsIn.Cells(outRow, 3 + m).Value = 0
            End If
        Next m
        
        outRow = outRow + 1
    Next sortIdx
    
    Application.ScreenUpdating = True
    
    MsgBox "OK: Simulation input rebuilt successfully." & vbCrLf & _
           "Start Month: " & startYM & vbCrLf & _
           "Month Count: " & monthCount & vbCrLf & _
           "Items: " & planDict.Count, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Function AddMonths(baseYM As Long, monthsToAdd As Integer) As Long
    Dim year As Integer
    Dim month As Integer
    Dim newMonth As Integer
    Dim newYear As Integer
    
    year = baseYM \ 100
    month = baseYM Mod 100
    
    newMonth = month + monthsToAdd
    newYear = year
    
    Do While newMonth > 12
        newMonth = newMonth - 12
        newYear = newYear + 1
    Loop
    
    AddMonths = newYear * 100 + newMonth
End Function

Sub QuickSortString(arr() As String, left As Integer, right As Integer)
    Dim pivot As String
    Dim temp As String
    Dim i As Integer
    Dim j As Integer
    
    If left < right Then
        pivot = arr((left + right) \ 2)
        i = left - 1
        j = right + 1
        
        Do
            Do
                i = i + 1
            Loop While arr(i) < pivot
            
            Do
                j = j - 1
            Loop While arr(j) > pivot
            
            If i < j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Loop While i < j
        
        Call QuickSortString(arr, left, j)
        Call QuickSortString(arr, j + 1, right)
    End If
End Sub
