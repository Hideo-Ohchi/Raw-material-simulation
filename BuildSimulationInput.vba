Attribute VB_Name = "Module1"
Sub BuildSimulationInput()
    Dim wsConfig As Worksheet
    Dim wsPlan As Worksheet
    Dim wsOut As Worksheet
    Dim wsMaster As Worksheet
    On Error GoTo ErrorHandler
    Set wsConfig = ThisWorkbook.Sheets(16)
    Set wsPlan = ThisWorkbook.Sheets(2)
    Set wsOut = ThisWorkbook.Sheets(17)
    Set wsMaster = ThisWorkbook.Sheets(6)
    Application.ScreenUpdating = False
    Dim startYM As Long
    Dim monthCount As Long
    Dim b1Value As Variant
    Dim b2Value As Variant
    b1Value = wsConfig.Range("B1").Value
    b2Value = wsConfig.Range("B2").Value
    If IsEmpty(b1Value) Or IsNull(b1Value) Or b1Value = "" Then
        MsgBox "Error: B1 is empty. Set start month (YYYYMM).", vbCritical, "Config Error"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    If IsEmpty(b2Value) Or IsNull(b2Value) Or b2Value = "" Then
        MsgBox "Error: B2 is empty. Set month count.", vbCritical, "Config Error"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    startYM = CLng(b1Value)
    monthCount = CLng(b2Value)
    If startYM < 200001 Or startYM > 210012 Then
        MsgBox "Error: B1 value invalid: " & startYM & vbCrLf & "Expected YYYYMM format like 202604.", vbCritical, "Config Error"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    If monthCount < 1 Or monthCount > 24 Then
        MsgBox "Error: B2 value invalid: " & monthCount & vbCrLf & "Expected month count like 4.", vbCritical, "Config Error"
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
    Dim endYM As Long
    endYM = AddMonths(startYM, monthCount)
    For r = 2 To planLastRow
        itemCode = CStr(wsPlan.Cells(r, 1).Value)
        itemCode = Trim(itemCode)
        If Len(itemCode) > 0 And Not IsError(wsPlan.Cells(r, 2).Value) Then
            On Error Resume Next
            ym = CLng(wsPlan.Cells(r, 2).Value)
            On Error GoTo ErrorHandler
            If ym >= startYM And ym < endYM Then
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
    Dim m As Long
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
    wsOut.Cells.Clear
    wsOut.Cells(1, 1).Value = "ItemCode"
    wsOut.Cells(1, 2).Value = "ItemName"
    For m = 0 To monthCount - 1
        wsOut.Cells(1, 3 + m).Value = CStr(months(m))
    Next m
    Dim outRow As Long
    outRow = 2
    If planDict.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "No data found in range " & startYM & " to " & AddMonths(startYM, monthCount - 1), vbInformation, "No Data"
        Exit Sub
    End If
    Dim sortedKeys() As String
    ReDim sortedKeys(0 To planDict.Count - 1)
    Dim keyIdx As Long
    Dim key As Variant
    keyIdx = 0
    For Each key In planDict.Keys
        sortedKeys(keyIdx) = CStr(key)
        keyIdx = keyIdx + 1
    Next key
    If planDict.Count > 1 Then
        Call QuickSortString(sortedKeys, 0, planDict.Count - 1)
    End If
    Dim sortIdx As Long
    Dim itemKey As String
    Dim monthYM As Long
    For sortIdx = 0 To planDict.Count - 1
        itemKey = sortedKeys(sortIdx)
        wsOut.Cells(outRow, 1).Value = itemKey
        If nameMap.Exists(itemKey) Then
            wsOut.Cells(outRow, 2).Value = nameMap(itemKey)
        Else
            wsOut.Cells(outRow, 2).Value = ""
        End If
        For m = 0 To monthCount - 1
            monthYM = months(m)
            If planDict(itemKey).Exists(CStr(monthYM)) Then
                wsOut.Cells(outRow, 3 + m).Value = planDict(itemKey)(CStr(monthYM))
            Else
                wsOut.Cells(outRow, 3 + m).Value = 0
            End If
        Next m
        outRow = outRow + 1
    Next sortIdx
    Application.ScreenUpdating = True
    MsgBox "OK: " & planDict.Count & " items, " & monthCount & " months. Start:" & startYM, vbInformation, "Success"
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "B1=" & b1Value & " B2=" & b2Value, vbCritical, "Error"
End Sub

Function AddMonths(baseYM As Long, monthsToAdd As Long) As Long
    Dim yr As Long
    Dim mo As Long
    Dim newMo As Long
    Dim newYr As Long
    yr = baseYM \ 100
    mo = baseYM Mod 100
    newMo = mo + monthsToAdd
    newYr = yr
    Do While newMo > 12
        newMo = newMo - 12
        newYr = newYr + 1
    Loop
    AddMonths = newYr * 100 + newMo
End Function

Sub QuickSortString(arr() As String, left As Long, right As Long)
    Dim pivot As String
    Dim temp As String
    Dim i As Long
    Dim j As Long
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