' VBScript to convert simulation data structure
' 現在：A=品目コード, B=品目名, C=対象年月, D=元生産量, E=シミュレーション生産量
' 新規：A=品目コード, B=品目名, C以降=月別生産量

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.DisplayAlerts = False

Dim objBook, wsIn, wsTemp
Dim lastRow, i, j, colIndex, tempRow
Dim itemCode, itemName, yearMonth, simQty
Dim monthArray, monthCount
Dim currentItem, foundMonth, k

Set objBook = objExcel.Workbooks.Open("BOM copy 1.xlsm")
Set wsIn = objBook.Sheets("シミュレーション入力")

lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(-4117).Row

If lastRow <= 1 Then
    MsgBox "No data to convert"
    objBook.Close False
    objExcel.Quit
    WScript.Quit
End If

' Create temp sheet
On Error Resume Next
objBook.Sheets("_Temp").Delete
On Error GoTo 0

Set wsTemp = objBook.Sheets.Add()
wsTemp.Name = "_Temp"

' Set headers
wsTemp.Cells(1, 1).Value = "品目コード"
wsTemp.Cells(1, 2).Value = "品目名"

' Collect unique months
Set monthArray = CreateObject("Scripting.Dictionary")
monthCount = 0

For i = 2 To lastRow
    yearMonth = wsIn.Cells(i, 3).Value
    If yearMonth <> "" Then
        If Not monthArray.exists(yearMonth) Then
            monthArray.Add yearMonth, monthCount
            wsTemp.Cells(1, monthCount + 3).Value = yearMonth
            monthCount = monthCount + 1
        End If
    End If
Next i

' Fill data
tempRow = 2
currentItem = ""

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
        
        ' Find month column and put value
        If monthArray.exists(yearMonth) Then
            j = monthArray(yearMonth)
            wsTemp.Cells(tempRow, j + 3).Value = simQty
        End If
    End If
Next i

' Copy back to original sheet
wsIn.Cells.Clear
lastRow = wsTemp.Cells(wsTemp.Rows.Count, 1).End(-4117).Row

wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(lastRow, monthCount + 2)).Copy
wsIn.Range("A1").PasteSpecial

' Delete temp sheet
wsTemp.Delete

' Save and close
objBook.Save
objBook.Close False
objExcel.Quit

MsgBox "Data structure converted successfully"
