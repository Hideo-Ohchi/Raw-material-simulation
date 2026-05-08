Attribute VB_Name = "シミュレーション"

Option Explicit

' =====================================================================
' シミュレーション入力シートを横型形式に変換
' 実行前：A=品目コード, B=品目名, C=年月, D=元生産量, E=シミュレーション生産量
' 実行後：A=品目コード, B=品目名, C以降=月別生産量
' =====================================================================
Public Sub ConvertToHorizontalFormat()

    Dim wsIn As Worksheet
    Dim wsOut As Worksheet
    On Error GoTo ErrH

    Set wsIn  = ThisWorkbook.Sheets("シミュレーションイン")
    Set wsOut = ThisWorkbook.Sheets("シミュレーションアウト")

    ' ---- 1. 元データを読み込み ----
    Dim lastRow As Long
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row

    ' データ開始行を検索（品目コードが数値っぽい行）
    Dim dataStart As Long
    dataStart = 0
    Dim r As Long
    For r = 1 To lastRow
        If IsNumeric(wsIn.Cells(r, 1).Value) And Len(CStr(wsIn.Cells(r, 1).Value)) >= 10 Then
            dataStart = r
            Exit For
        End If
    Next r

    If dataStart = 0 Then
        MsgBox "データが見つかりません", vbExclamation
        Exit Sub
    End If

    ' ---- 2. 月リストと品目リストを収集 ----
    Dim monthList()  As String
    Dim itemCodes()  As String
    Dim itemNames()  As String
    Dim monthCount   As Long
    Dim itemCount    As Long
    monthCount = 0
    itemCount  = 0

    ReDim monthList(0)
    ReDim itemCodes(0)
    ReDim itemNames(0)

    For r = dataStart To lastRow
        Dim code As String
        Dim ymStr As String
        code  = CStr(wsIn.Cells(r, 1).Value)
        ymStr = CStr(CLng(wsIn.Cells(r, 3).Value))

        ' 月リスト
        Dim found As Boolean
        found = False
        Dim m As Long
        For m = 0 To monthCount - 1
            If monthList(m) = ymStr Then found = True: Exit For
        Next m
        If Not found Then
            ReDim Preserve monthList(monthCount)
            monthList(monthCount) = ymStr
            monthCount = monthCount + 1
        End If

        ' 品目リスト
        found = False
        Dim i As Long
        For i = 0 To itemCount - 1
            If itemCodes(i) = code Then found = True: Exit For
        Next i
        If Not found Then
            ReDim Preserve itemCodes(itemCount)
            ReDim Preserve itemNames(itemCount)
            itemCodes(itemCount) = code
            itemNames(itemCount) = CStr(wsIn.Cells(r, 2).Value)
            itemCount = itemCount + 1
        End If
    Next r

    ' 月を昇順ソート（バブルソート）
    Dim tmp As String
    For i = 0 To monthCount - 2
        For m = i + 1 To monthCount - 1
            If monthList(i) > monthList(m) Then
                tmp = monthList(i): monthList(i) = monthList(m): monthList(m) = tmp
            End If
        Next m
    Next i

    ' ---- 3. シミュレーションイン を書き換え ----
    wsIn.Cells.Clear

    ' ヘッダー行（行1）
    wsIn.Cells(1, 1).Value = "品目コード"
    wsIn.Cells(1, 2).Value = "品目名"
    For m = 0 To monthCount - 1
        wsIn.Cells(1, m + 3).Value = CLng(monthList(m))
    Next m

    ' データ行
    For i = 0 To itemCount - 1
        wsIn.Cells(i + 2, 1).Value = itemCodes(i)
        wsIn.Cells(i + 2, 2).Value = itemNames(i)
    Next i

    ' 生産量を元データから転記
    For r = dataStart To lastRow
        code  = CStr(wsIn.Cells(r, 1).Value)  ' ← これは wsIn を参照してしまう
    Next r

    ' 正しくは元データのコピー（バッファに保存）してから転記が必要
    ' 再実行: 元データのバッファを作成する
    wsIn.Cells.Clear

    ' 元データを配列に読み込み直し
    Dim srcRange As Range
    Set srcRange = wsIn.Parent.Sheets("シミュレーションイン").Range( _
        wsIn.Cells(dataStart, 1), _
        wsIn.Cells(lastRow, 5))

    MsgBox "Error in logic - use ConvertToHorizontalFormatV2 instead", vbCritical
    Exit Sub
ErrH:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' =====================================================================
' 正しい変換ロジック（バッファ使用）
' =====================================================================
Public Sub ConvertToHorizontalFormatV2()

    Dim wsIn  As Worksheet
    Dim wsOut As Worksheet
    On Error GoTo ErrH

    Set wsIn  = ThisWorkbook.Sheets("シミュレーションイン")
    Set wsOut = ThisWorkbook.Sheets("シミュレーションアウト")

    ' ---- 1. 最終行を検索 ----
    Dim lastRow As Long
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row

    Dim dataStart As Long
    dataStart = 0
    Dim r As Long
    For r = 1 To lastRow
        If IsNumeric(wsIn.Cells(r, 1).Value) And Len(CStr(wsIn.Cells(r, 1).Value)) >= 10 Then
            dataStart = r
            Exit For
        End If
    Next r

    If dataStart = 0 Then
        MsgBox "品目コードが見つかりません（10桁以上の数値を検索）", vbExclamation
        Exit Sub
    End If

    ' ---- 2. 元データを2次元配列に読み込み ----
    Dim rowCount As Long
    rowCount = lastRow - dataStart + 1
    Dim srcData() As Variant
    ReDim srcData(1 To rowCount, 1 To 5)
    For r = 1 To rowCount
        Dim c As Long
        For c = 1 To 5
            srcData(r, c) = wsIn.Cells(dataStart + r - 1, c).Value
        Next c
    Next r

    ' ---- 3. 月・品目リスト収集 ----
    Dim monthList() As String
    Dim itemCodes()  As String
    Dim itemNames()  As String
    Dim monthCount  As Long
    Dim itemCount   As Long
    monthCount = 0
    itemCount  = 0
    ReDim monthList(0)
    ReDim itemCodes(0)
    ReDim itemNames(0)

    Dim i As Long, m As Long
    Dim code As String, ymStr As String
    Dim found As Boolean

    For r = 1 To rowCount
        code  = CStr(srcData(r, 1))
        ymStr = CStr(CLng(srcData(r, 3)))

        ' 月
        found = False
        For m = 0 To monthCount - 1
            If monthList(m) = ymStr Then found = True: Exit For
        Next m
        If Not found Then
            ReDim Preserve monthList(monthCount)
            monthList(monthCount) = ymStr
            monthCount = monthCount + 1
        End If

        ' 品目
        found = False
        For i = 0 To itemCount - 1
            If itemCodes(i) = code Then found = True: Exit For
        Next i
        If Not found Then
            ReDim Preserve itemCodes(itemCount)
            ReDim Preserve itemNames(itemCount)
            itemCodes(itemCount) = code
            itemNames(itemCount) = CStr(srcData(r, 2))
            itemCount = itemCount + 1
        End If
    Next r

    ' 月を昇順ソート
    Dim tmp As String
    For i = 0 To monthCount - 2
        For m = i + 1 To monthCount - 1
            If monthList(i) > monthList(m) Then
                tmp = monthList(i): monthList(i) = monthList(m): monthList(m) = tmp
            End If
        Next m
    Next i

    ' ---- 4. シミュレーションイン 書き換え ----
    wsIn.Cells.Clear

    ' ヘッダー行
    wsIn.Cells(1, 1).Value = "品目コード"
    wsIn.Cells(1, 2).Value = "品目名"
    For m = 0 To monthCount - 1
        wsIn.Cells(1, m + 3).Value = CLng(monthList(m))
    Next m

    ' データ行
    For i = 0 To itemCount - 1
        wsIn.Cells(i + 2, 1).Value = itemCodes(i)
        wsIn.Cells(i + 2, 2).Value = itemNames(i)

        ' 各月の生産量を埋める
        For m = 0 To monthCount - 1
            For r = 1 To rowCount
                If CStr(srcData(r, 1)) = itemCodes(i) Then
                    If CStr(CLng(srcData(r, 3))) = monthList(m) Then
                        wsIn.Cells(i + 2, m + 3).Value = srcData(r, 5)  ' E列 = シミュレーション生産量
                        Exit For
                    End If
                End If
            Next r
        Next m
    Next i

    ' ---- 5. シミュレーションアウト ヘッダー更新 ----
    wsOut.Cells.Clear
    wsOut.Cells(1, 1).Value = "品目コード"
    wsOut.Cells(1, 2).Value = "品目名"
    For m = 0 To monthCount - 1
        wsOut.Cells(1, m + 3).Value = CLng(monthList(m))
    Next m

    MsgBox "変換完了: " & itemCount & " 品目 × " & monthCount & " ヶ月", vbInformation
    Exit Sub
ErrH:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub ExecuteSimulation()

    Dim wsIn  As Worksheet
    Dim wsOut As Worksheet
    On Error GoTo ErrH

    Set wsIn  = ThisWorkbook.Sheets("シミュレーションイン")
    Set wsOut = ThisWorkbook.Sheets("シミュレーションアウト")

    If Not ValidateSimulationInput(wsIn) Then
        MsgBox "Check simulation input data", vbExclamation
        Exit Sub
    End If

    ' ヘッダー行から月列を取得
    Dim lastCol As Long
    lastCol = wsIn.Cells(1, wsIn.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long
    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row

    ' シミュレーションアウトにデータをコピー
    wsOut.Cells.Clear
    wsIn.Range(wsIn.Cells(1, 1), wsIn.Cells(1, lastCol)).Copy wsOut.Range("A1")

    Dim i As Long
    Dim outRow As Long
    outRow = 2

    For i = 2 To lastRow
        If wsIn.Cells(i, 1).Value <> "" Then
            wsOut.Cells(outRow, 1).Value = wsIn.Cells(i, 1).Value
            wsOut.Cells(outRow, 2).Value = wsIn.Cells(i, 2).Value

            Dim m As Long
            For m = 3 To lastCol
                wsOut.Cells(outRow, m).Value = wsIn.Cells(i, m).Value
            Next m

            outRow = outRow + 1
        End If
    Next i

    MsgBox "Simulation complete. Check output sheet.", vbInformation
    Exit Sub
ErrH:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function ValidateSimulationInput(ws As Worksheet) As Boolean
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow <= 1 Then
        ValidateSimulationInput = False
        Exit Function
    End If

    If ws.Cells(1, 1).Value <> "品目コード" Then
        ValidateSimulationInput = False
        Exit Function
    End If

    If ws.Cells(1, 2).Value <> "品目名" Then
        ValidateSimulationInput = False
        Exit Function
    End If

    If ws.Cells(1, 3).Value = "" Then
        ValidateSimulationInput = False
        Exit Function
    End If

    ValidateSimulationInput = True
End Function

Public Sub InitializeFromProductionPlan()
    On Error GoTo ErrH

    Dim wsProd As Worksheet
    Dim wsIn   As Worksheet
    Set wsProd = ThisWorkbook.Sheets("生産計画")
    Set wsIn   = ThisWorkbook.Sheets("シミュレーションイン")

    ' 生産計画から横型で読み込む
    ' 生産計画: A=品目コード, B=対象年月, C=計画数量

    Dim prodLastRow As Long
    prodLastRow = wsProd.Cells(wsProd.Rows.Count, 1).End(xlUp).Row

    ' 月・品目収集
    Dim monthList() As String
    Dim itemCodes()  As String
    Dim itemNames()  As String
    Dim monthCount  As Long
    Dim itemCount   As Long
    monthCount = 0
    itemCount  = 0
    ReDim monthList(0)
    ReDim itemCodes(0)
    ReDim itemNames(0)

    Dim r As Long, i As Long, m As Long
    Dim code As String, ymStr As String
    Dim found As Boolean

    For r = 2 To prodLastRow
        code  = CStr(wsProd.Cells(r, 1).Value)
        ymStr = CStr(CLng(wsProd.Cells(r, 2).Value))

        found = False
        For m = 0 To monthCount - 1
            If monthList(m) = ymStr Then found = True: Exit For
        Next m
        If Not found Then
            ReDim Preserve monthList(monthCount)
            monthList(monthCount) = ymStr
            monthCount = monthCount + 1
        End If

        found = False
        For i = 0 To itemCount - 1
            If itemCodes(i) = code Then found = True: Exit For
        Next i
        If Not found Then
            ReDim Preserve itemCodes(itemCount)
            ReDim Preserve itemNames(itemCount)
            itemCodes(itemCount) = code
            itemNames(itemCount) = GetItemName(code)
            itemCount = itemCount + 1
        End If
    Next r

    ' 月昇順ソート
    Dim tmp As String
    For i = 0 To monthCount - 2
        For m = i + 1 To monthCount - 1
            If monthList(i) > monthList(m) Then
                tmp = monthList(i): monthList(i) = monthList(m): monthList(m) = tmp
            End If
        Next m
    Next i

    ' シミュレーションイン 書き込み
    wsIn.Cells.Clear
    wsIn.Cells(1, 1).Value = "品目コード"
    wsIn.Cells(1, 2).Value = "品目名"
    For m = 0 To monthCount - 1
        wsIn.Cells(1, m + 3).Value = CLng(monthList(m))
    Next m

    For i = 0 To itemCount - 1
        wsIn.Cells(i + 2, 1).Value = itemCodes(i)
        wsIn.Cells(i + 2, 2).Value = itemNames(i)
        For m = 0 To monthCount - 1
            For r = 2 To prodLastRow
                If CStr(wsProd.Cells(r, 1).Value) = itemCodes(i) Then
                    If CStr(CLng(wsProd.Cells(r, 2).Value)) = monthList(m) Then
                        wsIn.Cells(i + 2, m + 3).Value = wsProd.Cells(r, 3).Value
                        Exit For
                    End If
                End If
            Next r
        Next m
    Next i

    MsgBox "Initialization complete: " & itemCount & " items, " & monthCount & " months", vbInformation
    Exit Sub
ErrH:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function GetItemName(itemCode As String) As String
    Dim wsMaster As Worksheet
    Dim lastRow  As Long
    Dim r        As Long

    On Error GoTo NotFound
    Set wsMaster = ThisWorkbook.Sheets("品目マスタ")
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(wsMaster.Cells(r, 1).Value) = itemCode Then
            GetItemName = wsMaster.Cells(r, 2).Value
            Exit Function
        End If
    Next r
NotFound:
    GetItemName = ""
End Function