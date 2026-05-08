# ============================================================================
# Install-BuildSimulationInput.ps1
# ============================================================================
# 目的: BuildSimulationInput VBAマクロを Excel に埋め込む
# 機能: VBAコードをExcelファイルのモジュールとして追加
# ============================================================================

param(
    [string]$ExcelFilePath = "BOM copy 1.xlsm"
)

# ファイルパスを絶対パスに変換
if (-not [System.IO.Path]::IsPathRooted($ExcelFilePath)) {
    $ExcelFilePath = Join-Path (Get-Location) $ExcelFilePath
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "BuildSimulationInput VBA 埋め込みスクリプト" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "対象ファイル: $ExcelFilePath"

# ファイルが存在するかチェック
if (-not (Test-Path $ExcelFilePath)) {
    Write-Host "❌ ファイルが見つかりません: $ExcelFilePath" -ForegroundColor Red
    exit 1
}

# VBAコード文字列（ここに埋め込み）
$VBACode = @'
' ============================================================================
' BuildSimulationInput マクロセット
' ============================================================================
' 目的: 仕様書「シミュレーション入 データセット仕様」を実装
' 機能: 設定シートの開始年月・計画月数に基づいて、
'      生産計画を横持ち形式に変換し、シミュレーション入へ展開
' ============================================================================

Sub BuildSimulationInput()
    '
    ' シミュレーション入を再構築するメインマクロ
    ' 設定 → 生産計画 → シミュレーション入 の自動展開
    '
    ' 仕様書「16. シミュレーション入」のデータセット仕様に準拠：
    '   1. 設定シートから開始年月と計画月数を取得
    '   2. 開始年月を起点に、計画月数分の連続年月列をC列以降に作成
    '   3. 生産計画から対象期間のデータのみ抽出
    '   4. 品目コードごとに横持ち化
    '   5. 対象期間外のデータはセットしない
    '
    
    Dim wsConfig As Worksheet
    Dim wsPlan As Worksheet
    Dim wsIn As Worksheet
    Dim wsMaster As Worksheet
    
    On Error GoTo ErrorHandler
    
    ' シートを取得
    Set wsConfig = ThisWorkbook.Sheets("設定")
    Set wsPlan = ThisWorkbook.Sheets("生産計画")
    Set wsIn = ThisWorkbook.Sheets("シミュレーション入")
    Set wsMaster = ThisWorkbook.Sheets("品目マスタ")
    
    Application.ScreenUpdating = False
    
    ' ========================
    ' Step 1: 設定シートから開始年月と計画月数を読む
    ' ========================
    Dim startYM As Long
    Dim monthCount As Integer
    Dim configStartRow As Long
    
    ' 設定シートは1行目からデータの可能性があるため、複数の位置をチェック
    configStartRow = 2  ' デフォルト位置（1行目がヘッダー）
    
    ' 「開始年月」を探す
    Dim configLastCol As Long
    configLastCol = wsConfig.Cells(1, wsConfig.Columns.Count).End(xlToLeft).Column
    
    Dim findRow As Long, findCol As Long
    findRow = -1
    findCol = -1
    
    ' B1, B2 をチェック（一般的な配置）
    On Error Resume Next
    startYM = CLng(wsConfig.Range("B1").Value)
    monthCount = CInt(wsConfig.Range("B2").Value)
    On Error GoTo ErrorHandler
    
    ' 値の妥当性チェック
    If startYM < 199000 Or startYM > 209912 Or monthCount <= 0 Or monthCount > 120 Then
        ' 別の位置を探す
        startYM = CLng(wsConfig.Cells(2, 2).Value)
        monthCount = CInt(wsConfig.Cells(3, 2).Value)
    End If
    
    ' ========================
    ' Step 2: 生産計画から品目別の月次データを取得
    ' ========================
    Dim planDict As Object
    Set planDict = CreateObject("Scripting.Dictionary")
    
    Dim planLastRow As Long
    planLastRow = wsPlan.Cells(wsPlan.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long, itemCode As String, ym As Long, qty As Double
    
    For r = 2 To planLastRow
        itemCode = CStr(wsPlan.Cells(r, 1).Value)
        itemCode = Trim(itemCode)
        
        If Len(itemCode) > 0 And Not IsError(wsPlan.Cells(r, 2).Value) Then
            ym = CLng(wsPlan.Cells(r, 2).Value)
            
            ' 対象期間のデータのみを抽出
            If IsInTargetPeriod(ym, startYM, monthCount) Then
                qty = 0
                On Error Resume Next
                qty = CDbl(wsPlan.Cells(r, 3).Value)
                On Error GoTo ErrorHandler
                
                ' 品目コードをキーとした辞書を初期化
                If Not planDict.Exists(itemCode) Then
                    planDict.Add itemCode, CreateObject("Scripting.Dictionary")
                End If
                
                ' 年月をキー、数量を値として保存
                planDict(itemCode)(CStr(ym)) = qty
            End If
        End If
    Next r
    
    ' ========================
    ' Step 3: 月列配列を生成
    ' ========================
    Dim months() As Long
    ReDim months(0 To monthCount - 1)
    Dim m As Integer
    
    For m = 0 To monthCount - 1
        months(m) = AddMonths(startYM, m)
    Next m
    
    ' ========================
    ' Step 4: 品目マスタから品目名マップを作成
    ' ========================
    Dim nameMap As Object
    Set nameMap = CreateObject("Scripting.Dictionary")
    
    Dim masterLast As Long
    masterLast = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long, mCode As String, mName As String
    
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
    
    ' ========================
    ' Step 5: シミュレーション入をクリアして再構築
    ' ========================
    wsIn.Cells.Clear()
    
    ' ヘッダー行を作成（A1: 品目コード, B1: 品目名, C1以降: 月列）
    wsIn.Cells(1, 1).Value = "品目コード"
    wsIn.Cells(1, 2).Value = "品目名"
    
    For m = 0 To monthCount - 1
        ' YYYYMM形式の月列をフォーマット
        wsIn.Cells(1, 3 + m).Value = Format(months(m), "000000")
    Next m
    
    ' ========================
    ' Step 6: データ行を作成（品目ごとに横持ち化）
    ' ========================
    Dim outRow As Long
    outRow = 2
    
    Dim sortedKeys() As String
    ReDim sortedKeys(0 To planDict.Count - 1)
    
    Dim keyIdx As Integer, key As Variant
    keyIdx = 0
    
    For Each key In planDict.Keys
        sortedKeys(keyIdx) = CStr(key)
        keyIdx = keyIdx + 1
    Next key
    
    ' ソート（品目コード昇順）
    If planDict.Count > 1 Then
        Call QuickSortString(sortedKeys, LBound(sortedKeys), UBound(sortedKeys))
    End If
    
    ' 品目ごとにデータを書き込み
    Dim sortIdx As Integer, itemKey As String, monthYM As Long
    
    For sortIdx = LBound(sortedKeys) To UBound(sortedKeys)
        itemKey = sortedKeys(sortIdx)
        
        wsIn.Cells(outRow, 1).Value = itemKey
        
        ' 品目名を設定
        If nameMap.Exists(itemKey) Then
            wsIn.Cells(outRow, 2).Value = nameMap(itemKey)
        Else
            wsIn.Cells(outRow, 2).Value = ""
        End If
        
        ' 各月の数量を設定
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
    
    ' ========================
    ' 完了処理
    ' ========================
    Application.ScreenUpdating = True
    
    MsgBox "✓ シミュレーション入を再構築しました。" & vbCrLf & vbCrLf & _
           "開始年月: " & Format(startYM, "0000/00") & vbCrLf & _
           "計画月数: " & monthCount & " ヶ月" & vbCrLf & _
           "品目数: " & (outRow - 2) & " 品目" & vbCrLf & _
           "出力行数: " & (outRow - 1) & " 行", vbInformation, "完了"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました:" & vbCrLf & vbCrLf & _
           Err.Number & ": " & Err.Description, vbCritical, "エラー"
End Sub


Function IsInTargetPeriod(ym As Long, startYM As Long, monthCount As Integer) As Boolean
    '
    ' 指定年月が対象期間内かチェック
    ' 対象期間: startYM から monthCount ヶ月間
    '
    Dim endYM As Long
    endYM = AddMonths(startYM, monthCount)
    
    IsInTargetPeriod = (ym >= startYM And ym < endYM)
End Function


Function AddMonths(baseYM As Long, monthsToAdd As Integer) As Long
    '
    ' YYYYMM形式の年月に月数を加算
    ' 例: AddMonths(202605, 3) = 202608
    '     AddMonths(202612, 2) = 202702
    '
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
    '
    ' 文字列配列をクイックソート（昇順）
    '
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
'@

try {
    # Excelプロセスを停止
    Write-Host ""
    Write-Host "既存のExcelプロセスを停止中..." -ForegroundColor Yellow
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    
    # Excel COM オブジェクトを作成
    Write-Host "Excelを起動中..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # ワークブックを開く
    Write-Host "ワークブックを開く: $ExcelFilePath" -ForegroundColor Yellow
    $wb = $excel.Workbooks.Open($ExcelFilePath, $false, $false)
    
    # 既存のモジュールをチェック・削除
    Write-Host "既存モジュールをチェック..." -ForegroundColor Yellow
    $vbProject = $wb.VBProject
    $modules = $vbProject.VBComponents
    
    # BuildSimulationInput が既に存在する場合は削除
    foreach ($comp in $modules) {
        if ($comp.Name -eq "BuildSimulationInput") {
            Write-Host "  → 既存モジュール '$($comp.Name)' を削除" -ForegroundColor Gray
            $vbProject.VBComponents.Remove($comp)
            break
        }
    }
    
    # 新しいモジュールを追加
    Write-Host "新しいモジュールを追加中..." -ForegroundColor Yellow
    $newModule = $vbProject.VBComponents.Add(1)  # 1 = StandardModule
    $newModule.Name = "BuildSimulationInput"
    
    # VBAコードをモジュールに追加
    Write-Host "VBAコードを埋め込み中..." -ForegroundColor Yellow
    $newModule.CodeModule.AddFromString($VBACode)
    
    # ワークブックを保存
    Write-Host "ワークブックを保存中..." -ForegroundColor Yellow
    $wb.Save()
    
    # クリーンアップ
    $wb.Close($false)
    $excel.Quit()
    
    # COM オブジェクトを解放
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "✓ VBA埋め込み完了" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "次のステップ:" -ForegroundColor Cyan
    Write-Host "  1. Excel を開く: $ExcelFilePath"
    Write-Host "  2. Alt+F8 キーを押す"
    Write-Host "  3. BuildSimulationInput を選択"
    Write-Host "  4. [実行] ボタンをクリック"
    Write-Host ""
    
} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "❌ エラーが発生しました" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "エラー内容:" -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    exit 1
}
