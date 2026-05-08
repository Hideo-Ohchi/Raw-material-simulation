# VBA修正内容

## 問題
設定シートの開始年月を202606に変更してもシミュレーション入は202605から始まっていた

## 原因
VBAコードの設定値読み込み部分が複雑で、値を正しく読み込めていなかった

## 修正内容

### 修正前のコード
```vba
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
```

### 修正後のコード
```vba
' 設定シート B1, B2 から直接読み込み
' B1: 開始年月（YYYYMM形式）
' B2: 計画月数

Dim b1Value As Variant
Dim b2Value As Variant

b1Value = wsConfig.Range("B1").Value
b2Value = wsConfig.Range("B2").Value

' 値の型変換と妥当性チェック
On Error Resume Next
startYM = CLng(b1Value)
monthCount = CInt(b2Value)
On Error GoTo ErrorHandler

' 妥当性チェック
If IsEmpty(startYM) Or startYM <= 0 Or monthCount <= 0 Then
    MsgBox "エラー: 設定シートのB1（開始年月）またはB2（計画月数）が正しく設定されていません。" & vbCrLf & vbCrLf & _
           "B1（開始年月）: " & b1Value & vbCrLf & _
           "B2（計画月数）: " & b2Value, vbCritical, "設定エラー"
    Application.ScreenUpdating = True
    Exit Sub
End If
```

## 改善点

| 項目 | 修正前 | 修正後 |
|------|-------|-------|
| **読み込み位置** | B1, B2または別の位置を探索 | B1, B2 に直接限定 |
| **妥当性チェック** | 複雑な条件分岐 | シンプルな条件判定 |
| **エラーメッセージ** | なし | 読み込んだ値を表示 |
| **デバッグ性** | 低い | 高い |

## 実行方法

```
1. Excel を開く：原料シミュレーション.xlsm
2. 設定シートで以下を確認
   ├─ A1: 「開始年月」
   ├─ B1: 202606（希望する開始年月）
   ├─ A2: 「計画月数」
   └─ B2: 3（計画月数）
3. Alt+F8 キー
4. BuildSimulationInput を選択
5. [実行] をクリック
6. 完了メッセージを確認
7. シミュレーション入シートで月列を確認
   → C1: 202606, D1: 202607, E1: 202608 になるはず
```

## 期待される動作

### 修正前
```
月列: 202605, 202606, 202607
      ↑ 設定が反映されていない
```

### 修正後
```
月列: 202606, 202607, 202608
      ↑ 設定値が正しく反映される
```

## エラー時の確認ポイント

エラーメッセージが出た場合：

```
エラー: 設定シートのB1（開始年月）またはB2（計画月数）が正しく設定されていません。

B1（開始年月）: [表示される値]
B2（計画月数）: [表示される値]
```

上記のメッセージから、実際に読み込まれた値を確認できます。

---

✅ **VBA更新完了**  
Excelファイルにコードを埋め込み済みです。マクロを再実行してください。
