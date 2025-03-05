
# 1.　選択範囲に"hi"を入力
```
Sub ran()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim rng As Range
    Set rng = Selection
    'MsgBox "present range:" & rng.Address
    ws.Range(rng.Address).Value = "hi"
    
End Sub

```
# 2. 選択範囲の数字に100を加算
```
Sub ran()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim rng As Range
    Set rng = Selection
    'MsgBox "present range:" & rng.Address
    
    ' 選択範囲の各セルに 100 を加算
    Dim cell As Range
    For Each cell In rng
        ' セルの値が数値の場合のみ加算
        If IsNumeric(cell.Value) Then
            cell.Value = cell.Value + 100
        End If
    Next cell
    
End Sub
```
# 3. 選択範囲の値の合計を求める
```
Sub wa()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim sum As Long
    
    Dim rng As Range
    Set rng = Selection
    'MsgBox "present range:" & rng.Address
    
    '
    Dim cell As Range
    For Each cell In rng
        ' セルの値が数値の場合のみ加算
        sum = sum + cell.Value
    Next cell
    MsgBox sum
    
End Sub
```

# [[TODO]] 
関数の作り方と細かい解説
