
### モジュール
ファイル分けをした方が見やすいから分けているだけ．
同じモジュールに複数の関数を用意してもよい

### 日付の挿入
### オブジェクト
オブジェクト型の変数にオブジェクトを代入するにはsetが必要
整数，小数，文字列の型以外は大体オブジェクト
```
Sub objtest()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") 'Worksheetはオブジェクト
    
	Dim rng As Range
	set rng =  ws.Cells(2,"B") 'Rangeもオブジェクト
End Sub
```
### A列最後尾の行の２つ下に日付を追加し，赤く塗る
```
Sub add_date3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim lastRow As Long
    lastRow = ws.Rows.count 'シートの行数 つまり，一番下の番号
    lastRow = ws.Cells(lastRow, "A").End(xlUp).Row 
	'ws.Cells(lastRow, "A") A最後のセル
	' .End(xlUp) ctrl+↑    .End(xlDown) ctrl+↓
	'.Row 行番号
	 
    ws.Cells(lastRow + 2, "A").Value = Now
    ws.Cells(lastRow + 2, "A").Interior.Color = RGB(255, 0, 0)

End Sub
```

### cell.end.rowを２行に分割
```
Sub add_date3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim lastRow As Long
    Dim rng As Range
    lastRow = ws.Rows.count
   ' lastRow = ws.Cells(lastRow, "A").End(xlUp).Row
    Set rng = ws.Cells(lastRow, "A").End(xlUp)
    lastRow = rng.Row
    ws.Cells(lastRow + 2, "A").Value = Now
    ws.Cells(lastRow + 2, "A").Interior.Color = RGB(255, 0, 0)

End Sub

```
### A列を上から見ていって初めて出た空のセルに日付を入れる
```
Sub add_date2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim i As Long
    i = 1
    Do
        If ws.Cells(i, "A") = "" Then
            ws.Cells(i, "A") = Now
            Exit Do
        End If
        i = i + 1
    Loop
    
End Sub

```