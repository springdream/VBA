
### ボタンの追加
![](attachments/Clipboard%20-%202025-03-27%2020.48.18.png)

開発タブの挿入
![](attachments/Pasted%20image%2020250327204953.png)
左上のボタンをクリック
その後，シート上をクリック
マクロ名を選択，OK


### 日付の挿入
```
Sub add_date()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Cells(1, "A").Value = Now
End Sub
```

### セルの色を塗る

```
Sub color()
	Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Cells(1,"A").Interior.Color = RGB(255,255,0)
```