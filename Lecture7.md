### これまでの復習


#### 経過時間を求めて右側に入れる
```
Sub test()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    For i = 1 To 3 - 1
        ws.Cells(i, "B") = ws.Cells(i + 1, "A") - ws.Cells(i, "A")
    Next i
End Sub
```