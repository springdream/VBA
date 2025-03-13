
#### 経過時間を求めて右側に入れる
![](attachments/Pasted%20image%2020250313204202.png)
```
Sub test()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    For i = 1 To 3 - 1
        ws.Cells(i, "B") = ws.Cells(i + 1, "A") - ws.Cells(i, "A")
    Next i
End Sub
```
![](attachments/Pasted%20image%2020250313204236.png)