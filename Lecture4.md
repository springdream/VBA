
# 関数(やってない)
関数とは，何らかの値を受け取って，何らかの処理をして，何らかの値を返すもの.
下は，数字を受け取って，受け取った数字を＋１して返す関数．
```
Function AddOne(ByVal num As Integer) As Integer
    AddOne = num + 1
End Function

```

# xlup
A1~A5に，1,2,3,4,5が入ってたら，A6にA1~A5の合計を追加する.
```
'ある範囲の合計を求める

Sub test()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim lastRow As Long
    Dim startRow As Long
    lastRow = ws.Rows.Count 'ブックそのものの一番下
    lastRow = ws.Cells(lastRow, "A").End(xlUp).Row 'A列の値の一番下のデータ
    startRow = ws.Cells(lastRow, "A").End(xlUp).Row
    
    Dim sum As Long
    sum = 0
    For i = startRow To lastRow
        sum = sum + ws.Cells(i, "A").Value
    Next i
    
    ws.Cells(lastRow + 1, "A").Value = sum
    
End Sub


```
#  Do Loop

```
'Loopについて
Sub test()
    Dim count As Long
    count = 0
    Do
        If count > 5 Then
            Exit Do
        End If
        MsgBox count
        count = count + 1
    Loop
    
End Sub

```

```
'Loopについて
Sub test()
    Dim i As Long
    i = 1
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    Do
        If ws.Cells(i, "D").Value <> "" Then
            MsgBox ws.Cells(i, "D").Value
            i = i + 1
        Else
            MsgBox "over range"
            Exit Do
        End If
    Loop
    
End Sub


```

# 時間
```
'時間
Sub test()
    Dim ws As Worksheet
    Dim n As Double
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ws.Cells(1, "E").Value = Time
    ws.Cells(1, "E").NumberFormat = "h:mm"
    
    ws.Cells(1, "F").Value = Now
    ws.Cells(1, "F").NumberFormat = "yyyy/mm/dd/h:mm:ss"
    
    MsgBox Now
    ws.Cells(1, "G").Value = Now
    n = Now
    MsgBox n
    ws.Cells(1, "H").Value = n
End Sub

```