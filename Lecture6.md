# 辞書の使用例

### 1.同じ名前のデータの合計を求める

### 2.同じ名前のデータの最大値を求める 


# 配列
```
Sub test()
	Dim scores(4) As Integer '
	scores(0) = 80
```

```
Sub test()
    Dim a(4) As Long
    Dim i  As Long
    i = 1
    a(i) = 100
    MsgBox a(i)
    
End Sub
```

```
Sub test()
    Dim a(4) As Long
    Dim i  As Long
    i = 1
    For i = 0 To 4
        a(i) = 100 * i
    Next i
    
    For i = 0 To 4
        MsgBox a(i)
    Next i
    
End Sub
```