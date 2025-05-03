# 辞書の使用例

### 1.同じ名前のデータの合計を求める

### 2.同じ名前のデータの最大値を求める 


# 配列(辞書とは違う)
配列は辞書のカギが0,1,2,... N-1 (N個)の連続した数字になったバージョン

```
Sub test()
	Dim scores(4) As Long '配列の宣言 0~3の４つの変数を作るイメージ
	scores(0) = 80
	scores(1) = 60
	scores(2) = 50
	scores(3) = 40
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