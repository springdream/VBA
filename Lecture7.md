# 1.初めに

一番簡単なプログラム
```
Sub test()
	Msgbox "hello world" 'hello worldと表示
End Sub
```
シングルクォーテーションの右側はプログラムに反映されない(コメントアウト)．

# 2.変数

```
Sub test()
	Dim str As String '文字列型
	Dim cnt As Long '整数型
	Dim fl As Double '浮動小数点型（要は小数）
	Dim bl As Boolean '真偽値型
End Sub
```
# 3.入出力

メッセージボックスの出し方
```
Sub test()
	MsgBox "hello world"
End Sub
```

セルの入力の仕方

```
Sub test()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Range("A1").Value = "hello"
End Sub
```
セルの出力の仕方 
```
Sub test()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    MsgBox ws.Range("A1").Value 'A1の内容がメッセージボックスに表示
End Sub
```
# 4. For文

```
Sub test()
    For i = 1 To 5
        MsgBox i
    Next i
End Sub
```

# 4.2 Do Loop
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
# 5. If文

```
Sub main()
    Dim count As Long
    count = 100
    If count < 50 Then
        MsgBox "50より小さいです"
	Else    'Else msgboxの部分は省略可能
        MsgBox "50以上です"
    End If
End Sub
```

If文のIf,thenで挟まれた部分には真偽値を入れる
```
Sub test()
	Dim bl As Boolean
	bl = True
	If bl Then 
		msgbox "真です"
	End If
End Sub
```
真偽値は以下がある
```
Sub test()
	Dim bl As Boolean 
	bl = True '真
	bl = False '偽
	bl = 3 < 5 '真
	bl = 3 <> 4 ' notイコール

End Sub
```
# 6.　選択範囲の取得
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

7