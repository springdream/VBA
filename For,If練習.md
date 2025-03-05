
# Q1. 1から10の和を求めよ
##### 方針
	合計を保存する変数を作って	
	そこに1から10を追加していく
```
Sub test1()
	Dim sum As Long
	For i = 1 to 10 
		sum = sum + i 
	Next i 
	MsgBox sum
End Sub
```

# Q2. 1から10の内偶数の和を求めよ
##### 方針 
	 Q1と基本は同じで，値を足すときにそれが偶数か考え，偶数なら足す 

```
Sub test2()
	Dim sum As Long 
	For i = 1 to 10 
		If i Mod 2 = 0 Then
			sum = sum + i
		End If
	Next i
	MsgBox sum
End Sub
```
 

# Q3. 九九の解全体の和を求めよ
##### 方針
	For文を2回使う
```
Sub kuku()
	Dim sum As long 
		For i = 1 to 9
			For j = 1 to 9
			sum = sum + i*j
			Next j
		Next i
	MsgBox sum
End Sub
```


