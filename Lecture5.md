# 辞書
名前と値の組み合わせ．名前を「鍵」，値は「値」と呼ぶ．

![|400x300](attachments/Clipboard%20-%202025-03-05%2003.24.53.png)


```
Sub test()
Dim taskDict As Object '辞書の変数を宣言 taskDictは何でもいい
Set taskDict = CreateObject("Scripting.Dictionary") '辞書本体の作成

taskDict.Add "A", 100 '辞書にA -> 100 を追加
taskDict.Add "B", 200 '辞書にB -> 200　を追加


MsgBox taskDict("A") '辞書のAを表示 (つまり100が表示)

For Each Key In taskDict.Keys 'For文で辞書の中身全部を表示
MsgBox "key:" + Key + ", Value:" + CStr(taskDict(Key))'CStrは数字を文字列にする
Next Key

End Sub

```

taskDict("A")は変数のように扱える．
```
Dim a as long
a = taskDict("A") 'a = 100 
taskDict("A") = 300 ' 
```
のようなことができる．


### 同じ名前のデータの合計を求める
(解説は後でする)
エクセルの一部(Sheet1)
![](attachments/Clipboard%20-%202025-03-05%2003.31.12%201.png)

```
Sub test()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    '辞書の作成
    Set taskDict = CreateObject("Scripting.Dictionary")
    Dim key1 As String
    Dim val As Long
    
    For i = 1 To 3
        key1 = ws.Cells(i, "A")
        val = ws.Cells(i, "B")
        
        If taskDict.Exists(key1) Then
            taskDict(key1) = taskDict(key1) + val
        Else
            taskDict.Add key1, val '
        End If
    Next i
      
    For Each heyano_namae In taskDict.Keys 'taskDict.Keys = {"cat","dog"}
    MsgBox "key:" + heyano_namae + ", Value:" + CStr(taskDict(heyano_namae))
    Next heyano_namae
    
    
End Sub


```

# Practice
以下の表から.上でやったように各動物の体重を求めよ．

| 名前     | 体重/Kg |
| ------ | ----- |
| dog    | 3     |
| dog    | 4     |
| cat    | 2     |
| rabbit | 1     |
| rabbit | 1     |
