# 予定
### 時間の集計を作る
開始時間 | 仕事内容 | 分類 | 
この表から

### 経過時間を求める
```
Sub keika_jikan()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim startRow As Long
    Dim lastRow As Long
    startRow = 1
    lastRow = 17
    For i = startRow To lastRow - 1
        ws.Cells(i, "E").Value = ws.Cells(i + 1, "B").Value - ws.Cells(i, "B").Value
        ws.Cells(i,"E").NumberFormat = "h:mm"
    Next i
        
End Sub
```

### 辞書を使って合計時間を取得

```
Sub sum_time()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set taskDict = CreateObject("Scripting.Dictionary")
    Dim key As String
    Dim val As Double
    
    For i = 1 To 16
        key = ws.Cells(i, "D")
        val = ws.Cells(i, "E")
        
        If taskDict.Exists(key) Then
            taskDict(key) = taskDict(key) + val
        Else
            taskDict.Add key, val
        End If
    Next i
      
    For Each cur_key In taskDict.Keys
	    MsgBox "key:" + cur_key + ", Value:" + CStr(taskDict(cur_key))
    Next cur_key
    
End Sub
```

### 辞書を使って合計時間をシートに挿入

```
Sub Insert_sum_Time()
 Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set taskDict = CreateObject("Scripting.Dictionary")
    Dim key As String
    Dim val As Double
    
    For i = 1 To 16
        key = ws.Cells(i, "D")
        val = ws.Cells(i, "E")
        
        If taskDict.Exists(key) Then
            taskDict(key) = taskDict(key) + val
        Else
            taskDict.Add key, val
        End If
    Next i
    
    Dim row_number As Long
    row_number = 1
    For Each cur_key In taskDict.Keys
        ws.Cells(row_number, "G").Value = cur_key
        ws.Cells(row_number, "H").Value = taskDict(cur_key)
        ws.Cells(row_number, "H").NumberFormat = "h:mm"
        row_number = row_number + 1
    Next cur_key
        
End Sub
```