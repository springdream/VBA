### 復習
前回時間の合計を求めるVBAを作った
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
        ws.Cells(row_number, "H").NumberFormat = "hh:mm"
        row_number = row_number + 1
    Next cur_key
        
End Sub
```
これを使い易く改造していく

```

Sub keika_jikan3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim startRow As Long
    Dim lastRow As Long
    startRow = Selection.Row
    lastRow = startRow + Selection.Rows.Count - 1
    
    For i = startRow To lastRow - 1
        ws.Cells(i, "E").Value = ws.Cells(i + 1, "B").Value - ws.Cells(i, "B").Value
        ws.Cells(i, "E").NumberFormat = "h:mm"
    Next i
        
End Sub

Sub Insert_sum_Time3()
 Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set taskDict = CreateObject("Scripting.Dictionary")
    Dim key As String
    Dim val As Double
    Dim startRow As Long
    Dim lastRow As Long
    startRow = Selection.Row
    lastRow = startRow + Selection.Rows.Count - 1
    
    For i = startRow To lastRow - 1
        key = ws.Cells(i, "D")
        val = ws.Cells(i, "E")
        
        If taskDict.Exists(key) Then
            taskDict(key) = taskDict(key) + val
        Else
            taskDict.Add key, val
        End If
    Next i
    
    Dim row_number As Long
    row_number = startRow
    For Each cur_key In taskDict.Keys
        ws.Cells(row_number, "G").Value = cur_key
        ws.Cells(row_number, "H").Value = taskDict(cur_key)
        ws.Cells(row_number, "H").NumberFormat = "hh:mm"
        row_number = row_number + 1
    Next cur_key

Sub cntl3()
    keika_jikan3
    Insert_sum_Time3
End Sub

```

