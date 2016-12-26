# excel-spreadsheet-copier
Sample code for writing an excel macro that copies and pastes rows based on certain conditions


```
Sub spreadsheet_copier()
For Each Cell In Sheets(1).Range("B:B")
    If InStr(1, Cell.Value, "someString") > 0 Then
        matchRow = Cell.Row
        Rows(matchRow & ":" & matchRow).Select
        Selection.Copy
        Sheets("Sheet2").Select
        ActiveSheet.Rows(Range("B" & Rows.Count).End(xlUp).Row + 1).Select
        ActiveSheet.Paste
        Sheets("Sheet1").Select
    End If
    
    If InStr(1, Cell.Value, "someOtherString") > 0 Then
        matchRow = Cell.Row
        Rows(matchRow & ":" & matchRow).Select
        Selection.Copy
        Sheets("Sheet3").Select
        ActiveSheet.Rows(Range("B" & Rows.Count).End(xlUp).Row + 1).Select
        ActiveSheet.Paste
        Sheets("Sheet1").Select
    End If
    
Next
End Sub
```
