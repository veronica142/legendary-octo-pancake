Attribute VB_Name = "Module4"
Sub DateDiff()
'changes the service date of sales orders in summary sheet to the date in data sheet
'finds all cells with the corresponding SO number in summary sheet, then mutates the date in next cell
Dim SalesOrder As String
Dim LastRow As Integer
Sheets(2).Activate
LastRow = ActiveSheet.UsedRange.Rows.Count
For i = 5 To LastRow
If Cells(i, "AE") <> 0 And IsEmpty(Cells(i, "AE")) = False Then
'Find and replace
SalesOrder = Worksheets(2).Cells(i, "I").Value
CurDate = Worksheets(2).Cells(i, "J").Value
With Worksheets(1).UsedRange
    Set c = .Find(SalesOrder, LookIn:=xlValues)
    If Not c Is Nothing Then
        FirstAddress = c.Address
        Do
            c.Offset(0, 1).Value = CurDate
            Set c = .FindNext(c)
        Loop While Not c.Address = FirstAddress And Not c Is Nothing
    End If
End With
End If
Next i
End Sub

