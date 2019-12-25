Attribute VB_Name = "Module5"
Option Explicit

Private Sub CopyData()
Dim Review As Workbook
Dim Netsuite As Workbook
Set Review = ThisWorkbook
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Dim NetStr As String
NetStr = Application.GetOpenFilename(Title:="Choose an Excel file to open")
If InStr(NetStr, "False") = 0 Then
Set Netsuite = Workbooks.Open(NetStr)
Netsuite.Sheets(1).Activate
Netsuite.Sheets(1).Range("A2:Q" & CStr(Range("A65536").End(xlUp).Row)).Copy
Review.Sheets(2).Range("B5").PasteSpecial
Netsuite.Close
End If
Call FormatData
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Private Sub FormatData()
Dim LastRow As Integer
LastRow = ActiveSheet.UsedRange.Rows.Count
Dim CusName As Range
Dim Quantity As Range
Set CusName = Range("C5:C" + CStr(LastRow))
Set Quantity = Range("P5:P" + CStr(LastRow))
CusName.Select
'changes the customer name field to right alignment
With Selection
.HorizontalAlignment = xlRight
End With
Quantity.Select
'changes the quantity field to center alignment
With Selection
.HorizontalAlignment = xlCenter
End With
Dim Rowx As Integer
Dim Day As String
'changes the service date to YYYY-MM-DD format
For Rowx = 5 To LastRow
Day = Cells(Rowx, 10).Text
Cells(Rowx, 10).Value = Right(Day, 4) + "-" + Left(Day, 2) + "-" + Mid(Day, 4, 2)
Next Rowx
'autofill formula
With ThisWorkbook.Sheets(2)
.Range("AC6:AE" & LastRow).FillDown
End With
End Sub


