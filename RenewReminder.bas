Attribute VB_Name = "Module6"
Sub RenewReminder()
Dim oWB As Workbook
Set oWB = Excel.ThisWorkbook
oWB.ResetColors
Dim LastRow As Integer
Worksheets(1).Activate
LastRow = ActiveSheet.UsedRange.Rows.Count
Dim Length As Integer
Length = InputBox("How many months do you want the reminder to include?", "Reminder Period")
Dim oDialog As Dialog
Set oDialog = Excel.Application.Dialogs(xlDialogEditColor)
If oDialog.Show(1) = True Then
    iColor = oWB.Colors(1)
    Worksheets(1).Activate
    For i = 6 To LastRow
    Dim Product As String
    Product = Cells(i, "E").Value
    Dim InitiateDate As String
    Dim Today As Date
    Today = Date
    InitiateDate = Cells(i, "I").Value
    'check air purifiers leasing
    If InStr(Product, "BlueAir") <> 0 And InStr(Product, "03") <> 0 Then
    If InitiateDate > Today - 365 And InitiateDate < Today + Length * 30 - 365 Then
    Cells(i, "A").Interior.Color = iColor
    End If
    End If
    Next i
End If
MsgBox ("Please filter on the Column A: all customers that need renewal are in the color you selected.")
End Sub


