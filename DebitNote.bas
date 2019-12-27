Attribute VB_Name = "Module1"
Option Explicit


Private Sub FormatCustomerName()
'formatting customers name with type (xx):(xx):xx(##-##-##) by seperating letters and numbers and spliting with respect to comma
'for strings that are splitted into multiple parts, cells to its right are used for the rest of the string
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim Rng As Range
Dim LastRow As Integer
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Set Rng = Range("A15:A" + CStr(LastRow))
Dim xRows As Integer
xRows = Rng.Cells.Count
For i = 1 To xRows
Dim a As Variant
Dim b As Integer
Dim Name As String
Name = Rng.Cells(i, 1).Value
a = Split(Name, ":")
b = UBound(a)
For j = 0 To b - 1
Rng.Offset(0, j).Cells(i, 1) = a(j)
Next j
Dim Customer As String
Dim Barcode As String
Customer = a(b)
Barcode = ""
For n = 1 To Len(Customer)
If Mid(Customer, n, 1) Like "[0-9]" Then
Barcode = Right(Customer, Len(Customer) - n + 1)
Customer = Left(Customer, n - 1)
Exit For
End If
Next n
Rng.Offset(0, b).Cells(i, 1) = Customer
Rng.Offset(0, b + 1).Cells(i, 1) = Barcode
Next i
Call ChangeDescript
End Sub

Private Sub ChangeDescript()
'changes the description for each service, leave the description as it is if it's not in the following types
'enter the range for which desciption needs change
Dim Rng As Range
Dim LastRow As Integer
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Set Rng = Range("F15:F" + CStr(LastRow))
Dim xRows As Integer
xRows = Rng.Cells.Count
For i = 1 To xRows
Dim Val As Variant
Val = Rng.Cells(i, 1).Value
    'IEQ Testing
    If InStr(Val, "Health") <> 0 Then
    Rng.Cells(i, 1).Value = Replace(Val, Val, "IEQ Testing")
    'PM2.5 Audition
    ElseIf InStr(Val, "PM2.5") <> 0 Then
    Rng.Cells(i, 1).Value = Replace(Val, Val, "PM2.5 Air Testing")
    'Lost Machine
    ElseIf InStr(Val, "lost") <> 0 Then
    Rng.Cells(i, 1).Value = Replace(Val, Val, "Lost Machine")
    'Technician Fee
    ElseIf InStr(Val, "echnician") <> 0 Then
    Rng.Cells(i, 1).Value = Replace(Val, Val, "Technician Fee")
    'Moving Fee
    ElseIf InStr(Val, "mov") <> 0 Or InStr(Val, "Mov") <> 0 Then
    Rng.Cells(i, 1).Value = Replace(Val, Val, "Moving Fee")
    'Second-year Renewal
    ElseIf InStr(Val, "renew") <> 0 Or InStr(Val, "2nd year") <> 0 Or InStr(Val, "second year") <> 0 Then
    Rng.Cells(i, 1).Value = Replace(Val, Val, "IEQ 12 month Installation - renewal")
    'First-year Installation
    ElseIf InStr(Val, "12 month leasing") <> 0 Or InStr(Val, "12 months leasing") <> 0 Or InStr(Val, "12 months leasing") <> 0 _
    Or InStr(Val, "cable extension") <> 0 Or InStr(Val, "dual faucet") <> 0 Then
    Rng.Cells(i, 1).Value = Replace(Val, Val, "IEQ 12 month Installation")
    'Water or Air Filter Replacement
    ElseIf InStr(Val, "rep") <> 0 Or InStr(Val, "Rep") <> 0 Then
    Dim newval As String
    If InStr(Val, "203") <> 0 Or InStr(Val, "403") <> 0 Or InStr(Val, "503") <> 0 Then
    newval = "Replacement Fee: Air Filters"
    ElseIf InStr(Val, "latai") <> 0 Or InStr(Val, "WB") <> 0 Or InStr(Val, "CF") <> 0 Or InStr(Val, "SK") <> 0 _
    Or InStr(Val, "aterbaby") <> 0 Or InStr(Val, "learfall") <> 0 Or InStr(Val, "howerking") <> 0 Then
    newval = "Replacement Fee: Water Filters"
    If InStr(Val, "203") <> 0 Or InStr(Val, "403") <> 0 Or InStr(Val, "503") Then
    newval = newval + "and Air Filters"
    End If
    End If
    Rng.Cells(i, 1).Value = Replace(Val, Val, newval)
    'Additional Machine
    ElseIf InStr(Val, "dditional") <> 0 Then
    Rng.Cells(i, 1).Value = Replace(Val, Val, "Additional Machine")
    End If
    Next
    Call MergeCustomerName
End Sub

Private Sub MergeCustomerName()
'merge cells for same customer name
Dim RngA As Range
Dim xRows As Integer
xRows = Cells(Rows.Count, 1).End(xlUp).Row - 2
Set RngA = Range("C15:C" + CStr(xRows))
For i = 1 To xRows
    For j = i + 1 To xRows
        If RngA.Cells(i, 1).Value <> RngA.Cells(j, 1).Value Then
            Exit For
        End If
    Next
    Range(RngA.Cells(i, 1), RngA.Cells(j - 1, 1)).merge
    Range(RngA.Cells(i, 1).Offset(0, -1), RngA.Cells(j - 1, 1).Offset(0, -1)).merge
    Range(RngA.Cells(i, 1).Offset(0, -2), RngA.Cells(j - 1, 1).Offset(0, -2)).merge
    Range(RngA.Cells(i, 1).Offset(0, 1), RngA.Cells(j - 1, 1).Offset(0, 1)).merge
    i = j - 1
Next
Call TaxRate
End Sub


Private Sub TaxRate()
'determine the tax rate of each service
'6% for testing and inspection and 13% for all others
Dim Rowx As Integer
Dim LastRow As Integer
LastRow = ActiveSheet.UsedRange.Rows.Count
Dim descript As String
For Rowx = 15 To LastRow
descript = Range("F" + CStr(Rowx)).Text
If InStr(descript, "esting") <> 0 Or InStr(descript, "echnician Fee") <> 0 Or InStr(descript, "nspection") Then
Range("H" + CStr(Rowx)).Value = 0.06
ElseIf InStr(descript, "ubtotal") <> 0 Or InStr(descript, "TOTAL") <> 0 Or IsEmpty(Range("F" + CStr(Rowx))) Then
Range("H" + CStr(Rowx)).ClearContents
Else
Range("H" + CStr(Rowx)).Value = 0.13
End If
Next Rowx
Call Fapiao
End Sub

Private Sub Fapiao()
'merges cells that should have the same fapiao number and enters the total fapiao amount in Column I
Dim xRows As Integer
xRows = Cells(Rows.Count, 1).End(xlUp).Row - 2
For i = 15 To xRows
    Dim FapiaoAmount As Long
    FapiaoAmount = Cells(i, "G").Value
    For j = i + 1 To xRows
        If Cells(i, "H").Value <> Cells(j, "H").Value Then
            Exit For
        Else
            FapiaoAmount = FapiaoAmount + Cells(j, "G").Value
        End If
    Next
    Range(Cells(i, "H"), Cells(j - 1, "H")).merge
    Range(Cells(i, "I"), Cells(j - 1, "I")).merge
    Cells(i, "I").Value = FapiaoAmount
    i = j - 1
Next
Call Subtotal
End Sub

Private Sub Subtotal()
'add subtotal under each customer
Range("C15").Select
    Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(5), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
        

