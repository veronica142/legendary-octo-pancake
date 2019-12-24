Attribute VB_Name = "Module5"
Option Explicit

Private Sub ChangeAll(Foldername As String)

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim Name As String
    Dim i As Integer, r As Integer, c As Integer
    Dim RawData As Workbook
    
    For i = 1 To 5
    Dim wb As Workbook
    Name = Foldername & "\" & CStr(i) & ".csv"
    Set RawData = Workbooks.Open(Name)
    Set wb = Workbooks.Add
    wb.Activate
    If i = 1 Then
    RawData.Worksheets(1).UsedRange.Copy Destination:=wb.Worksheets(1).Range("A1")
    Else
    Call Encoding(Name)
    End If
    wb.SaveAs (Foldername & "\" & "_" & CStr(i))
    wb.Close
    RawData.Close
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub


Private Sub Encoding(Name As String)

    ActiveWorkbook.Queries.Add Name:="2", Formula:= _
        "let" & Chr(13) & Chr(10) & "   Source = Csv.Document(File.Contents(""" & Name & """),[Delimiter="","", Encoding=65001])," & Chr(13) & Chr(10) & "   #""First Row as Header"" = Table.PromoteHeaders(Source)" & Chr(13) & Chr(10) & "in" & Chr(13) & Chr(10) & "    #""First Row as Header"""
    ActiveWorkbook.Worksheets.Add
    
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=2;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [2]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "_2"
        .Refresh BackgroundQuery:=False
    End With
    
End Sub


