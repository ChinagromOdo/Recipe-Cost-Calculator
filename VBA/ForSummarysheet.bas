Attribute VB_Name = "ForSummarysheet"
Option Explicit

Function ListSheets(RowNum As Long) As String
    On Error Resume Next
    ListSheets = ThisWorkbook.Sheets(RowNum).Name
    
End Function

Sub AutoSheetName()
    Dim RowNum As Long
    Dim TotalSheets As Long
    Dim c As Range
    Dim CellValue As String
    Dim LastRow As Long
    Dim CheckRange As Range

    TotalSheets = ThisWorkbook.Sheets.Count
    RowNum = 5
    
    Do While RowNum <= TotalSheets
        Range("B" & 1 + RowNum).Value = ListSheets(RowNum)
        RowNum = RowNum + 1
        
        
    Loop
    LastRow = Cells(Rows.Count, "B").End(xlUp).Row
    Set CheckRange = Range("B6:B" & LastRow)
    For Each c In Range("B6:B" & Cells(Rows.Count, "B").End(xlUp).Row)
    CellValue = c.Value
    If Application.WorksheetFunction.CountIf(CheckRange, CellValue) > 1 Then
        c.Delete Shift:=xlUp
    End If
    Next c
    Columns("B:B").Select
    Selection.Columns.AutoFit
End Sub

Sub HyperLinking()
Dim c As Range
For Each c In Range("B6:B" & Cells(Rows.Count, "B").End(xlUp).Row)
    If c <> "" Then
        c.Hyperlinks.Add Anchor:=c, Address:="", _
        SubAddress:="'" & c.Value & "'!A1", TextToDisplay:=c.Value
    End If
    Next c
End Sub


Sub UpdateSummarySheet()
    AutoSheetName
    HyperLinking
End Sub
    

