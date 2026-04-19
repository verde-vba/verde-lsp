Attribute VB_Name = "Utilities"
Option Explicit

Public Function FindSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    Set FindSheet = ws
End Function

Public Function FindRowsByPrefix(ws As Worksheet, ByVal col As Long, ByVal prefix As String) As Collection
    Dim result As New Collection
    Dim lastRow As Long
    Dim i As Long
    Dim cellVal As String

    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    For i = 1 To lastRow
        cellVal = CStr(ws.Cells(i, col).Value)
        If Left(cellVal, Len(prefix)) = prefix Then
            result.Add i
        End If
    Next i

    Set FindRowsByPrefix = result
End Function

Public Function FindTotalRow(ws As Worksheet, ByVal col As Long) As Long
    Dim lastRow As Long
    Dim i As Long

    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    For i = lastRow To 1 Step -1
        If Trim(CStr(ws.Cells(i, col).Value)) = "Total" Then
            FindTotalRow = i
            Exit Function
        End If
    Next i

    FindTotalRow = 0
End Function

Public Function FormatAsCurrency(ByVal amount As Double) As String
    FormatAsCurrency = Format(amount, "$#,##0.00")
End Function
