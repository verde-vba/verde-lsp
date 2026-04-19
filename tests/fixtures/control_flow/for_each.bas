Attribute VB_Name = "ForEach"
Option Explicit

' For Each with Range
Sub ForEachRange()
    Dim c As Range
    For Each c In Worksheets("Sheet1").Range("A1:D10")
        If c.Value < .001 Then
            c.Value = 0
        End If
    Next c
End Sub

' For Each counting blanks
Sub CountBlanks()
    Dim numBlanks As Long
    Dim c As Range
    numBlanks = 0
    For Each c In Range("TestRange")
        If c.Value = "" Then
            numBlanks = numBlanks + 1
        End If
    Next c
    MsgBox "There are " & numBlanks & " empty cells in this range."
End Sub

' For Each with Workbooks
Sub CloseAllWorkbooks()
    Dim w As Workbook
    For Each w In Workbooks
        If w.Name <> ThisWorkbook.Name Then
            w.Close savechanges:=True
        End If
    Next w
End Sub

' For Each with Worksheets
Sub DeleteAllSheets()
    Dim w As Worksheet
    Application.DisplayAlerts = False
    For Each w In Worksheets
        w.Delete
    Next w
    Application.DisplayAlerts = True
End Sub

' For Each with Names collection
Sub ListNames()
    Dim newSheet As Worksheet
    Set newSheet = ActiveWorkbook.Worksheets.Add
    Dim i As Long
    Dim nm As Name
    i = 1
    For Each nm In ActiveWorkbook.Names
        newSheet.Cells(i, 1).Value = nm.NameLocal
        newSheet.Cells(i, 2).Value = "'" & nm.RefersToLocal
        i = i + 1
    Next nm
End Sub

' For Each with array elements
Sub ForEachArray()
    Dim elements(5) As String
    Dim element As Variant
    For Each element In elements
        Debug.Print element
    Next element
End Sub
