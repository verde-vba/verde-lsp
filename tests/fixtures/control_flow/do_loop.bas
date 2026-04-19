Attribute VB_Name = "DoLoop"
Option Explicit

' Do While ... Loop
Sub DoWhileLoop()
    Dim currentCell As Range
    Dim nextCell As Range
    Worksheets("Sheet1").Range("A1").Sort _
        key1:=Worksheets("Sheet1").Range("A1")
    Set currentCell = Worksheets("Sheet1").Range("A1")
    Do While Not IsEmpty(currentCell)
        Set nextCell = currentCell.Offset(1, 0)
        If nextCell.Value = currentCell.Value Then
            currentCell.EntireRow.Delete
        End If
        Set currentCell = nextCell
    Loop
End Sub

' Do Until ... Loop
Sub DoUntilLoop()
    Dim Counter As Long
    Counter = 0
    Do Until Counter >= 20
        Counter = Counter + 1
    Loop
    Debug.Print Counter
End Sub

' Do ... Loop While
Sub DoLoopWhile()
    Dim x As Long
    x = 0
    Do
        x = x + 1
    Loop While x < 10
End Sub

' Do ... Loop Until
Sub DoLoopUntil()
    Dim x As Long
    x = 0
    Do
        x = x + 1
    Loop Until x >= 10
End Sub
