Attribute VB_Name = "IfThenElse"
Option Explicit

' Simple If
Sub SimpleIf()
    Dim x As Long
    If x = 1 Then
        Debug.Print "one"
    End If
End Sub

' If/ElseIf/Else
Sub IfElseIfElse()
    Dim x As Long
    If x = 1 Then
        Debug.Print "one"
    ElseIf x = 2 Then
        Debug.Print "two"
    Else
        Debug.Print "other"
    End If
End Sub

' Single-line If
Sub SingleLineIf()
    Dim Number As Integer, MyString As String
    Number = 1
    If Number = 1 Then GoTo Line1 Else GoTo Line2

Line1:
    MyString = "Number equals 1"
    GoTo LastLine
Line2:
    MyString = "Number equals 2"
LastLine:
    Debug.Print MyString
End Sub

' If with Excel objects
Sub IfWithExcel()
    For Each c In Worksheets("Sheet1").Range("A1:D10")
        If c.Value < .001 Then
            c.Value = 0
        End If
    Next c

    If Application.OperatingSystem Like "*Macintosh*" Then
        Application.StandardFont = "Geneva"
    Else
        Application.StandardFont = "Arial"
    End If
End Sub
