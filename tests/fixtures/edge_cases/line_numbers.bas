Attribute VB_Name = "LineNumbers"
Option Explicit

' Procedures with line numbers (used with Erl)
Sub SubWithLineNumbers()
    On Error GoTo err_
    
    Dim i As Long
    i = 42 / 0

    Debug.Print "This line is not reached"
    Exit Sub

err_:
    Debug.Print "*** Error ***"
    Debug.Print "   " & Err.Description & " on line " & Erl
End Sub
