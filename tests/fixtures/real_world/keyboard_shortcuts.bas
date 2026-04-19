Attribute VB_Name = "KeyboardShortcuts"
Option Explicit

Sub Auto_Open()
    Application.OnKey "^+b", "CalcBorrowingBaseShortcut"
End Sub

Sub Auto_Close()
    Application.OnKey "^+b", ""
End Sub
