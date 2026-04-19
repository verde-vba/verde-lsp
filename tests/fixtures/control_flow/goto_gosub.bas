Attribute VB_Name = "GotoGosub"

' GoTo statement
Sub GotoStatementDemo()
    Dim Number As Long, MyString As String
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

' GoSub/Return
Sub GosubDemo()
    Dim Num As Variant
    Num = InputBox("Enter a positive number to be divided by 2.")
    If Num > 0 Then GoSub MyRoutine
    Debug.Print Num
    Exit Sub
MyRoutine:
    Num = Num / 2
    Return
End Sub

' On...GoSub and On...GoTo
Sub OnGosubGotoDemo()
    Dim Number As Long, MyString As String
    Number = 2
    On Number GoSub Sub1, Sub2
    On Number GoTo Line1, Line2
    Exit Sub
Sub1:
    MyString = "In Sub1": Return
Sub2:
    MyString = "In Sub2": Return
Line1:
    MyString = "In Line1"
Line2:
    MyString = "In Line2"
End Sub
