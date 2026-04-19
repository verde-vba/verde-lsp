Attribute VB_Name = "MultiStatementLine"

' Multiple statements on one line with colon
Dim MyStr1, MyStr2
MyStr1 = "Hello": Rem Comment after a statement separated by a colon
MyStr2 = "Goodbye"    ' This is also a comment; no colon is needed

' Single-line If with colon
Sub MultiStatementDemo()
    Dim OldName As String, NewName As String
    OldName = "OLDFILE": NewName = "NEWFILE"
    Name OldName As NewName
End Sub

' GoSub targets on same line
Sub GoSubMultiLine()
    Dim MyString As String
    On 2 GoSub Sub1, Sub2
    Exit Sub
Sub1:
    MyString = "In Sub1": Return
Sub2:
    MyString = "In Sub2": Return
End Sub
