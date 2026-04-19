Attribute VB_Name = "StringOperations"

' Mid statement (not function - left-side assignment)
Dim MyString As String
MyString = "The dog jumps"
Mid(MyString, 5, 3) = "fox"
Mid(MyString, 5) = "cow"
Mid(MyString, 5) = "cow jumped over"
Mid(MyString, 5, 3) = "duck"

' LSet and RSet
Dim LeftStr As String
LeftStr = "0123456789"
LSet LeftStr = "<-Left"

Dim RightStr As String
RightStr = "0123456789"
RSet RightStr = "Right->"

' Name statement (rename file)
Dim OldName As String, NewName As String
OldName = "OLDFILE": NewName = "NEWFILE"
Name OldName As NewName
