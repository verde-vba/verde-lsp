Attribute VB_Name = "FileIO"
Option Explicit

Sub FileOperations()
    ' Open for various modes
    Open "TESTFILE" For Input As #1
    Close #1

    Open "TESTFILE" For Binary Access Write As #1
    Close #1

    Type Record
       ID As Integer
       Name As String * 20
    End Type

    Dim MyRecord As Record
    Open "TESTFILE" For Random As #1 Len = Len(MyRecord)
    Close #1

    Open "TESTFILE" For Output Shared As #1
    Close #1

    Open "TESTFILE" For Binary Access Read Lock Read As #1
    Close #1
End Sub

' Reading from file
Sub ReadFile()
    Dim MyString As String, MyNumber As Long
    Open "TESTFILE" For Input As #1
    Do While Not EOF(1)
       Input #1, MyString, MyNumber
       Debug.Print MyString, MyNumber
    Loop
    Close #1
End Sub

' Line Input
Sub ReadFileLines()
    Dim TextLine As String
    Open "TESTFILE" For Input As #1
    Do While Not EOF(1)
       Line Input #1, TextLine
       Debug.Print TextLine
    Loop
    Close #1
End Sub

' Get/Put for random access
Sub RandomAccess()
    Type Record
       ID As Integer
       Name As String * 20
    End Type

    Dim MyRecord As Record, RecordNumber As Long
    Open "TESTFILE" For Random As #1 Len = Len(MyRecord)
    For RecordNumber = 1 To 5
       MyRecord.ID = RecordNumber
       MyRecord.Name = "My Name" & RecordNumber
       Put #1, RecordNumber, MyRecord
    Next RecordNumber
    Close #1
End Sub

' Lock/Unlock
Sub LockUnlockDemo()
    Type Record
       ID As Integer
       Name As String * 20
    End Type

    Dim MyRecord As Record, RecordNumber As Long
    Open "TESTFILE" For Random Shared As #1 Len = Len(MyRecord)
    RecordNumber = 4
    Lock #1, RecordNumber
    Get #1, RecordNumber, MyRecord
    MyRecord.ID = 234
    MyRecord.Name = "John Smith"
    Put #1, RecordNumber, MyRecord
    Unlock #1, RecordNumber
    Close #1
End Sub

' Write statement
Sub WriteDemo()
    Open "TESTFILE" For Output As #1
    Write #1, "Hello World", 234
    Write #1,   ' Write blank line

    Dim MyBool As Boolean, MyDate As Date, MyNull As Variant, MyError As Variant
    MyBool = False: MyDate = #February 12, 1969#: MyNull = Null
    MyError = CVErr(32767)
    Write #1, MyBool; " is a Boolean value"
    Write #1, MyDate; " is a date"
    Write #1, MyNull; " is a null value"
    Write #1, MyError; " is an error value"
    Close #1
End Sub
