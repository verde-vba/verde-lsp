Attribute VB_Name = "ResumePatterns"
Option Explicit

Sub ResumeStatementDemo()
   On Error GoTo ErrorHandler
   Open "TESTFILE" For Output As #1
   Kill "TESTFILE"
   Exit Sub
ErrorHandler:
   Select Case Err.Number
      Case 55
         Close #1
      Case Else
         ' Handle other situations here....
   End Select
   Resume   ' Resume execution at same line that caused the error
End Sub

' Error handling with Resume Next
Sub SafeWorksheetLookup()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    On Error GoTo 0
End Sub

' Error handling with line numbers and Erl
Sub ErrorWithLineNumbers()
    On Error GoTo err_
    Dim i As Long
    i = 42 / 0   ' Division by Zero!
    Debug.Print "This line is not reached"
    Exit Sub
err_:
    Debug.Print "*** Error ***"
    Debug.Print "   " & Err.Description & " on line " & Erl
End Sub

' Complex error handling with multiple handlers
Sub ComplexErrorHandling()
    On Error GoTo err_C
    
    Dim a As Integer, b As Integer, c As Integer
    a = 5 - 3
    b = a - 2
    c = a / b  ' Boom!
    
done_C:
    Exit Sub
    
err_C:
    MsgBox "Error: " & Err.Description & " [" & Err.Number & "]"
    Resume done_C
    Resume   ' For debugging: use Set Next Statement here
End Sub
