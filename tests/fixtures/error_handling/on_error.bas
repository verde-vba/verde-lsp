Attribute VB_Name = "OnError"
Option Explicit

Sub OnErrorStatementDemo()
   On Error GoTo ErrorHandler
   Open "TESTFILE" For Output As #1
   Kill "TESTFILE"

   On Error GoTo 0   ' Turn off error trapping
   On Error Resume Next   ' Defer error trapping
   Dim ObjectRef As Object
   Set ObjectRef = GetObject("MyWord.Basic")

   If Err.Number = 440 Or Err.Number = 432 Then
      Dim Msg As String
      Msg = "There was an error attempting to open the Automation object!"
      MsgBox Msg, , "Deferred Error Test"
      Err.Clear
   End If
Exit Sub
ErrorHandler:
   Select Case Err.Number
      Case 55
         Close #1
      Case Else
         ' Handle other situations here...
   End Select
   Resume
End Sub
