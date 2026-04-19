Attribute VB_Name = "SelectCase"
Option Explicit

' Select Case with Is
Sub SelectCaseWithIs()
    Select Case Application.MailSystem
        Case Is = xlMAPI
            MsgBox "Mail system is Microsoft Mail"
        Case Is = xlPowerTalk
            MsgBox "Mail system is PowerTalk"
        Case Is = xlNoMailSystem
            MsgBox "No mail system installed"
    End Select
End Sub

' Select Case with error values
Sub SelectCaseErrors()
    Worksheets("Sheet1").Activate
    If IsError(ActiveCell.Value) Then
        Dim errval As Variant
        errval = ActiveCell.Value
        Select Case errval
            Case CVErr(xlErrDiv0)
                MsgBox "#DIV/0! error"
            Case CVErr(xlErrNA)
                MsgBox "#N/A error"
            Case CVErr(xlErrName)
                MsgBox "#NAME? error"
            Case CVErr(xlErrNull)
                MsgBox "#NULL! error"
            Case CVErr(xlErrNum)
                MsgBox "#NUM! error"
            Case CVErr(xlErrRef)
                MsgBox "#REF! error"
            Case CVErr(xlErrValue)
                MsgBox "#VALUE! error"
            Case Else
                MsgBox "Unknown error"
        End Select
    End If
End Sub

' Select Case with multiple values per case
Sub SelectCaseMultiple()
    Dim x As Long
    Select Case x
        Case 1, 2, 3
            Debug.Print "1-3"
        Case 4 To 10
            Debug.Print "4-10"
        Case Is > 10
            Debug.Print ">10"
        Case Else
            Debug.Print "other"
    End Select
End Sub

' Nested Select Case in Exit
Sub ExitStatementDemo()
    Dim I As Long, MyNum As Long
    Do
        For I = 1 To 1000
            MyNum = Int(Rnd * 1000)
            Select Case MyNum
                Case 7: Exit For
                Case 29: Exit Do
                Case 54: Exit Sub
            End Select
        Next I
    Loop
End Sub
