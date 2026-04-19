Attribute VB_Name = "CallPatterns"
Option Explicit

' Call with parentheses
Sub CallWithParens()
    Call PrintToDebugWindow("Hello World")
End Sub

Sub PrintToDebugWindow(AnyString)
   Debug.Print AnyString
End Sub

' Call without Call keyword
Sub CallWithoutKeyword()
    PrintToDebugWindow "Hello World"
End Sub

' Call with named arguments
Sub CallWithNamedArgs()
    Dim loanAmt As Variant
    loanAmt = Application.InputBox _
        (Prompt:="Loan amount (100,000 for example)", _
            Default:=loanAmt, Type:=1)
End Sub

' AddressOf operator
Sub UseAddressOf()
    Dim cbAddr As Long
    cbAddr = GetAddressOfCallback(AddressOf CallBackSub)
End Sub

Function GetAddressOfCallback(addr As Long) As Long
    GetAddressOfCallback = addr
End Function

Sub CallBackSub()
    MsgBox "xyz"
End Sub

' Call Shell
Sub CallShell()
    Dim AppName As String
    Call Shell(AppName, 1)
End Sub
