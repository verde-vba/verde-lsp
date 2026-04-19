Attribute VB_Name = "LineContinuation"
Option Explicit

' Line continuation with underscore
Sub LineContinuationDemo()
    Dim loanAmt As Variant
    loanAmt = Application.InputBox _
        (Prompt:="Loan amount (100,000 for example)", _
            Default:=loanAmt, Type:=1)
    
    MsgBox "The selection contains " & _
        Selection.Columns.Count & " columns."
    
    Worksheets("Sheet1").Range("A1").Sort _
        key1:=Worksheets("Sheet1").Range("A1")
End Sub
