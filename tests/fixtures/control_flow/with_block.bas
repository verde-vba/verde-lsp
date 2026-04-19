Attribute VB_Name = "WithBlock"
Option Explicit

' Basic With block
Sub BasicWith()
    Dim SomeObject As MyType
    With SomeObject
        .MyProperty = "SomeValue"
    End With
End Sub

' With block with Excel objects
Sub WithExcel()
    With Worksheets("Sheet1").Range("A1").CurrentRegion
        .Rows(1).Font.Bold = True
        .Columns(1).Font.Bold = True
        .Columns.AutoFit
    End With
End Sub
