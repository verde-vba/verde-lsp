Attribute VB_Name = "EmptyProcedures"
Option Explicit

' Empty Sub
Sub EmptySub()
End Sub

' Empty Function
Function EmptyFunc() As Variant
End Function

' Empty Property Get
Property Get EmptyProp() As String
End Property

' Empty Property Let
Property Let EmptyProp(val As String)
End Property

' Empty Property Set
Property Set EmptyObj(val As Object)
End Property

' Single-line Sub/Function
Public Sub SingleLine(): End Sub
Public Function SingleLineFunc() As Long: SingleLineFunc = 0: End Function
