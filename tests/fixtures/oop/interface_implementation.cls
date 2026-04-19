VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FooImpl"
Option Explicit

Implements IExample

Public Sub IExample_aSub()
    MsgBox "Foo says hello"
End Sub

Public Function IExample_aFunc(param As Double)
    IExample_aFunc = param * 2
End Function
