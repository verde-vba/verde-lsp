VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Predeclared"
Attribute VB_PredeclaredId = True
Option Explicit

Private member As Long

Sub Class_Initialize()
    Debug.Print "An instance of predeclared is being initialized"
End Sub

Sub Class_Terminate()
    Debug.Print "Terminate predeclared instance, member = " & member
End Sub

Sub SetMemberValue(m As Long)
    Debug.Print "setting member value to " & m
    member = m
End Sub

Function GetMemberValue() As Long
    GetMemberValue = member
End Function
