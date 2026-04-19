VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Resource"
Option Explicit

Private member As Long

Sub Class_Initialize()
    Debug.Print "Allocating resource"
End Sub

Sub Class_Terminate()
    Debug.Print "Freeing resource"
End Sub

Sub SetMemberValue(m As Long)
    Debug.Print "Setting member value to " & m
    member = m
End Sub

Function GetMemberValue() As Long
    GetMemberValue = member
End Function
