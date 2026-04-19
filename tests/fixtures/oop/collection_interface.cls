VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICollection"

Option Explicit

Public Function Count() As Long
End Function

Public Function Contains(value As Variant) As Boolean
End Function

Public Sub Clear()
End Sub

Public Function NewEnum() As IUnknown
    Attribute NewEnum.VB_UserMemId = -4
End Function
