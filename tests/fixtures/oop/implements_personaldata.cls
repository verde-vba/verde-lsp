VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomerData"
Option Explicit

Implements PersonalData

Private Property Get PersonalData_Address() As String
    PersonalData_Address = "CustomerAddress"
End Property

Private Property Let PersonalData_Address(ByVal RHS As String)
    '
End Property

Private Property Let PersonalData_Name(ByVal RHS As String)
    '
End Property

Private Property Get PersonalData_Name() As String
    PersonalData_Name = "CustomerName"
End Property
