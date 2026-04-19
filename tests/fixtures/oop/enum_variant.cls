VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumVARIANT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public interface_IEnumVARIANT As IEnumVARIANT

Public Function get_IEnumVARIANT() As IEnumVARIANT
    Attribute get_IEnumVARIANT.VB_UserMemId = -4
    Set get_IEnumVARIANT = interface_IEnumVARIANT
End Function
