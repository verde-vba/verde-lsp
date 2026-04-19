Attribute VB_Name = "AttributeLines"
Attribute VB_Description = "Module with various attribute lines"

' Module-level attributes are parsed before Option statements
Option Explicit

' Procedure with VB_Description attribute
Public Sub DocumentedSub()
    Attribute DocumentedSub.VB_Description = "This sub does something"
    Debug.Print "Hello"
End Sub

' Procedure with VB_UserMemId (default member)
Public Property Get Item(index As Long) As Variant
    Attribute Item.VB_UserMemId = 0
End Property

' VB_MemberFlags
Public Sub HiddenSub()
    Attribute HiddenSub.VB_MemberFlags = "40"
End Sub
