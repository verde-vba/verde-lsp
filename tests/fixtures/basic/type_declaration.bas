Attribute VB_Name = "TypeDeclaration"

' User-defined type
Type EmployeeRecord
   ID As Integer
   Name As String * 20
   Address As String * 30
   Phone As Long
   HireDate As Date
End Type

' Public type with nested types
Public Type MyType
    MyProperty As String
End Type

' Using a UDT
Sub CreateRecord()
   Dim MyRecord As EmployeeRecord
   MyRecord.ID = 12003
End Sub
