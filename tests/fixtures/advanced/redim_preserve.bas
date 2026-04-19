Attribute VB_Name = "ReDimPreserve"
Option Explicit

Sub ReDimDemo()
    Dim MyArray() As Integer
    ReDim MyArray(5)
    Dim I As Long
    For I = 1 To 5
       MyArray(I) = I
    Next I

    ReDim MyArray(10)
    For I = 1 To 10
       MyArray(I) = I
    Next I

    ReDim Preserve MyArray(15)
End Sub

Sub DynamicArrayWithErase()
    Dim NumArray(10) As Integer
    Dim StrVarArray(10) As String
    Dim StrFixArray(10) As String * 10
    Dim VarArray(10) As Variant
    Dim DynamicArray() As Integer
    ReDim DynamicArray(10)
    Erase NumArray
    Erase StrVarArray
    Erase StrFixArray
    Erase VarArray
    Erase DynamicArray
End Sub
