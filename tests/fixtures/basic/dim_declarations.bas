Attribute VB_Name = "DimDeclarations"
Option Explicit

' Variant by default
Dim AnyValue, MyValue

' Explicitly typed
Dim Number As Integer

' Multiple declarations on single line
Dim AnotherVar, Choice As Boolean, BirthDate As Date

' Fixed-size array
Dim DayArray(50)

' Multi-dimensional array
Dim Matrix(3, 4) As Integer

' Array with explicit bounds
Dim MyMatrix(1 To 5, 4 To 9, 3 To 5) As Double

' Array with index range
Dim BirthDay(1 To 10) As Date

' Dynamic array
Dim MyArray()

' Fixed-width string
Dim FixedStr As String * 3

' Type-declaration characters
Sub TypeDeclarationChars()
    Dim perc%           ' Integer
    Dim excl!           ' Single
    Dim hash#           ' Double
    Dim dollar$         ' String
    Dim amp&            ' Long
    Dim at@             ' Currency
End Sub
