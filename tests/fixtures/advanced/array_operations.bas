Attribute VB_Name = "ArrayOperations"
Option Explicit

' Creating arrays with Array()
Sub CreateArray()
    Dim ary As Variant
    ary = Array("zero", "one", "two", "three", "four", "five")
    Dim i As Long
    For i = LBound(ary) To UBound(ary)
        Debug.Print i & ": " & ary(i)
    Next i
End Sub

' Declaring arrays
Sub DeclareArrays()
    Dim arr1d(5) As Long                    ' 1D fixed array
    Dim arr2d(3, 4) As Integer              ' 2D array
    Dim arrBounds(1 To 5, 4 To 9) As Double ' Explicit bounds
    Dim arrDynamic() As Integer             ' Dynamic array
    
    ReDim arrDynamic(10)
    ReDim Preserve arrDynamic(15)
End Sub

' Passing arrays to functions
Function DotProductExplicit(ByRef vec_1() As Double, ByRef vec_2() As Double) As Double
    Dim i As Long
    DotProductExplicit = 0
    For i = LBound(vec_1) To UBound(vec_1)
        DotProductExplicit = DotProductExplicit + vec_1(i) * vec_2(i)
    Next i
End Function

Function DotProductVariant(ByRef vec_1 As Variant, ByRef vec_2 As Variant) As Double
    Dim i As Long
    DotProductVariant = 0
    For i = LBound(vec_1) To UBound(vec_1)
        DotProductVariant = DotProductVariant + vec_1(i) * vec_2(i)
    Next i
End Function

' Return array from function
Function Fibonacci(n As Integer) As Long()
    Dim f() As Long
    Dim i As Integer
    ReDim f(1 To n)
    f(1) = 1
    f(2) = 1
    For i = 3 To n
        f(i) = f(i - 2) + f(i - 1)
    Next i
    Fibonacci = f
End Function

' Option Base interaction
Sub OptionBaseDemo()
    ' With default Option Base 0:
    Dim MyArray(20) As Long           ' 0 to 20
    Dim TwoDArray(3, 4) As Long       ' 0 to 3, 0 to 4
    Dim ZeroArray(0 To 5) As Long     ' Explicit override
    Dim Lower As Long
    Lower = LBound(MyArray)
    Lower = LBound(TwoDArray, 2)
    Lower = LBound(ZeroArray)
End Sub
