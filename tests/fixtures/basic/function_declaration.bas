Attribute VB_Name = "FunctionDeclaration"
Option Explicit

' Function with return type
Function CalculateSquareRoot(NumberArg As Double) As Double
   If NumberArg < 0 Then
      Exit Function
   Else
      CalculateSquareRoot = Sqr(NumberArg)
   End If
End Function

' Function with ParamArray
Function CalcSum(ByVal FirstArg As Integer, ParamArray OtherArgs())
    Dim ReturnValue
    ReturnValue = CalcSum(4, 3, 2, 1)
End Function

' Public Function with ByVal parameter and return type
Public Function GetName(ByVal id As Long) As String
    GetName = "test"
End Function

' Function returning an array
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

' Function with New in Dim
Function MakeList() As Collection
    Dim result As New Collection
    Set MakeList = result
End Function
