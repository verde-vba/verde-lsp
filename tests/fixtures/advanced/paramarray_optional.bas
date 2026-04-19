Attribute VB_Name = "ParamArrayOptional"
Option Explicit

' Function with ParamArray
Function CalcSum(ByVal FirstArg As Integer, ParamArray OtherArgs())
    Dim ReturnValue As Variant
    ReturnValue = FirstArg
    Dim i As Long
    For i = LBound(OtherArgs) To UBound(OtherArgs)
        ReturnValue = ReturnValue + OtherArgs(i)
    Next i
    CalcSum = ReturnValue
End Function

' Function with Optional parameters
Public Function SheetExists(sheetName As String, Optional wb As Workbook) _
    As Boolean

    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim s As Object
    On Error GoTo notFound
    Set s = wb.Sheets(sheetName)
    SheetExists = True
    Exit Function

notFound:
    SheetExists = False
End Function

' Optional with default value
Function FormatValue(val As Double, Optional decimals As Integer = 2) As String
    FormatValue = Format(val, "0." & String(decimals, "0"))
End Function

' ByVal and ByRef parameters
Function DotProduct(ByRef vec_1() As Double, ByRef vec_2() As Double) As Double
    Dim i As Long
    DotProduct = 0
    For i = LBound(vec_1) To UBound(vec_1)
        DotProduct = DotProduct + vec_1(i) * vec_2(i)
    Next i
End Function
