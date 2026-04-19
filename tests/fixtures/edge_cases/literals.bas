Attribute VB_Name = "Literals"
Option Explicit

Sub LiteralTypes()
    Dim x As Variant

    ' Integer literal
    x = 42

    ' Hex literal
    x = &HFF
    x = &HE1E4FF&

    ' Octal literal
    x = &O77

    ' Float literal
    x = 3.14
    x = 0.75

    ' String literal
    x = "hello world"
    x = "He said ""hello"""   ' Escaped double quote
    x = ""                     ' Empty string

    ' Boolean literals
    x = True
    x = False

    ' Nothing literal
    Set x = Nothing

    ' Null
    x = Null

    ' Empty (default Variant value)
    ' x is implicitly Empty when Dim'd as Variant

    ' Date literals
    Dim d As Date
    d = #February 12, 1969#
    d = #1994-07-13#

    ' Special characters in strings
    Dim kappa As String
    kappa = ChrW(&H3BA)
End Sub
