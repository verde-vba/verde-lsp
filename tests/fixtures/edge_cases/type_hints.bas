Attribute VB_Name = "TypeHints"
Option Explicit

' Type declaration characters (suffix type hints)
Sub TypeHintDemo()
    Dim intVar%      ' Integer
    Dim longVar&     ' Long
    Dim singleVar!   ' Single
    Dim doubleVar#   ' Double
    Dim currVar@     ' Currency
    Dim strVar$      ' String
    
    intVar% = 42
    longVar& = 100000
    singleVar! = 3.14
    doubleVar# = 3.14159265358979
    currVar@ = 19.99
    strVar$ = "Hello"
End Sub
