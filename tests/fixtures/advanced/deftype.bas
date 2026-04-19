Attribute VB_Name = "Deftype"

' Variable names beginning with A through K default to Integer
DefInt A-K
' Variable names beginning with L through Z default to String
DefStr L-Z

Dim CalcVar
CalcVar = 4
Dim StringVar
StringVar = "Hello there"

Function ATestFunction(INumber)
   ATestFunction = INumber * 2
End Function
