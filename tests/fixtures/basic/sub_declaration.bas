Attribute VB_Name = "SubDeclaration"
Option Explicit

' Simple Sub with no parameters
Sub AutoOpen()
End Sub

' Sub with parameters and body
Sub SubComputeArea(Length, TheWidth)
   Dim Area As Double
   If Length = 0 Or TheWidth = 0 Then
      Exit Sub
   End If
   Area = Length * TheWidth
   Debug.Print Area
End Sub

' Private Sub with typed parameter
Private Sub RunTask(ws As Worksheet)
    ws.Activate
End Sub

' Public Sub with ByVal and ByRef
Public Sub ProcessData(ByVal Factor1 As Integer, ByRef Factor2 As Integer)
    Dim Product
    Product = Factor1 * Factor2
    Debug.Print Product
End Sub
