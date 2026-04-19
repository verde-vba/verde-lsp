VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyProcedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CurrentColor As Integer
Private Const BLACK = 0, RED = 1, GREEN = 2, BLUE = 3

' Property Get - returns value
Property Get PenColor() As String
   Select Case CurrentColor
      Case RED
         PenColor = "Red"
      Case GREEN
         PenColor = "Green"
      Case BLUE
         PenColor = "Blue"
   End Select
End Property

' Property Let - sets value
Property Let PenColor(ColorName As String)
   Select Case ColorName
      Case "Red"
         CurrentColor = RED
      Case "Green"
         CurrentColor = GREEN
      Case "Blue"
         CurrentColor = BLUE
      Case Else
         CurrentColor = BLACK
   End Select
End Property

' Property Set - sets object reference
Property Set Pen(P As Object)
   Set CurrentPen = P
End Property

' Property with default member attribute
Private m_name As String

Property Let Name(name_ As String)
    Attribute Name.VB_UserMemId = 0
    m_name = name_
End Property

Property Get Name() As String
    Name = m_name
End Property
