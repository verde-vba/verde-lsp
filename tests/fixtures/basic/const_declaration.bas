Attribute VB_Name = "ConstDeclaration"

' Constants are Private by default
Const MyVar = 459

' Public constant
Public Const MyString = "HELP"

' Private Integer constant
Private Const MyInt As Integer = 5

' Multiple constants on same line
Const MyStr = "Hello", MyDouble As Double = 3.4567

' Named constants used in practice
Private Const BLACK = 0, RED = 1, GREEN = 2, BLUE = 3
Private Const INVENTORY_RATE As Double = 0.75
Private Const AR_RATE As Double = 0.8
Private Const INVENTORY_PREFIX As String = "120"
