Attribute VB_Name = "ScopeAndLifetime"
Option Explicit

' Module-level variables
Public globalVar As String
Private moduleVar As Long
Dim defaultVar As Double

' Constants at module level
Public Const APP_NAME = "MyApp"
Private Const MAX_ITEMS As Long = 100

Sub DemoScope()
    ' Local variable
    Dim localVar As String
    localVar = "only visible here"
    
    ' Static variable (persists between calls)
    Static callCount As Long
    callCount = callCount + 1
    
    ' Block-level (VBA does NOT have block scope, 
    ' variable is visible throughout the entire procedure)
    If True Then
        Dim blockVar As String
        blockVar = "visible everywhere in this Sub"
    End If
    Debug.Print blockVar  ' This works in VBA
End Sub

' Procedure-level static
Static Function GetNextID() As Long
    Static id As Long
    id = id + 1
    GetNextID = id
End Function
