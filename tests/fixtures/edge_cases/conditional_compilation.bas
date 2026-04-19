Attribute VB_Name = "ConditionalCompilation"
Option Explicit

' Basic conditional compilation
#If VBA7 Then
    Debug.Print "VBA7 is defined"
#Else
    Debug.Print "VBA7 is not defined"
#End If

' Multiple ElseIf
#If Win64 Then
    Debug.Print "win64 is defined"
#ElseIf Win32 Then
    Debug.Print "win32 is defined"
#Else
    Debug.Print "neither win64 nor win32 is defined"
#End If

' Conditional compilation with 0 (dead code)
Sub DeadCodeExample()
    Debug.Print "one"
#If 0 Then
    Debug.Print "two"
    Debug.Print "three"
    Debug.Print "four"
#End If
    Debug.Print "five"
End Sub

' Conditional API declarations
#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

' Custom compilation constants
#Const DEBUG_MODE = 1
#If DEBUG_MODE Then
    Sub DebugOutput(msg As String)
        Debug.Print msg
    End Sub
#End If
