Attribute VB_Name = "DeclareAPI"
Option Explicit

' Windows API declarations
Declare Sub MessageBeep Lib "User" (ByVal N As Integer)
Declare Sub MessageBeep Lib "User" Alias "SomeBeep" (ByVal N As Integer)
Declare Function GetWinFlags Lib "Kernel" Alias "#132" () As Long

' Conditional compilation for 32/64-bit
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
#Else
    Private Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "User32" () As Long
    Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
#End If

' Win32/Win16 conditional
#If Win32 Then
   Declare Sub MessageBeep Lib "User32" (ByVal N As Long)
#Else
   Declare Sub MessageBeep Lib "User" (ByVal N As Integer)
#End If

' Call API
Sub CallMyDll()
   Call MessageBeep(0)
   MessageBeep 0   ' Call without Call keyword
End Sub
