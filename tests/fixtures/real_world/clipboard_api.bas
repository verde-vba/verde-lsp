Attribute VB_Name = "Clipboard"
Option Compare Database
Option Explicit

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

Function GetText()
   Dim hClipMemory As LongPtr
   Dim lpClipMemory As LongPtr
   Dim retval As LongPtr
   Dim MyString As String

   If OpenClipboard(0&) = 0 Then
      Err.Raise 5, "Clipboard", "Cannot open Clipboard. It may already be open."
   End If

   hClipMemory = GetClipboardData(1)
   If IsNull(hClipMemory) Then
      CloseClipboard
      GetText = ""
      Exit Function
   End If

   lpClipMemory = GlobalLock(hClipMemory)
   If Not IsNull(lpClipMemory) Then
      Dim nSize As Long
      nSize = GlobalSize(hClipMemory)
      MyString = Space$(nSize)
      retval = lstrcpy(MyString, lpClipMemory)
      retval = GlobalUnlock(hClipMemory)
   End If

   retval = CloseClipboard
   GetText = MyString
End Function
