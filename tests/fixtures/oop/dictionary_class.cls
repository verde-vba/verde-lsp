VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleDictionary"
Option Explicit

Private keys() As String
Private values() As Variant
Private mCount As Long

Private Sub Class_Initialize()
    mCount = 0
    ReDim keys(0)
    ReDim values(0)
End Sub

Public Sub Add(Key As String, value As Variant)
    mCount = mCount + 1
    ReDim Preserve keys(1 To mCount)
    ReDim Preserve values(1 To mCount)
    keys(mCount) = Key
    If IsObject(value) Then
        Set values(mCount) = value
    Else
        values(mCount) = value
    End If
End Sub

Public Function Item(Key As String) As Variant
    Dim i As Long
    For i = 1 To mCount
        If keys(i) = Key Then
            If IsObject(values(i)) Then
                Set Item = values(i)
            Else
                Item = values(i)
            End If
            Exit Function
        End If
    Next i
    Err.Raise 9, "SimpleDictionary", "Key not found: " & Key
End Function

Public Function Exists(Key As String) As Boolean
    Dim i As Long
    For i = 1 To mCount
        If keys(i) = Key Then
            Exists = True
            Exit Function
        End If
    Next i
    Exists = False
End Function

Public Property Get Count() As Long
    Count = mCount
End Property

Public Sub Remove(Key As String)
    Dim i As Long, j As Long
    For i = 1 To mCount
        If keys(i) = Key Then
            For j = i To mCount - 1
                keys(j) = keys(j + 1)
                values(j) = values(j + 1)
            Next j
            mCount = mCount - 1
            If mCount > 0 Then
                ReDim Preserve keys(1 To mCount)
                ReDim Preserve values(1 To mCount)
            End If
            Exit Sub
        End If
    Next i
End Sub
