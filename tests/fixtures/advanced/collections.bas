Attribute VB_Name = "Collections"
Option Explicit
Option Base 1

' Returns True if the Collection contains the specified key
Public Function HasKey(Key As Variant, col As Collection) As Boolean
    Dim obj As Variant
    On Error GoTo err
    HasKey = True
    obj = col(Key)
    Exit Function
err:
    HasKey = False
End Function

' Returns True if the Collection contains an element equal to value
Public Function Contains(value As Variant, col As Collection) As Boolean
    Contains = (IndexOf(value, col) >= 0)
End Function

' Returns the first index of an element equal to value
Public Function IndexOf(value As Variant, col As Collection) As Long
    Dim index As Long
    For index = 1 To col.Count Step 1
        If col(index) = value Then
            IndexOf = index
            Exit Function
        End If
    Next index
    IndexOf = -1
End Function

' Returns an array which exactly matches this collection
Public Function ToArray(col As Collection) As Variant
    Dim A() As Variant
    ReDim A(0 To col.Count)
    Dim i As Long
    For i = 0 To col.Count - 1
        A(i) = col(i + 1)
    Next i
    ToArray = A()
End Function

' Returns a Collection from an Array
Public Function FromArray(A() As Variant) As Collection
    Dim col As Collection
    Set col = New Collection
    Dim element As Variant
    For Each element In A
        col.Add element
    Next element
    Set FromArray = col
End Function
