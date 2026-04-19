Attribute VB_Name = "LetSetNew"

' Let statement (explicit and implicit)
Dim MyStr As String, MyInt As Long
Let MyStr = "Hello World"
Let MyInt = 5

' Implicit Let (no keyword)
MyStr = "Hello World"
MyInt = 5

' Set statement for objects
Sub SetDemo()
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.Add
    newSheet.Name = "1995 Budget"
    
    ' Set with New
    Dim col As Collection
    Set col = New Collection
    
    ' Set to Nothing
    Set col = Nothing
End Sub
