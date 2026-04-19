Attribute VB_Name = "ForNext"
Option Explicit

' Simple For loop
Sub SimpleFor()
    Dim i As Long
    For i = 1 To 10
        Debug.Print i
    Next i
End Sub

' For with Step
Sub ForWithStep()
    Dim i As Long
    For i = 10 To 1 Step -1
        Debug.Print i
    Next i
End Sub

' For without variable after Next
Sub ForWithoutVar()
    Dim i As Long
    For i = 1 To 10
        Beep
    Next
End Sub

' For with Excel objects
Sub ForWithExcel()
    Worksheets("Sheet1").Activate
    Dim areaCount As Long
    areaCount = Selection.Areas.Count
    If areaCount <= 1 Then
        MsgBox "The selection contains " & _
            Selection.Columns.Count & " columns."
    Else
        Dim i As Long
        For i = 1 To areaCount
            MsgBox "Area " & i & " of the selection contains " & _
                Selection.Areas(i).Columns.Count & " columns."
        Next i
    End If
End Sub

' Nested For with ReDim
Sub NestedFor()
    Dim items() As Boolean
    Dim lbox As Object
    Set lbox = Worksheets("Sheet1").ListBoxes(1)
    ReDim items(1 To lbox.ListCount)
    Dim i As Long
    For i = 1 To lbox.ListCount
        If i Mod 2 = 1 Then
            items(i) = True
        Else
            items(i) = False
        End If
    Next
End Sub
