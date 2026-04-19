Attribute VB_Name = "NestedStructures"
Option Explicit

' Deeply nested control flow
Sub DeeplyNested()
    Dim i As Long, j As Long, k As Long
    
    For i = 1 To 10
        For j = 1 To 10
            If i > j Then
                Select Case i Mod 3
                    Case 0
                        Do While k < 5
                            With Worksheets("Sheet1")
                                .Cells(i, j).Value = k
                            End With
                            k = k + 1
                        Loop
                    Case 1
                        For k = 1 To 3
                            If k = 2 Then
                                Exit For
                            End If
                        Next k
                    Case 2
                        On Error Resume Next
                        Debug.Print 1 / (i - j)
                        On Error GoTo 0
                End Select
            End If
        Next j
    Next i
End Sub

' Nested With blocks
Sub NestedWith()
    With Worksheets("Sheet1")
        With .Range("A1")
            With .Font
                .Name = "Arial"
                .Size = 12
                .Bold = True
            End With
            .Value = "Hello"
        End With
    End With
End Sub
