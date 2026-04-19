VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Check if the change was in column A
    If Not Intersect(Target, Me.Range("A:A")) Is Nothing Then
        Application.EnableEvents = False
        
        Dim cell As Range
        For Each cell In Intersect(Target, Me.Range("A:A"))
            ' Auto-uppercase column A
            If Not IsEmpty(cell.Value) Then
                cell.Value = UCase(cell.Value)
            End If
        Next cell
        
        Application.EnableEvents = True
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Highlight the active row
    Cells.Interior.ColorIndex = xlNone
    Target.EntireRow.Interior.Color = RGB(230, 240, 255)
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Me.Range("B:B")) Is Nothing Then
        MsgBox "You double-clicked on: " & Target.Value
        Cancel = True
    End If
End Sub

Private Sub Worksheet_Calculate()
    ' Runs after sheet recalculation
    Debug.Print "Sheet recalculated at " & Now
End Sub

Private Sub Worksheet_Activate()
    Debug.Print "Sheet activated"
End Sub

Private Sub Worksheet_Deactivate()
    Debug.Print "Sheet deactivated"
End Sub
