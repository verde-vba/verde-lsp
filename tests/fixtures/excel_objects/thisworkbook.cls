VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Auto_Open
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Auto_Close
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Validate data before saving
    If Not ValidateAllData() Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Data validation failed. Save anyway?", _
            vbYesNo + vbExclamation, "Validation Warning")
        If result = vbNo Then Cancel = True
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Log changes across all sheets
    Debug.Print "Changed: " & Sh.Name & "!" & Target.Address & " at " & Now
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    ' Apply default formatting to new sheets
    With Sh.Rows(1).Font
        .Bold = True
        .Size = 12
    End With
End Sub

Private Function ValidateAllData() As Boolean
    ValidateAllData = True
End Function
