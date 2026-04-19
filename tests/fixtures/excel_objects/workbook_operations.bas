Attribute VB_Name = "WorkbookOperations"
Option Explicit

' Create and manage workbooks
Sub WorkbookManagement()
    ' Create new workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    ' Add worksheets
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(Type:=xlWorksheet)
    ws.Name = "DataSheet"
    
    ' Save workbook
    wb.SaveAs Filename:="C:\Temp\TestWorkbook.xlsx", _
        FileFormat:=xlOpenXMLWorkbook
    
    ' Close workbook
    wb.Close SaveChanges:=True
    
    ' Open workbook
    Set wb = Workbooks.Open("C:\Temp\TestWorkbook.xlsx")
    
    ' Iterate sheets
    For Each ws In wb.Worksheets
        Debug.Print ws.Name
    Next ws
End Sub

' Application object usage
Sub ApplicationExamples()
    ' Screen updating
    Application.ScreenUpdating = False
    
    ' Calculation
    Application.Calculation = xlCalculationManual
    
    ' Status bar
    Application.StatusBar = "Processing..."
    
    ' Display alerts
    Application.DisplayAlerts = False
    
    ' Do work here...
    
    ' Restore settings
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

' Find and Replace
Sub FindReplaceExample()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim found As Range
    Set found = ws.Cells.Find( _
        What:="SearchText", _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
    
    If Not found Is Nothing Then
        Dim firstAddress As String
        firstAddress = found.Address
        Do
            found.Interior.Color = vbYellow
            Set found = ws.Cells.FindNext(found)
        Loop While Not found Is Nothing And found.Address <> firstAddress
    End If
End Sub
