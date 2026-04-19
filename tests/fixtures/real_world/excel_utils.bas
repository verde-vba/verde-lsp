Attribute VB_Name = "ExcelUtils"
Option Explicit

Private Declare Function CallNamedPipe Lib "kernel32" _
    Alias "CallNamedPipeA" ( _
        ByVal lpNamedPipeName As String, _
        ByVal lpInBuffer As Any, ByVal nInBufferSize As Long, _
        ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, _
        ByRef lpBytesRead As Long, ByVal nTimeOut As Long) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Function IsWorkbookOpen(wbFilename As String) As Boolean
    Dim w As Workbook
    On Error GoTo notOpen
    Set w = Workbooks(wbFilename)
    IsWorkbookOpen = True
    Exit Function
notOpen:
    IsWorkbookOpen = False
End Function

Public Function SheetExists(sheetName As String, Optional wb As Workbook) _
    As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim s As Object
    On Error GoTo notFound
    Set s = wb.Sheets(sheetName)
    SheetExists = True
    Exit Function
notFound:
    SheetExists = False
End Function

Public Function ChartExists(chartName As String, _
    Optional sheetName As String = "", Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim s As Worksheet
    Dim c As ChartObject
    ChartExists = False
    If sheetName = "" Then
        For Each s In wb.Sheets
            If ChartExists(chartName, s.Name, wb) Then
                ChartExists = True
                Exit Function
            End If
        Next
    Else
        Set s = wb.Sheets(sheetName)
        On Error GoTo notFound
        Set c = s.ChartObjects(chartName)
        ChartExists = True
notFound:
    End If
End Function

Private Sub DeleteSheetOrSheets(s As Object)
    Dim prevDisplayAlerts As Boolean
    prevDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    s.Delete
    On Error GoTo 0
    Application.DisplayAlerts = prevDisplayAlerts
End Sub

Public Function GetRealUsedRange(s As Worksheet, _
    Optional fromTopLeft As Boolean = True) As Range
    If fromTopLeft Then
        Set GetRealUsedRange = s.Range( _
            s.Cells(1, 1), _
            s.Cells( _
                s.UsedRange.Rows.Count + s.UsedRange.Row - 1, _
                s.UsedRange.Columns.Count + s.UsedRange.Column - 1))
    Else
        Set GetRealUsedRange = s.UsedRange
    End If
End Function

Public Function ExcelCol(c As Integer) As String
    ExcelCol = ExcelCol_ZeroBased(c - 1)
End Function

Private Function ExcelCol_ZeroBased(c As Integer) As String
    Dim c2 As Integer
    c2 = c \ 26
    If c2 = 0 Then
        ExcelCol_ZeroBased = Chr(65 + c)
    Else
        ExcelCol_ZeroBased = ExcelCol(c2) & Chr(65 + (c Mod 26))
    End If
End Function

Public Function ExcelErrorType(e As Variant) As String
    If IsError(e) Then
        Select Case e
            Case CVErr(xlErrDiv0): ExcelErrorType = "#DIV/0!"
            Case CVErr(xlErrNA): ExcelErrorType = "#N/A"
            Case CVErr(xlErrName): ExcelErrorType = "#NAME?"
            Case CVErr(xlErrNull): ExcelErrorType = "#NULL!"
            Case CVErr(xlErrNum): ExcelErrorType = "#NUM!"
            Case CVErr(xlErrRef): ExcelErrorType = "#REF!"
            Case CVErr(xlErrValue): ExcelErrorType = "#VALUE!"
            Case Else: ExcelErrorType = "#UNKNOWN_ERROR"
        End Select
    Else
        ExcelErrorType = "(not an error)"
    End If
End Function

Public Sub ShowStatusMessage(statusMessage As String)
    Application.StatusBar = statusMessage
    Application.Caption = Len(statusMessage) & ":" & statusMessage
End Sub

Public Sub ClearStatusMessage()
    Application.StatusBar = False
    Application.Caption = Empty
End Sub

Public Sub RefreshAccessConnections(Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim cn As WorkbookConnection
    On Error GoTo err_
    Application.Calculation = xlCalculationManual
    Dim numConnections As Integer, i As Integer
    For Each cn In wb.Connections
        If cn.Type = xlConnectionTypeOLEDB Then
            numConnections = numConnections + 1
        End If
    Next
    For Each cn In wb.Connections
        If cn.Type = xlConnectionTypeOLEDB Then
            i = i + 1
            ShowStatusMessage "Refreshing data connection '" _
                & cn.OLEDBConnection.CommandText _
                & "' (" & i & " of " & numConnections & ")"
            cn.OLEDBConnection.BackgroundQuery = False
            cn.Refresh
       End If
    Next
    GoTo done_
err_:
    MsgBox "Error " & Err.Number & ": " & Err.Description
done_:
    ShowStatusMessage "Recalculating"
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    ClearStatusMessage
End Sub
