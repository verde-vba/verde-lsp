Attribute VB_Name = "BorrowingBase"
Option Explicit

Private Const INVENTORY_RATE As Double = 0.75
Private Const AR_RATE As Double = 0.8
Private Const INVENTORY_PREFIX As String = "120"

Public Sub CalcBorrowingBase(control As IRibbonControl)
    RunBorrowingBase
End Sub

Public Sub CalcBorrowingBaseShortcut()
    RunBorrowingBase
End Sub

Private Sub RunBorrowingBase()
    Dim wsBalance As Worksheet
    Dim wsAR As Worksheet
    Dim inventoryTotal As Double
    Dim arTotal As Double
    Dim borrowingBase As Double

    Set wsBalance = FindSheet("BalanceSheet")
    If wsBalance Is Nothing Then
        MsgBox "Sheet 'BalanceSheet' not found in the active workbook.", vbExclamation
        Exit Sub
    End If

    Set wsAR = FindSheet("ARAgingSummary")
    If wsAR Is Nothing Then
        MsgBox "Sheet 'ARAgingSummary' not found in the active workbook.", vbExclamation
        Exit Sub
    End If

    inventoryTotal = CalcInventory(wsBalance)
    arTotal = CalcEligibleAR(wsAR)

    If inventoryTotal < 0 Or arTotal < 0 Then
        Exit Sub
    End If

    Dim inventoryComponent As Double
    Dim arComponent As Double
    inventoryComponent = inventoryTotal * INVENTORY_RATE
    arComponent = arTotal * AR_RATE
    borrowingBase = inventoryComponent + arComponent

    Dim msg As String
    msg = "Borrowing Base Calculation" & vbCrLf & vbCrLf
    msg = msg & "Inventory (120x accounts): " & FormatAsCurrency(inventoryTotal) & vbCrLf
    msg = msg & "  x " & Format(INVENTORY_RATE, "0%") & " = " & FormatAsCurrency(inventoryComponent) & vbCrLf & vbCrLf
    msg = msg & "Eligible AR (Current + 30 + 60): " & FormatAsCurrency(arTotal) & vbCrLf
    msg = msg & "  x " & Format(AR_RATE, "0%") & " = " & FormatAsCurrency(arComponent) & vbCrLf & vbCrLf
    msg = msg & "Borrowing Base: " & FormatAsCurrency(borrowingBase)

    MsgBox msg, vbInformation, "Borrowing Base"
End Sub

Private Function CalcInventory(ws As Worksheet) As Double
    Dim matchedRows As Collection
    Dim total As Double
    Dim rowNum As Variant

    Set matchedRows = FindRowsByPrefix(ws, 1, INVENTORY_PREFIX)

    If matchedRows.Count = 0 Then
        MsgBox "No inventory accounts (prefix '" & INVENTORY_PREFIX & "') found on BalanceSheet.", vbExclamation
        CalcInventory = -1
        Exit Function
    End If

    total = 0
    For Each rowNum In matchedRows
        Dim cellVal As Variant
        cellVal = ws.Cells(CLng(rowNum), 2).Value
        If IsNumeric(cellVal) Then
            total = total + CDbl(cellVal)
        End If
    Next rowNum

    CalcInventory = total
End Function

Private Function CalcEligibleAR(ws As Worksheet) As Double
    Dim totalRow As Long
    Dim arCurrent As Double
    Dim ar30 As Double
    Dim ar60 As Double

    totalRow = FindTotalRow(ws, 1)

    If totalRow = 0 Then
        MsgBox "Could not find 'Total' row on ARAgingSummary.", vbExclamation
        CalcEligibleAR = -1
        Exit Function
    End If

    Dim val As Variant
    val = ws.Cells(totalRow, 2).Value
    arCurrent = IIf(IsNumeric(val), CDbl(val), 0)
    val = ws.Cells(totalRow, 3).Value
    ar30 = IIf(IsNumeric(val), CDbl(val), 0)
    val = ws.Cells(totalRow, 4).Value
    ar60 = IIf(IsNumeric(val), CDbl(val), 0)

    CalcEligibleAR = arCurrent + ar30 + ar60
End Function
