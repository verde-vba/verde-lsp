Attribute VB_Name = "RangeOperations"
Option Explicit

' Working with Range objects
Sub RangeExamples()
    ' Single cell
    Range("A1").Value = "Hello"
    
    ' Range of cells
    Range("A1:D10").ClearContents
    
    ' Cells property
    Cells(1, 1).Value = "Row 1, Col 1"
    
    ' Offset
    Range("A1").Offset(1, 0).Value = "One row down"
    
    ' Resize
    Range("A1").Resize(5, 3).Select
    
    ' End property for finding last row
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' End property for finding last column
    Dim lastCol As Long
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' CurrentRegion
    Range("A1").CurrentRegion.Select
    
    ' SpecialCells
    Dim blankCells As Range
    On Error Resume Next
    Set blankCells = Range("A1:D10").SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
    
    ' Areas collection
    Dim area As Range
    For Each area In Selection.Areas
        Debug.Print area.Address
    Next area
    
    ' Intersect
    Dim rng As Range
    Set rng = Application.Intersect(Range("A1:C5"), Range("B3:D7"))
    If Not rng Is Nothing Then
        rng.Interior.Color = vbYellow
    End If
    
    ' Union
    Set rng = Application.Union(Range("A1:A5"), Range("C1:C5"))
    rng.Font.Bold = True
End Sub

' Working with named ranges
Sub NamedRangeExamples()
    ActiveWorkbook.Names.Add Name:="SalesData", RefersTo:="=Sheet1!$A$1:$D$100"
    Range("SalesData").Select
End Sub

' Copy/Paste operations
Sub CopyPasteExamples()
    Range("A1:D10").Copy
    Range("F1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    ' Copy to another sheet
    Worksheets("Sheet1").Range("A1:D10").Copy _
        Destination:=Worksheets("Sheet2").Range("A1")
End Sub

' Sorting
Sub SortExample()
    Dim ws As Worksheet
    Set ws = Worksheets("Sheet1")
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("A2:A100"), _
        SortOn:=xlSortOnValues, Order:=xlAscending
    
    With ws.Sort
        .SetRange ws.Range("A1:D100")
        .Header = xlYes
        .Apply
    End With
End Sub

' AutoFilter
Sub FilterExample()
    Dim ws As Worksheet
    Set ws = Worksheets("Sheet1")
    
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ws.Range("A1:D100").AutoFilter Field:=1, Criteria1:=">100"
    
    ' Count visible rows
    Dim visibleCount As Long
    visibleCount = ws.AutoFilter.Range.Columns(1) _
        .SpecialCells(xlCellTypeVisible).Count - 1
    
    ws.AutoFilterMode = False
End Sub
