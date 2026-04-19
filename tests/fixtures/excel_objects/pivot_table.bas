Attribute VB_Name = "PivotTableOps"
Option Explicit

Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pc As PivotCache
    Dim ptField As PivotField
    
    Set ws = ActiveSheet
    
    ' Create pivot cache
    Set pc = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=ws.Range("A1").CurrentRegion.Address)
    
    ' Create pivot table
    Set pt = pc.CreatePivotTable( _
        TableDestination:=Worksheets("PivotSheet").Range("A3"), _
        TableName:="SalesPivot")
    
    ' Configure pivot fields
    With pt
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Product").Orientation = xlColumnField
        .AddDataField .PivotFields("Revenue"), "Sum of Revenue", xlSum
        .AddDataField .PivotFields("Units"), "Sum of Units", xlSum
    End With
End Sub

Sub RefreshAllPivots()
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
End Sub
