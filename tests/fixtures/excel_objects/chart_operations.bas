Attribute VB_Name = "ChartOperations"
Option Explicit

Sub CreateChart()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add( _
        Left:=200, Top:=50, Width:=400, Height:=300)
    
    With cht.Chart
        .SetSourceData Source:=ws.Range("A1:B10")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Sales Data"
        
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Month"
        End With
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Revenue"
        End With
    End With
End Sub
