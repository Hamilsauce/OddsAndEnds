POption Explicit

Sub slicerCacheTimelines()
'
' Macro3 Macro
'

'
    ActiveWorkbook.SlicerCaches("Timeline_Date1").ClearDateFilter
    ActiveWorkbook.SlicerCaches("Timeline_Date1").TimelineState.SetFilterDateRange "7/1/2019", "7/31/2019"
    ActiveWorkbook.SlicerCaches("Timeline_Date1").Slicers("Date 1"). _
        TimelineViewState.ShowHeader = True
    ActiveWorkbook.SlicerCaches("Timeline_Date1").Slicers("Date 1"). _
        TimelineViewState.ShowHorizontalScrollbar = True
    ActiveWorkbook.SlicerCaches("Timeline_Date1").Slicers("Date 1"). _
        TimelineViewState.ShowHorizontalScrollbar = False
    ActiveWorkbook.SlicerCaches("Timeline_Date1").Slicers("Date 1"). _
        TimelineViewState.ShowHeader = False
    ActiveWorkbook.SlicerCaches("Timeline_Date1").Slicers("Date 1"). _
        TimelineViewState.ShowTimeLevel = True
    ActiveSheet.Shapes.Range(Array("Low", "High", "Date 2", "Date 3", "Date 4", _
        "Percent Change")).Select

    ActiveSheet.Shapes.Range(
        Array("High", "Date 2", "Date 3", "Date 4", "Percent Change")).Select

    For Each range In ranges 
        
    Next range


    Selection.ShapeRange.IncrementLeft 126.4285826772
    Selection.ShapeRange.IncrementTop 1.0714173228
    ActiveSheet.Shapes.Range(Array("Low")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Low").PivotTables.AddPivotTable (ActiveSheet _
        .PivotTables("PivotTable4"))
    ActiveWorkbook.SlicerCaches("Slicer_Low").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable4"))
    ActiveWorkbook.SlicerCaches("Slicer_Low").PivotTables.AddPivotTable ( _
        ActiveWorkbook.PivotTables("PivotChartTable3"))
    ActiveWorkbook.SlicerCaches("Slicer_Low").VisibleSlicerItemsList = Array( _
        "[newFinalARQL].[Low].&[1.]")
    ActiveWorkbook.SlicerCaches("Slicer_Low").VisibleSlicerItemsList = Array( _
        "[newFinalARQL].[Low].&[1.04]")
    Selection.Cut
    Sheets("ARQL Historical").Select
    ActiveSheet.ChartObjects("Chart 7").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.Paste
    Sheets("Close Price Graph").Select
    ActiveSheet.Shapes.Range(Array("High")).Select
    ActiveWorkbook.SlicerCaches("Slicer_High").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable4"))
    ActiveSheet.Shapes("High").IncrementLeft 410.3571653543
    ActiveSheet.Shapes("High").IncrementTop -260.3571653543
    ActiveWorkbook.SlicerCaches("Slicer_High").VisibleSlicerItemsList = Array( _
        "[stocksHistoricalCVRS].[High].&[7.9E-1]")
    ActiveWorkbook.SlicerCaches("Slicer_High").VisibleSlicerItemsList = Array( _
        "[stocksHistoricalCVRS].[High].&[8.1E-1]")
    ActiveWorkbook.SlicerCaches("Slicer_High").VisibleSlicerItemsList = Array( _
        "[stocksHistoricalCVRS].[High].&[7.7E-1]")
    ActiveWorkbook.SlicerCaches("Slicer_High").ClearManualFilter
    ActiveSheet.Shapes("High").ScaleWidth 2.9836307962, msoFalse, _
        msoScaleFromTopLeft
    ActiveWorkbook.SlicerCaches("Slicer_High").VisibleSlicerItemsList = Array( _
        "[stocksHistoricalCVRS].[High].&[8.3E-1]")
    ActiveSheet.Shapes.Range(Array("Date 1")).Select
    ActiveWorkbook.SlicerCaches("Timeline_Date1").ClearDateFilter
    ActiveWindow.SmallScroll Down:=-9
    ActiveSheet.Shapes.Range(Array("High")).Select
    ActiveWorkbook.SlicerCaches("Slicer_High").ClearManualFilter
    ActiveSheet.Shapes.Range(Array("Percent Change")).Select
    ActiveSheet.Shapes("Percent Change").IncrementLeft 409.2857480315
    ActiveSheet.Shapes("Percent Change").IncrementTop 51.4285826772
    ActiveSheet.Shapes.Range(Array("Date 4")).Select
    ActiveSheet.Shapes("Date 4").IncrementLeft 166.0714173228
    ActiveSheet.Shapes("Date 4").IncrementTop -38.5714173228
    ActiveWorkbook.SlicerCaches("Slicer_Date2").VisibleSlicerItemsList = Array( _
        "[stocksHistoricalHSGX].[Date].&[2017-09-27T00:00:00]")
    ActiveWorkbook.SlicerCaches("Slicer_Date2").VisibleSlicerItemsList = Array( _
        "[stocksHistoricalHSGX].[Date].&[2017-10-02T00:00:00]")
    ActiveWorkbook.SlicerCaches("Slicer_Date1").VisibleSlicerItemsList = Array( _
        "[newFinalARQL].[Date].&[2017-10-02T00:00:00]")
    ActiveWorkbook.SlicerCaches("Slicer_Date1").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Date2").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Date2").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable4"))
    ActiveWorkbook.SlicerCaches("Slicer_Date2").PivotTables.AddPivotTable ( _
        ActiveWorkbook.PivotTables("PivotChartTable3"))
    ActiveWorkbook.SlicerCaches("Slicer_Date2").VisibleSlicerItemsList = Array( _
        "[stocksHistoricalHSGX].[Date].&[2017-09-29T00:00:00]")
    Sheets("Close Price Graph").Select
    ActiveSheet.Shapes.Range(Array("Date 3")).Select
    ActiveSheet.Shapes.Range(Array("Date 4")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Date2").Slicers("Date 4").Name = _
        "HSGX Date"
    With ActiveWorkbook.SlicerCaches("Slicer_Date2").Slicers("HSGX Date")
        .Caption = "Date"
        .DisplayHeader = True
        .SlicerCacheLevel.CrossFilterType = _
        xlSlicerCrossFilterShowItemsWithDataAtTop
        .SlicerCacheLevel.SortItems = xlSlicerSortDataSourceOrder
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Date2").Slicers("HSGX Date")
        .Caption = "HSGX Date"
        .DisplayHeader = True
        .SlicerCacheLevel.CrossFilterType = _
        xlSlicerCrossFilterShowItemsWithDataAtTop
        .SlicerCacheLevel.SortItems = xlSlicerSortDataSourceOrder
    End With
    ActiveSheet.Shapes.Range(Array("Date 3")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Date1").VisibleSlicerItemsList = Array( _
        "[newFinalARQL].[Date].&[2017-09-29T00:00:00]")
    ActiveWorkbook.SlicerCaches("Slicer_Date1").VisibleSlicerItemsList = Array( _
        "[newFinalARQL].[Date].&[2017-10-04T00:00:00]")
    ActiveWorkbook.SlicerCaches("Slicer_Date1").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Date1").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable4"))
    ActiveWorkbook.SlicerCaches("Slicer_Date1").VisibleSlicerItemsList = Array( _
        "[newFinalARQL].[Date].&[2017-10-04T00:00:00]")
    ActiveSheet.Shapes.Range(Array("HSGX Date")).Select
    ActiveSheet.Shapes("HSGX Date").IncrementLeft 8.5714173228
    ActiveSheet.Shapes("HSGX Date").IncrementTop -2.1428346457
    ActiveWorkbook.SlicerCaches("Slicer_Date2").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Date1").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Slicer_Date2").VisibleSlicerItemsList = Array( _
        "[stocksHistoricalHSGX].[Date].&[2017-09-28T00:00:00]")
    ActiveSheet.Shapes.Range(Array("Date 3")).Select
    ActiveSheet.Shapes("Date 3").IncrementLeft 710.3571653543
    ActiveSheet.Shapes("Date 3").IncrementTop 19.2857480315
    ActiveWorkbook.SlicerCaches("Slicer_Date").VisibleSlicerItemsList = Array( _
        "[stocksHistoricalCVRS].[Date].&[2017-09-28T00:00:00]")
    ActiveWorkbook.SlicerCaches("Slicer_Date").ClearManualFilter
    ActiveSheet.Shapes.Range(Array("Date 2")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Date").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable4"))
    ActiveWorkbook.SlicerCaches("Slicer_Date").PivotTables.AddPivotTable ( _
        ActiveWorkbook.PivotTables("PivotChartTable3"))
    ActiveWorkbook.SlicerCaches("Slicer_Date").VisibleSlicerItemsList = Array( _
        "[stocksHistoricalCVRS].[Date].&[2017-10-02T00:00:00]")
    ActiveSheet.Shapes.Range(Array("Date 3")).Select
    ActiveWorkbook.SlicerCaches("Slicer_Date1").VisibleSlicerItemsList = Array( _
        "[newFinalARQL].[Date].&[2017-09-28T00:00:00]")
    ActiveSheet.Shapes.Range(Array("High")).Select
    ActiveSheet.Shapes("High").IncrementLeft -610.7143307087
    ActiveSheet.Shapes("High").IncrementTop 545.3571653543
    ActiveWorkbook.SlicerCaches("Slicer_Date1").VisibleSlicerItemsList = Array( _
        "[newFinalARQL].[Date].&[2017-10-04T00:00:00]")
    ActiveWorkbook.SlicerCaches("Slicer_Date1").ClearManualFilter
    Range("J9").Select
    Sheets("HSGX History").Select
    Range("A1:I7").Select
    Range("I7").Activate
    Selection.ClearContents
    Sheets("HSGX History").Select
    Range("C1").Select
    ActiveWindow.SmallScroll Down:=-6
    Application.ActiveProtectedViewWindow.Edit
    Sheets("Learn More").Select
    ActiveSheet.Shapes.Range(Array("TextBox 29")).Select
    Range("C9").Select
    Selection.ShapeRange.Item(1).Hyperlink.Follow NewWindow:=False, AddHistory _
        :=True
    Selection.ShapeRange.Item(1).Hyperlink.Follow NewWindow:=False, AddHistory _
        :=True
    ActiveWindow.Close
    Sheets("Sample").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Sample").Select
    ActiveWindow.SelectedSheets.Delete
    Range("A1").Select
    Windows("ARQL graph with View Switch.xlsm").Activate
    Windows("ARQL graph with View Switch.xlsm").Activate
    Range("L25").Select
    ActiveWindow.Close
    ActiveWindow.Close
    ActiveWindow.Close
    ActiveWindow.Close
    Range("B9").Select
    ActiveCell.FormulaR1C1 = "Input:"
    Range("A9").Select
    Selection.Delete Shift:=xlToLeft
    Range("A9").Select
    Selection.Font.Bold = True
    Range("A9").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A9").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub wsArray(dfs)
Dim sheetGroup() As Variant

Dim groupedSheets As Worksheets
Dim groupSh As Variant

Set groupSh = ThisWorkbook.Worksheets(Array("Sheet11", "Sheet9", "Sheet10"))
'Set groupSh = Sheets(Array("Sheet11", "Sheet9", "Sheet10"))

    With groupSh(0).Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With

End Sub

Sub Macro5()
'
' Macro5 Macro
'

'
    Sheets(Array("Sheet8", "Sheet9", "Sheet10")).Select
    Sheets("Sheet10").Activate
    Range("E5").Select
    ActiveCell.FormulaR1C1 = "asas"
    Range("E6").Select
    Sheets(Array("Sheet8", "Sheet9", "Sheet10")).Select
    Sheets("Sheet8").Activate
    Range("E5").Select
    ActiveWorkbook.Names.Add Name:="NAMEDRANGE", RefersToR1C1:="=Sheet8!R5C5"

    Sheets(Array("Sheet8", "Sheet9", "Sheet10")).Select
    Sheets("Sheet10").Activate
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "=LEN(Sheet9!R[-3]C[-10])"

    Sheets(Array("Sheet8", "Sheet9", "Sheet10")).Select
    Sheets("Sheet9").Activate
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "asda"
    Range("A2").Select
    Sheets(Array("Sheet8", "Sheet9", "Sheet10")).Select
    Sheets("Sheet8").Activate
    With ActiveWorkbook.Sheets("Sheet8").Tab
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
    End With
    With ActiveWorkbook.Sheets("Sheet10").Tab
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
    End With
    With ActiveWorkbook.Sheets("Sheet9").Tab
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
    End With
    Sheets("Sheet9").Select
    Range("M30").Select
    Sheets("Sheet10").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "55"
    Range("A2").Select
    Sheets("Sheet8").Select
    Sheets("Sheet8").Select
    Sheets.Add
End Sub
