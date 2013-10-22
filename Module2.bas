Attribute VB_Name = "Module2"
'Graph result
Sub Graph(exportChart As Boolean, result As Worksheet, scl As Double, topAsBottom As Boolean)
    'modified data for graphing
    Dim count As Integer, plotRange As Range, force As Boolean
    Set plotRange = result.Range("Force").Cells(1, 1).Offset(0, 1)
'    temp = MsgBox(Prompt:="make a force graph?(instead of displacement)", _
'        Buttons:=vbYesNo, Title:="Force vs. Displacement")
'    If temp = vbYes Then
'        force = True
'    End If
    Call prepareData(result, count, force, topAsBottom)
    'sort data for graphing
    Dim col As Range, sortRange As Range, lastRow As Range
    Set col = result.Range(plotRange, plotRange.Offset(3 * count - 1))
    'select sortRange
    Set sortRange = result.Range(plotRange.Offset(0), plotRange.Offset(3 * count - 1, 2))

    'Sort data
    With result.Sort
    .SortFields.Clear
    .SortFields.Add Key:=col, _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange sortRange
    .Header = xlGuess
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
    End With
    'Add Chart
    With result.ChartObjects.Add _
        (Left:=100, Width:=375, top:=75, Height:=225)
        .Name = "displacement"
        .chart.ChartType = xlXYScatter
        .chart.SetSourceData Source:=result.Range(plotRange.Offset(0, 1), plotRange.Offset(3 * count - 1, 2))
        End With
    Dim chartobj As ChartObject
    Set chartobj = result.ChartObjects("displacement")
    Call formatChart(chartobj.chart)
    Call graphDLine(result.ChartObjects)
    If exportChart Then Call exportChartf(chartobj)
End Sub
'Prepare data for graphing
Sub prepareData(result As Worksheet, count As Integer, force As Boolean, top As Boolean)
    Dim i As Integer, j As Integer
    Dim plotRange As Range, xt As Range, yt As Range, xb As Range, yb As Range
    Set plotRange = result.Range("Force").Cells(1, 1).Offset(0, 1)
    If top Then
        Set xb = result.Range("XT").Cells(1, 1)
        Set yb = result.Range("YT").Cells(1, 1)
    Else
        Set xb = result.Range("XB").Cells(1, 1)
        Set yb = result.Range("YB").Cells(1, 1)
    End If
    
    Set plotRange = result.Range("Force").Cells(1, 1).Offset(0, 1)

    For j = 0 To 2
        For i = 0 To count - 1
            plotRange.Offset(i + j * count).Value = i
            Next i
        Next j

        Set xt = result.Range("scaled_XT").Cells(1, 1)
        Set yt = result.Range("scaled_YT").Cells(1, 1)
        For i = 0 To count - 1
            plotRange.Offset(i, 1).Value = xb.Offset(i)
            plotRange.Offset(i, 2).Value = yb.Offset(i)
            plotRange.Offset(i + count, 1).Value = xt.Offset(i)
            plotRange.Offset(i + count, 2).Value = yt.Offset(i)
            Next i
'    End If

End Sub


'Format chart, line, arrow. Background image has to be named cell
Sub formatChart(chart As chart)
    'Do not show legend
    chart.Legend.Clear
    'Format arrow and line
    With chart.SeriesCollection(1)
        .MarkerStyle = -4142
        .Format.Line.EndArrowheadStyle = msoArrowheadStealth
        .Format.Line.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .Format.Line.ForeColor.TintAndShade = 0
        .Format.Line.ForeColor.Brightness = 0
        .Format.Line.Transparency = 0
        
        .Format.Glow.Color.ObjectThemeColor = msoThemeColorAccent1
        .Format.Glow.Color.TintAndShade = 0
        .Format.Glow.Color.Brightness = 0.400000006
        .Format.Glow.Transparency = 0.4800000191
        .Format.Glow.Radius = 26
    End With
    
    'Insert image background to chart
        'get current directory
        Dim currentDir As String, picDir As String
        currentDir = ThisWorkbook.path
        picDir = currentDir & "\cell.tif"
    With chart.PlotArea.Format.Fill
        .Visible = msoTrue
        .UserPicture picDir
    End With
    'Set the xy- scale of chart to match that of the picture
    'coFactor is the conversion factor from vba to inches
    Dim pic As Object, result As Worksheet, coFactor
    coFactor = 140 / 105
    Set result = ThisWorkbook.Worksheets("result")
    Set pic = result.Pictures.Insert(picDir)
 '   MsgBox pic.Width & "A" & pic.Height
    pic.ShapeRange.ScaleHeight 1, msoTrue
    pic.ShapeRange.ScaleWidth 1, msoTrue
    pic.Visible = msoTrue
    chart.Axes(xlValue).MinimumScale = 0
    chart.Axes(xlValue).MaximumScale = pic.Height * coFactor
    chart.Axes(xlCategory).MinimumScale = 0
    chart.Axes(xlCategory).MaximumScale = pic.Width * coFactor
    For Each ax In chart.Axes
        ax.HasMajorGridlines = False
        ax.HasMinorGridlines = False
        Next
End Sub
'graph the boundary of d-region
Sub graphDLine(chartobjs As ChartObjects)
    Dim region As Worksheet, pRange As Range, i As Integer
    Dim chartobj As ChartObject
    Dim chrt As chart
    Set chrt = chartobjs("displacement").chart
    Set region = ThisWorkbook.Worksheets("Region")
    For i = 1 To 2
        region.Range("dBoundary").Cells(6 * i - 4, 2).Value = chrt.Axes(xlCategory).MaximumScale
        region.Range("dBoundary").Cells(6 * i - 4 + 3, 1).Value = chrt.Axes(xlValue).MaximumScale
        Next i
    Set pRange = region.Range("dBoundary")
    pRange.Select
    With chrt.SeriesCollection.NewSeries
        .Name = "dboundary"
        .XValues = pRange.Columns(1)
        .Values = pRange.Columns(2)
        End With
    
End Sub
'Export chart to image if user say yes
Sub exportChartf(chartobj As ChartObject)
    Name = "result.png"
    On Error Resume Next
    Kill ThisWorkbook.path & "\" & Name
    On Error GoTo 0
    chartobj.Activate
    chartobj.chart.Export Filename:=ThisWorkbook.path & "\" & Name, Filtername:="PNG"
End Sub

