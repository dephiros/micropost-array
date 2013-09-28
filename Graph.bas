Attribute VB_Name = "Module2"
'Graph result
Sub Graph(result As Worksheet)
    'modified data for graphing
    Dim count As Integer, i As Integer, j As Integer
    Dim plotRange As Range, xt As Range, yt As Range, xb As Range, yb As Range
    Set xt = result.Range("scaled_XT").Cells(1, 1)
    Set yt = result.Range("scaled_YT").Cells(1, 1)
    Set xb = result.Range("XB").Cells(1, 1)
    Set yb = result.Range("YB").Cells(1, 1)
    Set plotRange = result.Range("Force").Cells(1, 1).Offset(0, 1)
    count = Application.WorksheetFunction.count(result.Range("XT"))
    For j = 0 To 2
        For i = 0 To count - 1
            plotRange.Offset(i + j * count).Value = i
            Next i
        Next j
    For i = 0 To count - 1
        plotRange.Offset(i, 1).Value = xb.Offset(i)
        plotRange.Offset(i, 2).Value = yb.Offset(i)
        plotRange.Offset(i + count, 1).Value = xt.Offset(i)
        plotRange.Offset(i + count, 2).Value = yt.Offset(i)
        Next i
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
        .name = "displacement"
        .chart.ChartType = xlXYScatter
        .chart.SetSourceData Source:=result.Range(plotRange.Offset(0, 1), plotRange.Offset(3 * count - 1, 2))
        End With
    Dim chart As chart
    Set chart = result.ChartObjects("displacement").chart
    Call formatChart(chart)
End Sub
'Forat chart, line, arrow. Background image has to be named cell
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
        currentDir = ThisWorkbook.Path
        picDir = currentDir & "\cell.tif"
    With chart.PlotArea.Format.Fill
        .Visible = msoTrue
        .UserPicture picDir
    End With
    'Set the xy- scale of chart to match that of the picture
    'coFactor is the conversion factor from vba to inches
    Dim pic As Object, result As Worksheet, coFactor
    coFactor = 0.78 / 56.25
    Set result = ThisWorkbook.Worksheets("result")
    Set pic = result.Pictures.Insert(picDir)
    MsgBox pic.Width & "A" & pic.Height
    pic.ShapeRange.ScaleHeight 1, msoTrue
    pic.ShapeRange.ScaleWidth 1, msoTrue
    pic.Visible = msoTrue
    chart.Axes(xlValue).MinimumScale = 0
    chart.Axes(xlValue).MaximumScale = pic.Height * 100 * coFactor
    chart.Axes(xlCategory).MinimumScale = 0
    chart.Axes(xlCategory).MaximumScale = pic.Width * 100 * coFactor
End Sub
