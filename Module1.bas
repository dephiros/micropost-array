Attribute VB_Name = "Module1"
Public Sub ProcessData(exportChart As Boolean, scl As Double, topAsBase As Boolean)
    Application.ScreenUpdating = False
    Dim top As Worksheet, base As Worksheet, result As Worksheet
    Set top = Worksheets("top")
    Set base = Worksheets("base")
    Call sortCol(top, "Y")
    Call sortCol(base, "Y")
    'Name All column in worksheet
    Call clearName(top)
    Call nameCol(top)
    Call nameCol(base)
    'Create new sheet
    Application.DisplayAlerts = False
    For Each sh In Worksheets
        If sh.Name Like "result" Then sh.Delete
    Next
    Set result = Sheets.Add
    result.Name = "result"
    Application.DisplayAlerts = True
    Dim rowNum As Integer
    If top.Cells(1, 1).End(xlDown).Row > base.Cells(1, 1).End(xlDown).Row Then
        rowNum = base.Cells(1, 1).End(xlDown).Row - 1
        Else
        rowNum = top.Cells(1, 1).End(xlDown).Row - 1
        End If
    'create named column for result
    Call makeNamedResult(result, rowNum)
    'match top and base data
    Call matchData(top, base, result, scl)
    'calculate displacement and force data
    Call Calculation(result)
    Call scaleG(3)
    'user input
    Call Module3.testRegion
    Call Module2.Graph(exportChart, result, scl, topAsBase)
    Application.ScreenUpdating = True
End Sub
Sub clearName(ws As Worksheet)
    Dim nm As Name
    On Error Resume Next
    For Each nm In ThisWorkbook.names
        nm.Delete
    Next nm
    On Error GoTo 0

End Sub
Sub nameCol(ws As Worksheet)
    Dim wb
    Dim nameRange As Range, i As Range
    Set data = ThisWorkbook
    Set nameRange = ws.Range("A1", ws.Cells(1, 1).End(xlToRight))
    For Each i In nameRange
        If hasValue(i) Then ws.Range(i.Offset(1, 0), i.End(xlDown)).Name = i.Value & ws.Name
        Next i
End Sub
Sub makeNamedResult(result As Worksheet, rowNum As Integer)
    result.UsedRange.Clear
    Dim data As Workbook, names() As Variant
    Set data = ThisWorkbook
    names() = Array("AreaT", "XT", "YT", "Scaled_XT", "Scaled_YT", "MajorT", "MinorT", "AreaB", "XB", "YB", "MajorB", _
        "MinorB", "Displacement", "Theta", "kn", "kd", "k", "Force")
    'starting header, starting range, counter.
    Dim iRange As Range, iHeader As Range, counter As Integer
    counter = 0
    Set iHeader = result.Range("B1")
    Set iRange = result.Range("B2", "B" & rowNum + 1)
    For Each Name In names
        iHeader.Offset(0, counter).Value = Name
        iRange.Offset(0, counter).Name = Name
        counter = counter + 1
    Next Name

End Sub
'Display a userform for user to choose option as well as scale
Sub optionSelect(exportChart As Boolean, scl As Double, topAsBase As Boolean)
    Options.Show
    exportChart = Options.exportChart.Value
    topAsBase = Options.topAsBase.Value
    scl = Options.scale_um / Options.scale_pixel
End Sub


'sort NAME column in workseet in descending order
    Public Sub sortCol(ws As Worksheet, colName As String)
    Dim col As Range, sortRange As Range, lastRow As Range
    Set col = ws.Cells.Find(colName, , xlValues, xlWhole).EntireColumn.SpecialCells(xlCellTypeConstants).Offset(1).SpecialCells(xlCellTypeConstants)
    'select sortRange
    Set sortRange = ws.Range("A1", ws.Cells(1, 1).End(xlDown).End(xlToRight))

    'Sort data
    With ws.Sort
    .SortFields.Clear
    .SortFields.Add Key:=col, _
    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    .SetRange sortRange
    .Header = xlGuess
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
    End With
 
End Sub

'Search and match data according to dx, dy and put to result worksheet.
Public Sub matchData(top As Worksheet, base As Worksheet, result As Worksheet, scl As Double)
    Dim tX As Range, tY As Range, bX As Range, bY As Range, average() As Double, count As Integer
    'Set tX = top.Cells.Find("X", , xlValues, xlWhole).EntireColumn.SpecialCells(xlCellTypeConstants).Offset(1).SpecialCells(xlCellTypeConstants).Cells
    'Set tY = top.Cells.Find("Y", , xlValues, xlWhole).EntireColumn.SpecialCells(xlCellTypeConstants).Offset(1).SpecialCells(xlCellTypeConstants).Cells
    'Set bX = base.Cells.Find("X", , xlValues, xlWhole).EntireColumn.SpecialCells(xlCellTypeConstants).Offset(1).SpecialCells(xlCellTypeConstants).Cells
    'Set bY = base.Cells.Find("Y", , xlValues, xlWhole).EntireColumn.SpecialCells(xlCellTypeConstants).Offset(1).SpecialCells(xlCellTypeConstants).Cells
    count = 0
    Set tX = Range("Xtop")
    Set tY = Range("Ytop")
    Set bX = Range("Xbase")
    Set bY = Range("Ybase")
    average = averageR(top, "Major", "Minor")
    Dim i As Integer, j As Integer, k As Integer, min As Integer, usedJ() As Integer
    Dim last As Double, dx As Double, dy As Double, current As Double
    ReDim usedJ(1 To tX.Rows.count)
    current = 0
    dx = 0
    dy = 0
    For i = 1 To tX.Rows.count
        j = 1
        min = 0
        last = 0
        While j <= bX.Rows.count
            
            If Not isUSedJ(j, usedJ) Then
                dx = tX.Cells(i, 1).Value - bX.Cells(j, 1).Value
                dy = tY.Cells(i, 1).Value - bY.Cells(j, 1).Value
                current = Sqr((dx ^ 2) + (dy ^ 2))
                    'Out i & j
                If (Not last <> 0 Or current < last) And current < Abs(average(i)) Then
                    last = current
                    min = j
                                        'Out j & ":" & min
                End If
            End If
            j = j + 1
        Wend
        If min <> 0 Then
            count = count + 1
            Call wResult(top, base, result, count, i, min, scl)
            Dim l As Integer
            For l = 1 To UBound(usedJ)
                If usedJ(l) = 0 Then
                    usedJ(l) = min
                    Exit For
                    End If
                Next l
        End If
    Next i
End Sub
'Check if a cell has value
Function hasValue(cell As Range)
    hasValue = Not IsEmpty(cell.Value) And cell.Value <> ""
End Function

'Check if j is in used j
Function isUSedJ(j As Integer, usedJ() As Integer) As Boolean
    Dim used
    used = False
    For Each i In usedJ
        If i = j Then
            used = True
            Exit For
            End If
        Next i
isUSedJ = used
End Function
'Write result to result worksheet in approriate units. Coordinate is kept as measured for graphing.
's is the scale
Sub wResult(top As Worksheet, base As Worksheet, result As Worksheet, count As Integer, rowT As Integer, rowB As Integer, s As Double)
    'More accurate scale needed.
    Range("AreaT").Cells(count, 1).Value = Range("Areatop").Cells(rowT, 1).Value * s ^ 2 'um^2
    Range("XT").Cells(count, 1).Value = Range("Xtop").Cells(rowT, 1).Value
    Range("YT").Cells(count, 1).Value = Range("Ytop").Cells(rowT, 1).Value
    Range("MajorT").Cells(count, 1).Value = Range("Majortop").Cells(rowT, 1).Value * s 'um
    Range("MinorT").Cells(count, 1).Value = Range("Minortop").Cells(rowT, 1).Value * s 'um
    Range("AreaB").Cells(count, 1).Value = Range("Areabase").Cells(rowB, 1).Value * s ^ 2 'um
    Range("XB").Cells(count, 1).Value = Range("Xbase").Cells(rowB, 1).Value
    Range("YB").Cells(count, 1).Value = Range("Ybase").Cells(rowB, 1).Value
    Range("MajorB").Cells(count, 1).Value = (Range("Majorbase").Cells(rowB, 1).Value) * s 'um
    Range("MinorB").Cells(count, 1).Value = (Range("Minorbase").Cells(rowB, 1).Value) * s 'um
    End Sub

'Calculate average of Major and Minor and put in column r.
Public Function averageR(ws As Worksheet, majorN As String, minorN As String) As Double()
    Dim major As Range, minor As Range, average() As Double
    Set major = ws.Cells.Find(majorN, , xlValues, xlWhole).EntireColumn.SpecialCells(xlCellTypeConstants).Offset(1).SpecialCells(xlCellTypeConstants).Cells
    Set minor = ws.Cells.Find(minorN, , xlValues, xlWhole).EntireColumn.SpecialCells(xlCellTypeConstants).Offset(1).SpecialCells(xlCellTypeConstants).Cells
    Dim i, Range, c As Integer, cell As Range
    ReDim average(0 To major.Rows.count)
    c = 0
    Set cell = minor.Cells(1, 1)
    For Each i In major
            average(c) = (CDbl(i) + CDbl(cell.Value)) / 2
            c = c + 1
            Set cell = cell.Offset(1, 0)
    Next i
    averageR = average
End Function
'Calculate displacement and force and put in RESULT. Scaling is done in wResult.
Public Sub Calculation(result As Worksheet)
    Dim constant As Worksheet
    'variable for xy coordinate of top and base
    Dim xt As Range, yt As Range, xb As Range, yb As Range, rb() As Variant, ra() As Variant
    Set xt = Range("XT")
    Set yt = Range("YT")
    Set xb = Range("XB")
    Set yb = Range("YB")
    rb = Range("MajorB")
    ra = Range("MajorT")
    Range("Displacement").Formula = "=SQRT((XT - XB)^2 + (YT - YB)^2)"
    'Calculate Force
    'set variable
    Application.DisplayAlerts = False
    For Each sh In Worksheets
        If sh.Name Like "Constant" Then sh.Delete
    Next
    Set constant = Sheets.Add
    constant.Name = "constant"
    Application.DisplayAlerts = True
    Dim E As Double, G As Double, kappa As Double, pi As Double
    pi = Application.WorksheetFunction.pi
    E = 750 'kPa
    G = 250 'kPa
    kappa = 27 / 28
    H = 7 'um
    'k=((3*pi*E*G.*a.*b)*((a^2)*((cos(theta))^2)+(b^2)*(sin(theta)^2)))/...
    '((4*kappa*G*(H^2))+3*E*H*((cos(theta))^2)+(b^2)*(sin(theta)^2));
    'Range("Theta").Formula = "=IF(YT - YB <> 0, ATAN((YT - YB) / (XT - XB)), 0)"
    Range("Theta").Formula = "=ATAN((YT - YB) / (XT - XB))"
    Range("kn").Formula = "=((3*PI()*" & E & "*" & G & ") * MajorB * MinorB" _
        & "* ((MajorB^2)*(COS(Theta)^2) + (MinorB^2)*(SIN(Theta)^2)))"
        
    Range("kd").Formula = "=(((4 *" & kappa & "*" & G & "* (" & H & "^3)) + 3 *" & E & "*" & H & "*" _
        & "((MajorB^2)*(COS(Theta)^2) + (MinorB^2)*(SIN(Theta)^2))))"
    
    Range("k").Formula = "=kn/kd"
    Range("Force").Formula = "=k*Displacement"
End Sub

'Change the top coordinate to scale up the displacement vector by N
Sub scaleG(n As Integer)
'syt, sxt is the scaled yt and xt coordinate
Dim sxt As Range, syt As Range
Set syt = ThisWorkbook.Worksheets("result").Range("Scaled_YT")
Set sxt = ThisWorkbook.Worksheets("result").Range("Scaled_XT")
sxt.Formula = "=XB + (XT - XB) *" & n
syt.Formula = "=YB + (YT - YB) *" & n
End Sub

'Print to specified column for debugging.
Sub Out(item As Variant, Optional column As String = "A")
Dim ws As Worksheet
Set ws = ActiveSheet
count = Application.WorksheetFunction.count(ws.Range(column & ":" & column))
ws.Range(column & 1).Offset(count + 1, 0).Value = item
End Sub

'Find min of a column
Function minCol(col As Range)
    Dim i As Integer, count As Integer, min As Double
    count = Application.WorksheetFunction.count(col)
    On Error Resume Next
    min = col.Cells(1, 1).Value
    For i = 1 To count
        If col.Cells(i, 1) < min Then min = col.Cells(i, 1)
    Next
    minCol = min
End Function

Sub start()
    Options.Show
End Sub

