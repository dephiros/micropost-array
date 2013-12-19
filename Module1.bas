Attribute VB_Name = "Module1"
'This is the start function. Every other sub/func after this is organized in
'alphabetical order.
Sub start()
    Options.Show
End Sub

'Calculate average of Major and Minor and return an array
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
    'variable for xy coordinate of top and bottom
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
        If (sh.Name Like "constant" Or sh.Name Like "Constant") Then sh.Delete
    Next
    Set constant = Sheets.Add
    constant.Name = "Constant"
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
Sub clearName(ws As Worksheet)
    Dim nm As Name
    On Error Resume Next
    For Each nm In ThisWorkbook.names
        nm.Delete
    Next nm
    On Error GoTo 0

End Sub

'Check if a cell has value
Function hasValue(cell As Range) As Boolean
    hasValue = Not IsEmpty(cell.Value) And cell.Value <> ""
End Function

'Import the top/bottom worksheet to this current worksbook.
'If the current workbook already has top/bottom then delete them
'before import
Sub importSheet(path As String)
    Application.DisplayAlerts = False
    Set result = ThisWorkbook
    top = "top.xls"
    bottom = "bottom.xls"
    Workbooks.Open Filename:=path & "\" & top, ReadOnly:=True
    Workbooks.Open Filename:=path & "\" & bottom, ReadOnly:=True
    'check if the top and bottom sheet currently exist in this workbook.
    Set temp = result.Sheets.Add
    For Each sh In result.Worksheets
        If (sh.Name Like "top" Or sh.Name Like "bottom") Then sh.Delete
    Next
    'Copy top and bottom sheet
    Workbooks(top).Worksheets(1).Copy After:=result.Worksheets(result.Sheets.count)
    Workbooks(bottom).Worksheets(1).Copy After:=result.Worksheets(result.Sheets.count)
    temp.Delete
    Workbooks(top).Close
    Workbooks(bottom).Close
    Application.DisplayAlerts = True

End Sub

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

'Search and match data according to dx, dy and put to result worksheet.
Public Sub matchData(top As Worksheet, bottom As Worksheet, result As Worksheet, scl As Double)
    Dim tX As Range, tY As Range, bX As Range, bY As Range, average() As Double
    Dim count As Integer, count0 As Integer, count1 As Integer
    Dim i As Integer, j As Integer
    'proposalA, distanceA is to store list of matched bottom/top.
    'bachelor is stack to match the next unmatched bottom.
    Dim proposalA() As Integer, distanceA() As Double, bachelor() As Integer, lid As Integer
    count = 0 'number of post getting matched.
    Set tX = Range("Xtop")
    Set tY = Range("Ytop")
    Set bX = Range("Xbottom")
    Set bY = Range("Ybottom")
    'Remove data with too small area
    Dim aB As Range, aT As Range
    Set aB = Range("Areabottom")
    Set aT = Range("Areatop")
    averageB = WorksheetFunction.average(aB)
    averageT = WorksheetFunction.average(aT)
    For i = 1 To aB.Rows.count
        If aB.Cells(i, 1) < averageB / 2 Then
            aB.Cells(i, 1).EntireRow.Delete
            End If
        Next i
    For i = 1 To aT.Rows.count
        If aT.Cells(i, 1) < averageT / 2 Then
            aT.Cells(i, 1).EntireRow.Delete
            End If
        Next i
    
    'Generate distance from bottom to each top post.
    Dim matchSheet As Worksheet
    Set matchSheet = ThisWorkbook.Worksheets.Add
    'Match bottom to top.
    count = 0 ' the number of post matched.
    count0 = bX.Rows.count 'the number of bottom/proposer
    count1 = tX.Rows.count ' the number of top/proposee
    ReDim proposalA(1 To count1) As Integer, distanceA(1 To count1) As Double
    ReDim bachelor(1 To count0) As Integer
    'initialize proposalA and distanceA, 0 means no proposer
    For i = 1 To count1
        proposalA(i) = 0
        distanceA(i) = 0
        Next i
    lid = 0
    'store each column to process
    Dim col() As Range, bottomMatch() As Integer
    ReDim col(t To count0) As Range, bottomMatch(1 To count0) As Integer
    For i = 1 To count0
        Set temp = matchSheet.Range("A1").Offset(0, 2 * i - 2)
        'generate count from 1 to row number of top
        temp.Value = 1
        temp.DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, _
            Step:=1, Stop:=count1, Trend:=False
        Set temp = temp.Resize(count1, 2)
        tempX = bX.Cells(i)
        tempY = bY.Cells(i)
        For j = 1 To count1
            temp.Cells(j, 2) = ((tempY - tY.Cells(j)) ^ 2 + (tempX - tX.Cells(j)) ^ 2) ^ 0.5
            Next j
        'Sort each post in ascending order
        temp.Sort key1:=temp.Columns(2), order1:=xlAscending, MatchCase:=False
        'make sure the distance between two post is not greater than max
        'of major or minor of a post.
        limit = WorksheetFunction.Max(Range("Majorbottom").Cells(i).Value, _
            Range("Minorbottom").Cells(i).Value)
        If temp.Cells(1, 2).Value >= limit Then
            temp.ClearContents
            Set col(i) = temp.Resize(1, 2)
        Else 'if it passes, initialize proposalA, distanceA and bachelorA
            Set col(i) = temp
            'Initialize proposalA and distanceA with unchecked proposer and bachelor.
            j = temp.Cells(1, 1).Value
            proposalA(temp.Cells(1, 1).Value) = i 'the first proposee of current proposer.
            distanceA(temp.Cells(1, 1).Value) = temp.Cells(1, 2).Value
            bachelor(lid + 1) = i 'the stack will go backward from the last proposer to first.
            lid = lid + 1
        End If
        Next i
        
    'Match bottom with top and write to RESULT sheet.

    While lid > 0
        i = bachelor(lid)
        If hasValue(col(i).Cells(1, 1)) Then 'there is a proposee
            j = proposeMatch(i, col(i).Cells(1, 1).Value, col(i).Cells(1, 2).Value, _
                proposalA, distanceA, count1)
            If j <> 0 Then
                bachelor(lid) = j
            Else
                bachelor(lid) = 0
                lid = lid - 1
            End If
            col(i).Rows(1).ClearContents 'clear the row and shorten column i
            If (col(i).Rows.count > 1) Then
                Set col(i) = col(i).Rows(2).Resize(col(i).Rows.count - 1, 2)
            End If
        Else
            bachelor(lid) = 0
            lid = lid - 1
        End If
    Wend
    'Write result to RESULT, j keeps track of many have been written.
    j = 0
    For i = 1 To count1
        If (proposalA(i) > 0) Then
            j = j + 1
            Call wResult(top, bottom, result, j, i, proposalA(i), scl)
        End If
    Next i
    matchSheet.Delete
        
End Sub
'Each top/bottom post will propose to a bottom/top post. The proposal will be checked
'the proposal array. If there is conflict, the distance will be compared; if the
'current proposal has shorter distance to the proposed post.
'Whoever lose will be return
'The function return index of loser, 0 for no loser
'COUNT is the number of row of proposalA and distanceA
'NOTE: proposalA use proposee as index, col1 is proposer. distance A use proposee as
' index and col1 is distance to proposer
Function proposeMatch(proposer As Integer, proposee As Integer, _
    distance As Double, proposalA() As Integer, distanceA() As Double _
    , count As Integer) As Integer
    loser = 0
    If (proposalA(proposee) <> proposer) Then
        If (distanceA(proposee) > distance) Then
            loser = proposalA(proposee)
            distanceA(proposee) = distance
        Else
            loser = proposer
        End If
    End If
    proposeMatch = loser
End Function
'Write result to result worksheet in approriate units. Coordinate is kept as measured for graphing.
's is the scale
Sub wResult(top As Worksheet, bottom As Worksheet, result As Worksheet, count As Integer, rowT As Integer, rowB As Integer, s As Double)
    'More accurate scale needed.
    Range("AreaT").Cells(count, 1).Value = Range("Areatop").Cells(rowT, 1).Value * s ^ 2 'um^2
    Range("XT").Cells(count, 1).Value = Range("Xtop").Cells(rowT, 1).Value
    Range("YT").Cells(count, 1).Value = Range("Ytop").Cells(rowT, 1).Value
    Range("MajorT").Cells(count, 1).Value = Range("Majortop").Cells(rowT, 1).Value * s 'um
    Range("MinorT").Cells(count, 1).Value = Range("Minortop").Cells(rowT, 1).Value * s 'um
    Range("AreaB").Cells(count, 1).Value = Range("Areabottom").Cells(rowB, 1).Value * s ^ 2 'um
    Range("XB").Cells(count, 1).Value = Range("Xbottom").Cells(rowB, 1).Value
    Range("YB").Cells(count, 1).Value = Range("Ybottom").Cells(rowB, 1).Value
    Range("MajorB").Cells(count, 1).Value = (Range("Majorbottom").Cells(rowB, 1).Value) * s 'um
    Range("MinorB").Cells(count, 1).Value = (Range("Minorbottom").Cells(rowB, 1).Value) * s 'um
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

Sub nameCol(ws As Worksheet)
    Dim wb
    Dim nameRange As Range, i As Range
    Set data = ThisWorkbook
    Set nameRange = ws.Range("A1", ws.Cells(1, 1).End(xlToRight))
    For Each i In nameRange
        If hasValue(i) Then ws.Range(i.Offset(1, 0), i.End(xlDown)).Name = i.Value & ws.Name
        Next i
End Sub


'Display a userform for user to choose option as well as scale
Sub optionSelect(exportChart As Boolean, scl As Double, topAsBottom As Boolean)
    Options.Show
    exportChart = Options.exportChart.Value
    topAsBottom = Options.topAsBottom.Value
    scl = Options.scale_um / Options.scale_pixel
End Sub

'Print to specified column for debugging.
Sub Out(item As Variant, Optional column As String = "A")
Dim ws As Worksheet
Set ws = ActiveSheet
count = Application.WorksheetFunction.count(ws.Range(column & ":" & column))
ws.Range(column & 1).Offset(count + 1, 0).Value = item
End Sub

Public Sub ProcessData(exportChart As Boolean, scl As Double, topAsBottom As Boolean)
    Application.ScreenUpdating = False
    Call importSheet(ThisWorkbook.path)
    Dim top As Worksheet, bottom As Worksheet, result As Worksheet
    Set top = Worksheets("top")
    Set bottom = Worksheets("bottom")
    Call sortCol(top, "Y")
    Call sortCol(bottom, "Y")
    'Name All column in worksheet
    Call clearName(top)
    Call nameCol(top)
    Call nameCol(bottom)
    'Create new sheet
    Application.DisplayAlerts = False
    For Each sh In Worksheets
        If sh.Name Like "result" Then sh.Delete
    Next
    Set result = Sheets.Add
    result.Name = "result"
    Application.DisplayAlerts = True
    Dim rowNum As Integer
    If top.Cells(1, 1).End(xlDown).Row > bottom.Cells(1, 1).End(xlDown).Row Then
        rowNum = bottom.Cells(1, 1).End(xlDown).Row - 1
        Else
        rowNum = top.Cells(1, 1).End(xlDown).Row - 1
        End If
    'create named column for result
    Call makeNamedResult(result, rowNum)
    'match top and bottom data
    Call matchData(top, bottom, result, scl)
    'calculate displacement and force data
    Call Calculation(result)
    Call scaleG(3)
    'user input
    Call Module3.testRegion
    Call Module2.Graph(exportChart, result, scl, topAsBottom)
    Application.ScreenUpdating = True
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


