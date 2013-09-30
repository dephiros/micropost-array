Attribute VB_Name = "Module3"
'Calculate centroid of cell from the base coordinate of post. Modified CENTROID array to be
'double array with two elelemnt storing x and y coordinate of centroid
Sub centroidCal(xBase() As Variant, yBase() As Variant, centroid() As Double)
    Dim i As Integer, length As Integer
    centroid(1) = 0
    centroid(2) = 0
    length = UBound(xBase) - LBound(xBase) + 1
    For i = 1 To length
        centroid(1) = centroid(1) + xBase(i, 1)
        centroid(2) = centroid(2) + yBase(i, 1)
        Next i
    centroid(1) = centroid(1) / length
    centroid(2) = centroid(2) / length
End Sub

'Accept a REG array that store the index post in each region. Each region is one column with 1-A, 2-B, 3-C, 4-D, 5-E, 6-F
'REG_COUNT array that store the number of posts in each region
'S is the result worksheet
'Sub will modify reg and reg_count appropriately for all regions

Sub region(reg() As Integer, reg_count() As Integer, s As Worksheet)
    Application.ScreenUpdating = False
    'allocated space as 1/2 the total number of post for each region
    Dim x() As Variant, y() As Variant
    x = s.Range("XB").Value
    y = s.Range("YB").Value
    ReDim reg(1 To Round((1 / 3) * (UBound(x) - LBound(x))), _
        1 To 6) As Integer
    ReDim reg_count(1 To 6) As Integer
    Dim centroid(1 To 2) As Double, ind() As Integer
    Dim dBoundary() As Double
    Dim i As Integer
    For i = 1 To 5
        reg_count(i) = 0
        Next i
    Call centroidCal(x, y, centroid)
    dBoundary = regionD(centroid, ind, reg, reg_count, _
    x, y)
    Call regionA(dBoundary, ind, reg, reg_count, x, y)
    Call regionE(dBoundary, ind, reg, reg_count, x, y)
    Call regionB(dBoundary, ind, reg, reg_count, x, y)
    Call regionF(dBoundary, ind, reg, reg_count, x, y)
    Call regionC(dBoundary, ind, reg, reg_count, x, y)
    Call writeRegion(dBoundary, UBound(x) - LBound(x) + 1, reg, reg_count)
    Application.ScreenUpdating = True
End Sub
'Figure out the region A(1). Region A composes of all posts that are to the top left and bottom left
' of region D
'dBoundary is the boundary of region d(refer to regionD)
'ind() is to keep track of unassigned index sorted by distance to centroid
'reg(), reg_count() - refer to region()
'x, y contains x, y coordinate of post
Sub regionA(dBoundary() As Double, ind() As Integer, reg() As Integer, reg_count() As Integer, _
    x() As Variant, y() As Variant)
    Dim i As Integer
    For i = LBound(ind) To UBound(ind)
        'check if ind(i) as already been assigned region
        If ind(i) = -1 Then
            GoTo continue
            End If
        'if ind(i) is top-left of regionD
        If (x(ind(i), 1) <= dBoundary(1)) And (y(ind(i), 1) >= dBoundary(4)) Then
            reg_count(1) = reg_count(1) + 1
            reg(reg_count(1), 1) = ind(i)
            ind(i) = -1
            GoTo continue
            End If
        If (x(ind(i), 1) <= dBoundary(1)) And (y(ind(i), 1) <= dBoundary(2)) Then
            reg_count(1) = reg_count(1) + 1
            reg(reg_count(1), 1) = ind(i)
            ind(i) = -1
            GoTo continue
            End If
continue:
        Next i
End Sub

'regionB(2) is posts that are to the left of region D but are not in region A or D
Sub regionB(dBoundary() As Double, ind() As Integer, reg() As Integer, reg_count() As Integer, _
    x() As Variant, y() As Variant)
    Dim i As Integer
    For i = LBound(ind) To UBound(ind)
            'check if ind(i) as already been assigned region
        If ind(i) = -1 Then
            GoTo continue
            End If
        If (x(ind(i), 1) <= dBoundary(1)) Then
            reg_count(2) = reg_count(2) + 1
            reg(reg_count(2), 2) = ind(i)
            ind(i) = -1
            GoTo continue
            End If
continue:
            Next i
        
End Sub
'regionC(3) is the rest of the posts
Sub regionC(dBoundary() As Double, ind() As Integer, reg() As Integer, reg_count() As Integer, _
    x() As Variant, y() As Variant)
        Dim i As Integer
    For i = LBound(ind) To UBound(ind)
            'check if ind(i) as already been assigned region
        If ind(i) = -1 Then
            GoTo continue
            End If
        reg_count(3) = reg_count(3) + 1
        reg(reg_count(3), 3) = ind(i)
        ind(i) = -1
continue:
            Next i
        
End Sub





'Figure out the region D. region D composes of 1/3 numbers of of posts that are closest to the center.
'Index for dregion is 4 in reg array
'regionD returns the boundary of the regionD (smallestx, smallesty, biggestx, biggest y)
Function regionD(centroid() As Double, ind() As Integer, reg() As Integer, reg_count() As Integer, _
    x() As Variant, y() As Variant)
    Dim distance() As Double
    ReDim distance(LBound(x) To UBound(x)) As Double, _
        ind(LBound(x) To UBound(x)) As Integer
    Dim oneThird As Integer
    oneThird = Round((1 / 3) * (UBound(x) - LBound(x)))
    Dim i As Integer
    For i = LBound(x) To UBound(x)
        distance(i) = ((x(i, 1) - centroid(1)) ^ 2 + (y(i, 1) - centroid(2)) ^ 2) ^ (1 / 2)
        ind(i) = i
        Next i
    
    Call Module4.SortViaWorksheet(distance, ind)
    For i = LBound(ind) To oneThird
        reg(i, 4) = ind(i)
        ind(i) = -1 'mark ind(i) as used already assigned to a region
        Next i
    reg_count(4) = oneThird
    'Calculate the boundary of d_region
    Dim boundary(1 To 4) As Double
    Dim dx_sort() As Double, dy_sort() As Double, indtemp() As Integer
    ReDim dx_sort(1 To oneThird) As Double, dy_sort(1 To oneThird) As Double, _
        dind(1 To oneThird) As Integer
    For i = 1 To oneThird
        dx_sort(i) = x(reg(i, 4), 1)
        dy_sort(i) = y(reg(i, 4), 1)
        Next i
    Call Module4.SortViaWorksheet(dx_sort, dind)
    Call Module4.SortViaWorksheet(dy_sort, dind)
    boundary(1) = dx_sort(1)
    boundary(2) = dy_sort(1)
    boundary(3) = dx_sort(UBound(dx_sort))
    boundary(4) = dy_sort(UBound(dy_sort))
    regionD = boundary
End Function

'regionE(5) is similar to regions A but contains post that are top right and bottom right to region D
'refer to regionA for variable doc
Sub regionE(dBoundary() As Double, ind() As Integer, reg() As Integer, reg_count() As Integer, _
    x() As Variant, y() As Variant)
    Dim i As Integer
    For i = LBound(ind) To UBound(ind)
        'check if ind(i) as already been assigned region
        If ind(i) = -1 Then
            GoTo continue
            End If
        'if ind(i) is top-left of regionD
        If (x(ind(i), 1) >= dBoundary(3)) And (y(ind(i), 1) >= dBoundary(4)) Then
            reg_count(5) = reg_count(5) + 1
            reg(reg_count(5), 5) = ind(i)
            ind(i) = -1
            GoTo continue
            End If
        If (x(ind(i), 1) >= dBoundary(3)) And (y(ind(i), 1) <= dBoundary(2)) Then
            reg_count(5) = reg_count(5) + 1
            reg(reg_count(5), 5) = ind(i)
            ind(i) = -1
            GoTo continue
            End If
continue:
        Next i
End Sub

'regionF(6) is posts that are to the right of region D but are not in region E or D
Sub regionF(dBoundary() As Double, ind() As Integer, reg() As Integer, reg_count() As Integer, _
    x() As Variant, y() As Variant)
    Dim i As Integer
    For i = LBound(ind) To UBound(ind)
            'check if ind(i) as already been assigned region
        If ind(i) = -1 Then
            GoTo continue
            End If
        If (x(ind(i), 1) >= dBoundary(3)) Then
            reg_count(6) = reg_count(6) + 1
            reg(reg_count(6), 6) = ind(i)
            ind(i) = -1
            GoTo continue
            End If
continue:
            Next i
        
End Sub
'write region data to a spreadsheet for graphing purpose to a worksheet called region
Sub writeRegion(dBoundary() As Double, post_num As Integer, reg() As Integer, reg_count() As Integer)
    Dim region As Worksheet, boundaryX As Range, boundaryY As Range
    Dim iRange As Range, iHeader As Range
    Dim i As Integer, j As Integer, regionNum As Integer
    regionNum = UBound(reg_count) - LBound(reg_count) + 1
    Dim names() As Variant
    Application.DisplayAlerts = False
    For Each sh In Worksheets
        If sh.Name Like "Region" Then sh.Delete
        Next
    Set region = ThisWorkbook.Sheets.Add
    region.Name = "Region"
    Application.DisplayAlerts = True
    'plus 2 at the end is for boundaryx and boundaryy
    ReDim names(1 To regionNum + 2) As Variant
    'make the array of name
    For i = LBound(names) To regionNum
        names(i) = "Region" & Chr(i + 64)
        Next i
    names(regionNum + 1) = "dBoundaryX"
    names(regionNum + 2) = "dBoundaryY"
    Set iRange = region.Range("A2")
    Set iHeader = region.Range("A1")
    'create the range with name according to the array
    For i = LBound(names) To regionNum
        Set iHeader = iHeader.Offset(0, 1)
        iHeader.Value = names(i)
        Set iRange = iRange.Offset(0, 1)
        Set iRange = iRange.Resize(reg_count(i), 1)
        iRange.Name = names(i)
        For j = 1 To iRange.Rows.count
            iRange.Cells(j, 1).Value = reg(j, i)
            Next j
        Next i
    j = 1
    For i = regionNum + 1 To UBound(names)
        Set iHeader = iHeader.Offset(0, 1)
        iHeader.Value = names(i)
        Set iRange = iRange.Offset(0, 1)
        Set iRange = iRange.Resize(2, 1)
        iRange.Name = names(i)
        iRange.Cells(1, 1).Value = dBoundary(j)
        iRange.Cells(2, 1).Value = dBoundary(j + 2)
        j = j + 1
        Next i
End Sub


'Test region calculation
Sub testRegion()
    Dim s As Worksheet
    Set s = Worksheets("result")
    Dim centroid(1 To 2) As Double
    Dim xBase() As Variant, yBase() As Variant
    Dim reg() As Integer, reg_count() As Integer
    xBase = s.Range("XB").Value
    yBase = s.Range("YB").Value
    Call centroidCal(xBase, yBase, centroid)
    Call region(reg, reg_count, s)
End Sub
