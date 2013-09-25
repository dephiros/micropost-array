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

'Accept a REG array that store the index post in each region. Each region is one column with 1-A, 2-B, 3-C, 4-D, 5-F
'REG_COUNT array that store the number of posts in each region
'S is the result worksheet
'Sub will modify reg and reg_count appropriately for all regions

Sub region(reg() As Integer, reg_count() As Integer, s As Worksheet)
    ReDim reg_count(1 To 5) As Integer
    Dim centroid(1 To 2) As Double
    Dim xBase() As Variant, yBase() As Variant
    xBase = s.Range("XB").Value
    yBase = s.Range("YB").Value
    Call centroidCal(xBase, yBase, centroid)
    Call dregion(centroid, reg, reg_count, _
    xBase, yBase)
End Sub

'Figure out the D-region. D-region composes of 1/3 numbers of of posts that are closest to the center.
'Index for dregion is 4 in reg array
Sub dregion(centroid() As Double, reg() As Integer, reg_count() As Integer, _
    x() As Variant, y() As Variant)
    Dim distance() As Double, ind() As Integer
    ReDim distance(LBound(x) To UBound(x)) As Double, _
        ind(LBound(x) To UBound(x)) As Integer
    Dim oneThird As Integer
    oneThird = Round((1 / 3) * (UBound(x) - LBound(x)))
    ReDim reg(1 To oneThird, 1 To 5) As Integer
    Dim i As Integer
    For i = LBound(x) To UBound(x)
        distance(i) = ((x(i, 1) - centroid(1)) ^ 2 + (y(i, 1) - centroid(2)) ^ 2) ^ (1 / 2)
        ind(i) = i
        Next i
    
    Call Module4.SortViaWorksheet(distance, ind)
    For i = LBound(ind) To oneThird
        reg(i, 4) = ind(i)
        Next i
    For i = LBound(reg) To UBound(reg)
        Debug.Print ind(i) & ": "; reg(i, 4)
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
