Attribute VB_Name = "Module4"
' Sort an array by exporting it to a newly created worksheet.
' from http://www.cpearson.com/excel/SortingArrays.aspx
' A is the array need to be sorted
' IND is the original index of the array
Sub SortViaWorksheet(a() As Double, ind() As Integer)
Dim ws As Worksheet 'temporatory worksheet
Dim r As Range 'Range in the temp sheet to store the array and index
Dim i As Integer 'counter
For i = LBound(a) To UBound(a)
    Debug.Print ind(i) & " :"; a(i)
    Next i

'Turn off screen updating to speed up the running of macro
Application.ScreenUpdating = False

'Create a new sheet
Set ws = ThisWorkbook.Worksheets.Add

'Put the array on the worksheet
Set r = ws.Range("A1").Resize(UBound(a) - LBound(a) + 1, 2)
For i = 1 To r.Rows.count
    r.Cells(i, 1) = a(i)
    r.Cells(i, 2) = ind(i)
    Next i

'Sort the range
r.Sort key1:=r.Columns(1), order1:=xlAscending, MatchCase:=False


'Load the worksheet value back to the array
For i = 1 To r.Rows.count
    a(i) = r(i, 1)
    ind(i) = r(i, 2)
    Next i
    
'delete the temporary sheet
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True

'check
'For i = LBound(a) To UBound(a)
'    Debug.Print ind(i) & " :"; a(i)
'    Next i
    
End Sub
Sub te()
ThisWorkbook.Worksheets("Sheet3").Range("A1:B20").Columns(1).Select
End Sub
