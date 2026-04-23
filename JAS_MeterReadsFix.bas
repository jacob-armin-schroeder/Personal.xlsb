Attribute VB_Name = "JAS_MeterReadsFix"
Sub MeterReadFix()

''''''''''''''''''''''''''''''''''''''''
' Created by Jacob Schroeder, May 2016 '
''''''''''''''''''''''''''''''''''''''''
'Last update: 11/18/2017 - Changed macro to skip first line in order
'to better accomodate new Benton PUD file format.

'To use this macro, select only two columns, including ONE row at the
'top that has the header. The left column must be the timestamp. The
'right column must be the CUMULATIVE meter reading. The macro will
'fill in any missing rows, using the interval between the first two
'readings in the file (in rows 2 and 3). It may be necessary to man-
'ually change the first or second reading in if they are not already
'in the desired interval.
'
'ALSO NOTE: all daylight savings times must be converted to standard
'time manually BEFORE running the macro. This macro is clunky, but
'it can be a time-saver if used cautiously.

'This section added on June 7, 2016 to facilitate fixing Benton PUD data

    Cells.Replace What:="-", Replacement:="/", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select

'End of Benton PUD section

Dim Data, SortRange, SortKey As Range, n As Long, delta As Double, Col_1, Col_2 As Variant

Set Data = Selection

n = Application.CountA(Data) / 2
ReDim Col_1(1 To n, 1 To 1)
ReDim Col_2(1 To n, 1 To 1)

'Added timestamp sorting on Dec. 12, 2016

    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range(ActiveCell, ActiveCell.Offset(n, 1))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Creating the new data

Dim RowsToAdd As Variant, i, p As Long

For i = 1 To n
    Col_1(i, 1) = Data(i, 1)
    Col_2(i, 1) = Data(i, 2)
Next

delta = Round(((Col_1(3, 1) - Col_1(2, 1))) * 1440, 0) / 1440

p = Round((Col_1(n, 1) - Col_1(2, 1)) / delta, 0) + 2

ReDim RowsToAdd(1 To n - 2, 1 To 1)
For i = 1 To n - 2
    RowsToAdd(i, 1) = Round((Data(i + 2, 1) - Data(i + 1, 1)) / delta, 0)

Next

Dim Output As Variant
ReDim Output(1 To p, 1 To 2)

Output(1, 1) = Col_1(1, 1)
For i = 2 To p
    Output(i, 1) = Col_1(2, 1) + (i - 2) * delta
Next

Dim j, k As Double
k = 1
Output(1, 2) = Col_2(1, 1)
For i = 1 To n - 2
    For j = 1 To RowsToAdd(i, 1)
        k = k + 1
        Output(k, 2) = Col_2(i + 1, 1) + (j - 1) * (Col_2(i + 2, 1) - Col_2(i + 1, 1)) / RowsToAdd(i, 1)
    Next
Next
Output(p, 2) = Col_2(n, 1)

Sheets.Add After:=Sheets(Sheets.Count)
Sheets(Sheets.Count).Name = "MeterReadsFixed"
Sheets("MeterReadsFixed").Range("A1").Select
Sheets("MeterReadsFixed").Range(ActiveCell, ActiveCell.Offset(p - 1, 1)) = Output
Range("A1").Select
    ActiveCell.EntireColumn.AutoFit
Range("B1").Select
    ActiveCell.EntireColumn.AutoFit
End Sub
