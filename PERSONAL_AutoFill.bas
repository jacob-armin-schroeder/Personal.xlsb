Attribute VB_Name = "PERSONAL_AutoFill"
Option Explicit

Sub AutoFill()
' Recommended Shortcut: Ctrl+D
'
' Fills the active cell's contents and formatting downward, using an adjacent
' column to determine fill length. Prefers the left column; falls back to right.
' Column A always uses the right column as reference.
' Treats error values (#N/A, #REF!, etc.) as non-empty.

    Dim startCell As Range
    Set startCell = Selection.Cells(1, 1)

    Dim selCols As Long
    selCols = Selection.Columns.Count 
    
    Dim leftCol As Long
    Dim rightCol As Long
    leftCol = startCell.Column
    rightCol = startCell.Column + selCols - 1


    ' Identify reference cell in adjacent column
    Dim refCell As Range
    If leftCol = 1 Then
        Set refCell = Cells(startCell.Row, rightCol + 1)            ' Column A: right only
    ElseIf CellHasContent(Cells(startCell.Row + 1, leftCol - 1)) Then
        Set refCell = Cells(startCell.Row, leftCol - 1)             ' Left column preferred
    ElseIf CellHasContent(Cells(startCell.Row + 1, rightCol + 1)) Then
        Set refCell = Cells(startCell.Row, rightCol + 1)            ' Right column fallback
    Else
        Exit Sub                                                     ' No adjacent data found
    End If
    
    ' Determine number of rows to fill based on reference column depth
    Dim fillRows As Long
    If Not CellHasContent(refCell.Offset(1, 0)) Then
        Exit Sub                                                     ' Nothing below reference cell
    ElseIf Not CellHasContent(refCell.Offset(2, 0)) Then
        fillRows = 1                                                 ' Only one row below
    Else
        fillRows = refCell.End(xlDown).Row - startCell.Row          ' Fill to bottom of ref column
    End If
    
    startCell.Resize(fillRows + 1, selCols).FillDown

End Sub


Private Function CellHasContent(c As Range) As Boolean
' Returns True if the cell contains any value, including error values.
' A direct comparison like c.Value <> "" throws a type mismatch on error values.
    If IsError(c.Value) Then
        CellHasContent = True
    Else
        CellHasContent = (c.Value <> "")
    End If
End Function
