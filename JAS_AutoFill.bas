Sub AutoFill()
' AutoFill Macro
' Created:  February 8, 2017 by Jacob Schroeder
' Revised:  April 23, 2026
' Shortcut: Ctrl+Shift+D
'
' Fills the active cell's contents and formatting downward, using an adjacent
' column to determine fill length. Prefers the left column; falls back to right.
' Column A always uses the right column as reference.
' Treats error values (#N/A, #REF!, etc.) as non-empty.

    Dim startCell As Range
    Set startCell = ActiveCell
    
    ' Identify reference cell in adjacent column
    Dim refCell As Range
    If startCell.Column = 1 Then
        Set refCell = startCell.Offset(0, 1)                        ' Column A: right only
    ElseIf CellHasContent(startCell.Offset(1, -1)) Then
        Set refCell = startCell.Offset(0, -1)                       ' Left column preferred
    ElseIf CellHasContent(startCell.Offset(1, 1)) Then
        Set refCell = startCell.Offset(0, 1)                        ' Right column fallback
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
    
    startCell.Resize(fillRows + 1, 1).FillDown

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
