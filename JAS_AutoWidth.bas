' Created: 4/3/2018 by Jacob Schroeder
' Revised: 4/23/2026 by Jacob Schroeder

Sub AutoWidth()
    With Selection
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
    End With
End Sub

Sub WidthIncrease()
    AdjustColumnWidth 1
End Sub

Sub WidthDecrease()
    AdjustColumnWidth -1
End Sub

Sub HeightIncrease()
    AdjustRowHeight 5
End Sub

Sub HeightDecrease()
    AdjustRowHeight -5
End Sub

Private Sub AdjustColumnWidth(delta As Double)
    Dim w As Double
    w = ActiveCell.ColumnWidth
    Selection.ColumnWidth = Application.Round(Application.Max(w + delta, 1), 0)
End Sub

Private Sub AdjustRowHeight(delta As Double)
    Dim h As Double
    h = ActiveCell.RowHeight
    Selection.RowHeight = Application.Round(Application.Max(h + delta, 5), 0)
End Sub
