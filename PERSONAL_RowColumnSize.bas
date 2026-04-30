Attribute VB_Name = "PERSONAL_RowColumnSize"
Option Explicit

Sub Autofit()
' Recommended Shortcut: Ctrl+Shift+W
    With Selection
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
    End With
End Sub

Sub ColumnWidthIncrease()
' Recommended Shortcut: Ctrl+Q
    AdjustColumnWidth 1
End Sub

Sub ColumnWidthDecrease()
' Recommended Shortcut: Ctrl+Shift+Q
    AdjustColumnWidth -1
End Sub

Sub RowHeightIncrease()
' Recommended Shortcut: Ctrl+R
    AdjustRowHeight 5
End Sub

Sub RowHeightDecrease()
' Recommended Shortcut: Ctrl+Shift+R
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
