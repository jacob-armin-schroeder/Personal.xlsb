Attribute VB_Name = "JAS_AutoFill"
Sub AutoFill()
Attribute AutoFill.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' AutoFill Macro
' Created on February 8, 2017 by Jacob Schroeder

' Keyboard Shortcut: Ctrl+Shift+D

'This macro fills a cell's contents and formatting down the length of an adjacent column, with
'priority given to the column to the left of the active cell.

Dim ColumnID As Integer

ColumnID = ActiveCell.Column

Select Case ColumnID

    Case 1
        
        'Checks to see if the column to the RIGHT extends TWO rows beyond the active cell
        If Not ActiveCell.Offset(2, 1).Value = "" And Not ActiveCell.Offset(1, 1).Value = "" Then
            
            Selection.Offset(0, 1).Select   'Same as above, for column to the right
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Offset(0, -1).Select
            Selection.FillDown

        'Checks to see if the column to the RIGHT only extends ONE row beyond the active cell
        ElseIf ActiveCell.Offset(2, 1).Value = "" And Not ActiveCell.Offset(1, 1).Value = "" Then
            
            Range(Selection, Selection.Offset(1, 0)).Select
            Selection.FillDown
    
        End If
    
    Case Else

    'First checking to see if the column to the LEFT extends TWO rows beyond the active cell
    If Not ActiveCell.Offset(2, -1).Value = "" And Not ActiveCell.Offset(1, -1).Value = "" Then
    
            Selection.Offset(0, -1).Select                  'Selects cell to the left of the active cell
            Range(Selection, Selection.End(xlDown)).Select  'Extends selection to bottom of left column
            Selection.Offset(0, 1).Select       'Offsets entire selection back to the column to be filled
            Selection.FillDown  'Fills the originally selected cell down the length of the column to the left
        
        'Checks to see if the column to the LEFT only extends ONE row beyond the active cell
        ElseIf ActiveCell.Offset(2, -1).Value = "" And Not ActiveCell.Offset(1, -1).Value = "" Then
            
            Range(Selection, Selection.Offset(1, 0)).Select 'Extends the selection downward by one row
            Selection.FillDown  'Fills the originally selected cell down one row
            
                'Checks to see if the column to the RIGHT extends TWO rows beyond the active cell
        ElseIf Not ActiveCell.Offset(2, 1).Value = "" And Not ActiveCell.Offset(1, 1).Value = "" Then
            
            Selection.Offset(0, 1).Select   'Same as above, for column to the right
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Offset(0, -1).Select
            Selection.FillDown

        'Checks to see if the column to the RIGHT only extends ONE row beyond the active cell
        ElseIf ActiveCell.Offset(2, 1).Value = "" And Not ActiveCell.Offset(1, 1).Value = "" Then
            
            Range(Selection, Selection.Offset(1, 0)).Select
            Selection.FillDown
    
        End If

End Select
    
End Sub

