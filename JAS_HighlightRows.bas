Attribute VB_Name = "JAS_HighlightRows"
Sub HighlightRows()
Attribute HighlightRows.VB_ProcData.VB_Invoke_Func = "H\n14"
'
' Highlight Rows Macro
'Created by Jacob Schroeder on 2/14/2017
'
'Allows user to select a color and then highlights the table in alternating three-row
'bands of the default color (no fill) and the selected color

Application.ScreenUpdating = False
    Dim i, j, n As Integer, FinalRange As Range
    Set FinalRange = Selection
    Selection.CurrentRegion.Select
    n = Selection.Rows.Count

    Dim intResult As Long
    Application.Dialogs(xlDialogEditColor).Show (30)
    intResult = ActiveWorkbook.Colors(30)
  
    For j = 5 To 7
        i = j
        Selection.CurrentRegion.Select
        Range(ActiveCell.Rows(i), ActiveCell.Rows(i).End(xlToRight)).Select
        
        With Selection.Interior
           .Color = intResult

        End With
        Do Until i > n - 6
            Selection.Offset(6, 0).Select
            With Selection.Interior
                .Color = intResult
            End With
            i = i + 6
        Loop
    Next
    FinalRange.Select
Application.ScreenUpdating = True
End Sub
