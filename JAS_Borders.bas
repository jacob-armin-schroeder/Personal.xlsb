Attribute VB_Name = "JAS_Borders"
'
' Border_Table_Heading Macro
' Macro created 10/6/2015 by Jacob A. Schroeder
'
' Keyboard Shortcut: Ctrl+h
'
Sub Border_Table_Heading()
Attribute Border_Table_Heading.VB_ProcData.VB_Invoke_Func = "h\n14"
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = xlHorizontal
        .Font.Bold = True
    End With
    Selection.Borders(xlLeft).LineStyle = xlNone
    Selection.Borders(xlRight).LineStyle = xlNone
    Selection.Borders(xlTop).LineStyle = xlNone
    Selection.Borders(xlBottom).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.BorderAround Weight:=xlThin, ColorIndex:=xlAutomatic
End Sub
'
' BorderLinesVertical Macro
' Macro edited 5/30/2018 by Jacob A. Schroeder
'
' Keyboard Shortcut: Ctrl+e
'
Sub VerticalLines()
Attribute VerticalLines.VB_ProcData.VB_Invoke_Func = "e\n14"

  With Selection
    
    If .Borders(xlInsideVertical).LineStyle = xlNone Then
            .Borders(xlInsideVertical).Weight = xlThin
    
    ElseIf .Borders(xlInsideVertical).Weight = xlThin Then
            .Borders(xlInsideVertical).Weight = xlMedium
    
    Else: .Borders(xlInsideVertical).LineStyle = xlNone
    
    End If
  
  End With

End Sub

'
' BorderLinesHorizontal Macro
' Macro created 5/30/2018 by Jacob A. Schroeder
'
' Keyboard Shortcut: Ctrl+r
'
Sub HorizontalLines()
Attribute HorizontalLines.VB_ProcData.VB_Invoke_Func = "r\n14"

  With Selection
    
    If .Borders(xlInsideHorizontal).LineStyle = xlNone Then
            .Borders(xlInsideHorizontal).Weight = xlHairline
    
    ElseIf .Borders(xlInsideHorizontal).Weight = xlHairline Then
            .Borders(xlInsideHorizontal).Weight = xlThin
    
    ElseIf .Borders(xlInsideHorizontal).Weight = xlThin Then
            .Borders(xlInsideHorizontal).Weight = xlMedium
    
    Else: .Borders(xlInsideHorizontal).LineStyle = xlNone
    
    End If
  
  End With

End Sub
'
' BorderLinesOutline Macro
' Macro edited 5/30/2018 by Jacob A. Schroeder
'
' Keyboard Shortcut: Ctrl+o
'
Sub Border_Outline()
Attribute Border_Outline.VB_ProcData.VB_Invoke_Func = "o\n14"

  With Selection
    
    If .Borders(xlEdgeLeft).LineStyle = xlNone Then
                .BorderAround Weight:=xlThin
    
    ElseIf .Borders(xlEdgeLeft).Weight = xlThin Then
                .BorderAround Weight:=xlMedium
    
    Else
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
        
    End If
  
  End With

End Sub
'
' Border_Remove_All Macro
' Macro edited 10/6/2015 by Jacob A. Schroeder
'
' Keyboard Shortcut: Ctrl+n
'
Sub Border_Remove_All()
Attribute Border_Remove_All.VB_Description = "Macro recorded 6/28/96 by Marcus H. Wilcox"
Attribute Border_Remove_All.VB_ProcData.VB_Invoke_Func = "n\n0"
    Selection.Borders(xlLeft).LineStyle = xlNone
    Selection.Borders(xlRight).LineStyle = xlNone
    Selection.Borders(xlTop).LineStyle = xlNone
    Selection.Borders(xlBottom).LineStyle = xlNone
    Selection.BorderAround LineStyle:=xlNone
With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
End Sub
