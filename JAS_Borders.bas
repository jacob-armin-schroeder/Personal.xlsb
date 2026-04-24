Attribute VB_Name = "JAS_Borders"
'
Option Explicit
' Border_Table_Heading Macro
'Created 10/6/2015 by Jacob Schroeder
'Revised 4/24/2026 by Jacob Schroeder
'
' Keyboard Shortcut: Ctrl+h
'
Sub Border_Table_Heading()
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = xlHorizontal
        .Font.Bold = True
        .Borders.LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .BorderAround Weight:=xlThin, ColorIndex:=xlAutomatic
    End With
End Sub

' BorderLinesVertical Macro
' Edited 4/24/2026 by Jacob A. Schroeder
'
' Keyboard Shortcut: Ctrl+e
'
Sub VerticalLines()
    Dim b As Border
    Set b = Selection.Borders(xlInsideVertical)
    If b.LineStyle = xlNone Then
        b.Weight = xlThin
    ElseIf b.Weight = xlThin Then
        b.Weight = xlMedium
    Else
        b.LineStyle = xlNone
    End If
End Sub

' BorderLinesHorizontal Macro
' Created 5/30/2018 by Jacob Schroeder
' Edited 4/24/2026 by Jacob Schroeder
'
' Keyboard Shortcut: Ctrl+r
'
Sub HorizontalLines()
    Dim b As Border
    Set b = Selection.Borders(xlInsideHorizontal)
    ' Cycles hairline > thin > medium > none
    ' Hairline is included for horizontal only; useful for dense row data
    If b.LineStyle = xlNone Then
        b.Weight = xlHairline
    ElseIf b.Weight = xlHairline Then
        b.Weight = xlThin
    ElseIf b.Weight = xlThin Then
        b.Weight = xlMedium
    Else
        b.LineStyle = xlNone
    End If
End Sub

' BorderLinesOutline Macro
' Macro edited 4/24/2026 by Jacob Schroeder
'
' Keyboard Shortcut: Ctrl+o
'
Sub Border_Outline()
  With Selection
    If .Borders(xlEdgeLeft).LineStyle = xlNone Then
        .BorderAround Weight:=xlThin
    ElseIf .Borders(xlEdgeLeft).Weight = xlThin Then
        .BorderAround Weight:=xlMedium
    Else
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
    End If
  End With
End Sub

' Border_Remove_All Macro
' Edited 4/24/2026 by Jacob Schroeder
'
' Keyboard Shortcut: Ctrl+n
'
Sub Border_Remove_All()
' Clears all borders and fill from selection
    Selection.Borders.LineStyle = xlNone
    With Selection
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.ColorIndex = xlAutomatic
    End With
End Sub

Sub FillBright()
' Cycles selection background through light colors, then clears.
' Font is always set to automatic.
' Cycle: None > #ECECEC > #BFE9FF > #FDEAD7 > #DCEFD8 > #FFFFFF > None

    Dim colors(0 To 5) As Long
    colors(0) = -1                      ' Sentinel for "no fill"
    colors(1) = RGB(236, 236, 236)      ' #ECECEC
    colors(2) = RGB(191, 233, 255)      ' #BFE9FF
    colors(3) = RGB(253, 234, 215)      ' #FDEAD7
    colors(4) = RGB(220, 239, 216)      ' #DCEFD8
    colors(5) = RGB(255, 255, 255)      ' #FFFFFF

    Dim currentIndex As Integer
    currentIndex = GetColorIndex(colors)

    Dim nextIndex As Integer
    nextIndex = (currentIndex + 1) Mod 6

    If nextIndex = 0 Then
        Selection.Interior.Pattern = xlNone
    Else
        Selection.Interior.Color = colors(nextIndex)
    End If

    Selection.Font.ColorIndex = xlAutomatic

End Sub

Sub FillDark()
' Cycles selection background through dark colors, then clears.
' Font is set to white for all dark colors; automatic when fill is removed.
' Cycle: None > #262626 > #005677 > #D6700A > #417A34 > #000000 > None

    Dim colors(0 To 5) As Long
    colors(0) = -1                      ' Sentinel for "no fill"
    colors(1) = RGB(38, 38, 38)         ' #262626
    colors(2) = RGB(0, 86, 119)         ' #005677
    colors(3) = RGB(214, 112, 10)       ' #D6700A
    colors(4) = RGB(65, 122, 52)        ' #417A34
    colors(5) = RGB(0, 0, 0)            ' #000000

    Dim currentIndex As Integer
    currentIndex = GetColorIndex(colors)

    Dim nextIndex As Integer
    nextIndex = (currentIndex + 1) Mod 6

    If nextIndex = 0 Then
        Selection.Interior.Pattern = xlNone
        Selection.Font.ColorIndex = xlAutomatic
    Else
        Selection.Interior.Color = colors(nextIndex)
        Selection.Font.Color = RGB(255, 255, 255)
    End If

End Sub

Private Function GetColorIndex(colors() As Long) As Integer
' Returns the index of the active cell's current fill color within the provided
' color array. Returns 0 (the "no fill" position) if the fill is absent or
' does not match any color in the array, so the next cycle step is always valid.

    Dim i As Integer

    If ActiveCell.Interior.Pattern = xlNone Then
        GetColorIndex = 0
        Exit Function
    End If

    Dim currentColor As Long
    currentColor = ActiveCell.Interior.Color

    For i = 1 To UBound(colors)
        If currentColor = colors(i) Then
            GetColorIndex = i
            Exit Function
        End If
    Next i

    ' Current fill color is not in this cycle; treat as unrecognized and
    ' restart from the beginning of the cycle on next call.
    GetColorIndex = 0

End Function
