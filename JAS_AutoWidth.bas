Attribute VB_Name = "JAS_AutoWidth"
Sub AutoWidth()
Attribute AutoWidth.VB_ProcData.VB_Invoke_Func = "W\n14"
'
' Macro created 4/3/2018 by Jacob A. Schroeder

    With Selection
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
    End With
End Sub
Sub WidthIncrease()
Attribute WidthIncrease.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Macro created 4/8/2019 by Jacob A. Schroeder

    With Selection
        w = ActiveCell.ColumnWidth
            'Debug.Print w
        
        .ColumnWidth = Application.Round(w + 1, 0)
            'w = ActiveCell.Width
            'Debug.Print
    End With
    
End Sub
Sub WidthDecrease()
Attribute WidthDecrease.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' Macro created 4/8/2019 by Jacob A. Schroeder

    With Selection
        w = ActiveCell.ColumnWidth
            'Debug.Print w
        
        .ColumnWidth = Application.Round(Application.Max(w - 1, 1), 0)
            'w = ActiveCell.Width
            'Debug.Print
    End With
    
End Sub
Sub HeightIncrease()
Attribute HeightIncrease.VB_ProcData.VB_Invoke_Func = "j\n14"
'
' Macro created 4/8/2019 by Jacob A. Schroeder

    With Selection
        h = ActiveCell.RowHeight
            'Debug.Print h
        
        .RowHeight = Application.Round(h + 5, 0)
            'h = ActiveCell.Height
            'Debug.Print
    End With
    
End Sub
Sub HeightDecrease()
Attribute HeightDecrease.VB_ProcData.VB_Invoke_Func = "J\n14"
'
' Macro created 4/8/2019 by Jacob A. Schroeder

    With Selection
        h = ActiveCell.RowHeight
            'Debug.Print h
        
        .RowHeight = Application.Round(Application.Max(h - 5, 5), 0)
            'h = ActiveCell.Width
            'Debug.Print
    End With
    
End Sub

