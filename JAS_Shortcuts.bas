Attribute VB_Name = "JAS_Shortcuts"
Sub NumberFormatDecimal()
Attribute NumberFormatDecimal.VB_Description = "Macro created on 3/7/2017 by Jacob Schroeder"
Attribute NumberFormatDecimal.VB_ProcData.VB_Invoke_Func = "A\n14"
'
' Macro1 Macro
' Macro created 3/7/2017 by Jacob Schroeder
'
' Keyboard Shortcut: Ctrl+Shift+A
'
  Dim x, myFormats
  
  myFormats = Array("#,##0", "#,##0.0", "#,##0.00", "#,##0.000")
  With ActiveCell
    x = Application.Match(.NumberFormat, myFormats, False)
    If IsError(x) Then x = 0
    Selection.NumberFormat = myFormats((x Mod (UBound(myFormats) + 1)))
  End With

End Sub

Sub NumberFormatPercentage()
Attribute NumberFormatPercentage.VB_Description = "Macro created on 3/7/2017 by Jacob Schroeder"
Attribute NumberFormatPercentage.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' NumberFormatPercentageOneDecimal Macro
' Macro created 3/7/2017 by Jacob Schroeder
'
' Keyboard Shortcut: Ctrl+Shift+P
'
  Dim x, myFormats
  
  myFormats = Array("#,##0%", "#,##0.0%", "#,##0.00%", "#,##0.000%")
  With ActiveCell
    x = Application.Match(.NumberFormat, myFormats, False)
    If IsError(x) Then x = 0
    Selection.NumberFormat = myFormats((x Mod (UBound(myFormats) + 1)))
  End With
  
End Sub
Sub NumberFormatDateTime()
Attribute NumberFormatDateTime.VB_Description = "Created on 3/7/2017 by Jacob Schroeder"
Attribute NumberFormatDateTime.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' NumberFormatPercentageOneDecimal Macro
' Macro created 3/7/2017 by Jacob Schroeder
'
' Keyboard Shortcut: Ctrl+Shift+T
'
  Dim x, myFormats
  
'Specifying the list of formats to cycle through
    
    myFormats = Array("m/d/yyyy", "m/d/yy", "mm/dd/yyyy", "m/d/yy h:mm", "mm/dd/yyyy hh:mm", "hh:mm", "yyyy-mm-dd hh:mm:ss")
  
'Checking current formatting and applying new formatting
    
    With ActiveCell
              
        'Identifying the index number "x" of the format above that matches the active cell's current format
            x = Application.Match(.NumberFormat, myFormats, False)
        
        'If the cell's current format doesn't match any in the list above, the index "x" is set to 0.
        
            If IsError(x) Then x = 0
        
        'Changing the format of the selected range to format "x+1"
        'i.e., if x = 0, the first format in the list will be applied to the selection
        
            Selection.NumberFormat = myFormats((x Mod (UBound(myFormats) + 1)))
    
    End With
  
End Sub
Sub TogglePageBreaks()
Attribute TogglePageBreaks.VB_Description = "Created on 10/7/2015 by Jacob Schroeder"
Attribute TogglePageBreaks.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' HidePageBreaks Macro
' Macro recorded 10/7/2015 by Jacob Schroeder

' Keyboard Shortcut: Ctrl+p
'
    
'This is a simple toggle macro. Whatever the current status is, it sets it to the opposite ("Not")
    
    ActiveSheet.DisplayPageBreaks = Not (ActiveSheet.DisplayPageBreaks)

End Sub
