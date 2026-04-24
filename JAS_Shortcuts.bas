End Sub
Sub TogglePageBreaks()
' TogglePageBreaks Macro
' Keyboard Shortcut: Ctrl+p
'This simple macro toggles the current DisplayPageBreaks status.
    
    ActiveSheet.DisplayPageBreaks = Not (ActiveSheet.DisplayPageBreaks)

End Sub

Sub ToggleGridlines()
'
' TogglePageBreaks Macro
' Keyboard Shortcut: Ctrl+Shift+P
' This simple macro toggles the current DisplayGridlines status.

    ActiveWindow.DisplayGridlines = Not (ActiveWindow.DisplayGridlines)
    
End Sub
