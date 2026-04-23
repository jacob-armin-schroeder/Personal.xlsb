Attribute VB_Name = "JAS_PasteMacros"
Sub PasteTextOnly()
Attribute PasteTextOnly.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' PasteTextOnly Macro
'
' Keyboard Shortcut: Ctrl+Shift+V
'
On Error GoTo nonexcel

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Exit Sub

nonexcel:
    Dim DataObj As MSForms.DataObject
    Set DataObj = New MSForms.DataObject
    DataObj.GetFromClipboard

    Selection.Value = DataObj.GetText(1)
End Sub
