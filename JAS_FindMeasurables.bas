Attribute VB_Name = "JAS_FindMeasurables"
Sub FindNextMeas()
Attribute FindNextMeas.VB_ProcData.VB_Invoke_Func = "m\n14"
'Shortcut: Ctrl + m

' Created by Jacob Schroeder on May 9, 2016
' This macro finds the NEXT entry in the active column with a value that is different
' from the initially-selected cell. It was intended as a quick way to navigate to the NEXT
' measureable in a SENSEI upload document.

    Application.ScreenUpdating = False
    
    Dim M_0, M_1 As String
    M_0 = ActiveCell.Value
    M_1 = M_0
    
    Selection.Offset(1, 0).Select
    If M_0 = "" And Not ActiveCell.Value = "" Then
        M_0 = ActiveCell.Value
        M_1 = ActiveCell.Value
    Else
       
        Selection.Offset(-1, 0).Select
        Do Until Not (M_0 = M_1) Or M_1 = ""
        
            Selection.Offset(1, 0).Select
            M_1 = ActiveCell.Value
        Loop
    
    End If
    
    Application.ScreenUpdating = True
    ActiveCell.Activate
    
End Sub

Sub FindPrevMeas()
Attribute FindPrevMeas.VB_ProcData.VB_Invoke_Func = "M\n14"
'Shortcut: Ctrl + Shift + M

' Created by Jacob Schroeder on May 9, 2016
' This macro finds the NEXT PREVIOUS entry in the active column with a value that is different
' from the initially-selected cell. It was intended as a quick way to navigate to the PREVIOUS
' measureable in a SENSEI upload document.

    Application.ScreenUpdating = False
    
    Dim M_0, M_1 As String
    M_0 = ActiveCell.Value
    M_1 = M_0
    
    Selection.Offset(-1, 0).Select
    If M_0 = "" And Not ActiveCell.Value = "" Then
        M_0 = ActiveCell.Value
        M_1 = ActiveCell.Value
    Else
       
        Selection.Offset(1, 0).Select
        Do Until Not (M_0 = M_1) Or M_1 = ""
        
            Selection.Offset(-1, 0).Select
            M_1 = ActiveCell.Value
        Loop
    
    End If
    
    Application.ScreenUpdating = True
    ActiveCell.Activate
    
End Sub

