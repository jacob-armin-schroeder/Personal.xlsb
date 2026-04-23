Attribute VB_Name = "JAS_NewSENSEIUpload"
Sub New_SENSEI_Upload()
'
' New_SENSEI_Upload Macro
'

'
'First checks to make sure columns A through C are empty. This prevents overwriting
'any data or cell formatting.

'If Application.CountA(Range("A:C")) = 0 Then
    
    'Entering the column headers in row 1
    
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "ImportCode"
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "Timestamp"
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "Value"
    
    'Setting columns to the required SENSEI upload format
    
        Columns("A:A").Select
        Selection.NumberFormat = "General"
        Columns("B:B").Select
        Selection.NumberFormat = "yyyy-mm-dd hh:mm:ss"
        Columns("C:C").Select
        Selection.NumberFormat = "General"
        
    'Increasing width of the first two columns to an appropriate size
        
        Columns("A:B").Select
        Selection.ColumnWidth = 18
        
    'Activates the first cell available for input data
        
        Range("A2").Select

'If columns A through C are not empty, the macro stops.

'Else
'End If

End Sub

