Attribute VB_Name = "JAS_CustomFunctions"
Function DistanceToLine(LineXs As Variant, LineYs As Variant, PointX As Variant, PointY, Optional SegmentTF As Boolean = False) As Variant

''''''''''''''''''''''''''''''''
''Created by JAS on 12/17/2018''
''''''''''''''''''''''''''''''''

    
    Dim X_rows, X_columns, Y_rows, Y_columns As Integer
    X_rows = LineXs.Rows.Count: X_columns = LineXs.Columns.Count
    Y_rows = LineYs.Rows.Count: Y_columns = LineYs.Columns.Count
    
'Input range error checking

    If X_columns * X_rows <> 2 Then
        MsgBox ("Error! The LineXs range must include exactly two cells.")
        Return
    End If

    
    If Y_columns * Y_rows <> 2 Then
        MsgBox ("Error! The LineXs range must include exactly two cells.")
        Return
    End If
    
    Dim x_0, y_0, x, y As Variant: ReDim x(1 To 2), y(1 To 2)
    
    x_0 = PointX: y_0 = PointY
    x(1) = Application.Min(LineXs): x(2) = Application.Max(LineXs)
    y(1) = Application.Index(LineYs, Application.Match(x(1), LineXs, True))
    y(2) = Application.Index(LineYs, Application.Match(x(2), LineYs, True))
    
'Calculate line coefficients

    Dim a, b, c As Variant
    a = (y(2) - y(1)) / (x(2) - x(1)): b = -1: c = y(1) - a * x(1)
    
'Identify closest point on line

    Dim x_perp, y_perp As Variant
    x_perp = (b * (b * x_0 - a * y_0) - a * c) / (a * a + b * b)
    y_perp = (a * (-b * x_0 + a * y_0) - b * c) / (a * a + b * b)
    
    If SegmentTF = True Then
    
        If x_perp > x(2) Then
            x_perp = x(2): y_perp = y(2)
            
        ElseIf x_perp < x(1) Then
            x_perp = x(1): y_perp = y(1)
        
        End If
    
    End If
    
'Calculate distance to infinite line
    dist_inf = ((x_perp - x_0) ^ 2 + (y_perp - y_0) ^ 2) ^ 0.5
        
DistanceToLine = dist_inf

End Function
