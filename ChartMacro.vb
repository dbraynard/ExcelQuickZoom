Sub Left_Scroll()
    Scroll ("left")
End Sub

Sub Right_Scroll()
    Scroll ("right")
End Sub


Sub Scroll(direction As String)

    'reference the scroll amount
    Dim scroll_amount As Integer
    scroll_amount = CInt(ActiveSheet.Range("Scroll_Amount"))
    
    'MsgBox ("scrolling " & scroll_amount)
    
    If ActiveChart Is Nothing Then
        MsgBox ("No chart is selected")
        Exit Sub
    End If
      
    'only adjust the x range once
    Dim firstIteration As Boolean
       
    firstIteration = True
       
    
    For Each ser In ActiveChart.SeriesCollection
    ' Perform desired processing on each item.
        
        'The following code uses the split function because VBA does seem to have lookbehind/lookahead in
        'its regex library
        
        Dim newFormula As String
        
        If (StrComp(direction, "left") = 0) Then
            'left scroll
            newFormula = ChangeFormulaXRange(ser.formula, -1 * scroll_amount, -1 * scroll_amount, "adjust", firstIteration)
        Else
            'right scroll
            newFormula = ChangeFormulaXRange(ser.formula, scroll_amount, scroll_amount, "adjust", firstIteration)
        End If
                                        
        'MsgBox (newFormula)
        
        ser.formula = newFormula
        
        firstIteration = False
                        
    Next ser
    
    
End Sub

Sub Zoom_In()
    Zoom ("in")
End Sub

Sub Zoom_Out()
    Zoom ("out")
End Sub

Sub Zoom(direction As String)
    'reference the scroll amount
    Dim zoom_amount As Integer
    zoom_amount = CInt(Worksheets("analysis").Range("Zoom_Amount"))
    
    'MsgBox ("zooming " & zoom_amount)
    
    If ActiveChart Is Nothing Then
        MsgBox ("No chart is selected")
        Exit Sub
    End If
    
    'only adjust the x range once
    Dim firstIteration As Boolean
    
    firstIteration = True
        
    'loop all series for the active chart
    For Each ser In ActiveChart.SeriesCollection
    ' Perform desired processing on each item.
        
        'The following code uses the split function because VBA does seem to have lookbehind/lookahead in
        'its regex library
        
        Dim newFormula As String
        
        If (StrComp(direction, "in") = 0) Then
            'zoom in
            newFormula = ChangeFormulaXRange(ser.formula, zoom_amount, -1 * zoom_amount, "adjust", firstIteration)
        Else
            'zoom out
            newFormula = ChangeFormulaXRange(ser.formula, -1 * zoom_amount, zoom_amount, "adjust", firstIteration)
        End If
        
        'assign adjusted formula to the series of this loop iteration
        ser.formula = newFormula
        
        firstIteration = False
                        
    Next ser

End Sub

Sub Go_To_X_Range()
    
    'reference the scroll amount
    Dim set_x_start As Integer
    set_x_start = CInt(Worksheets("analysis").Range("Set_X_Start"))
    
    Dim set_x_end As Integer
    set_x_end = CInt(Worksheets("analysis").Range("Set_X_End"))
    
        
    If ActiveChart Is Nothing Then
        MsgBox ("No chart is selected")
        Exit Sub
    End If
    
    Dim firstIteration As Boolean
    
    firstIteration = True
    
    'loop all series for the active chart
    For Each ser In ActiveChart.SeriesCollection
    ' Perform desired processing on each item.
        
        'The following code uses the split function because VBA does seem to have lookbehind/lookahead in
        'its regex library
        
        Dim newFormula As String
        
        newFormula = ChangeFormulaXRange(ser.formula, set_x_start, set_x_end, "set", firstIteration)
               
        'assign adjusted formula to the series of this loop iteration
        ser.formula = newFormula
        
        firstIteration = False
                        
    Next ser

    
End Sub


Sub Set_Series_Thickness()
    
    'reference the thickness amount
    Dim set_thickness As Double
    set_thickness = CDbl(Worksheets("analysis").Range("Set_Thickness"))
        
        
    If ActiveChart Is Nothing Then
        MsgBox ("No chart is selected")
        Exit Sub
    End If

    'loop all series for the active chart
    For Each ser In ActiveChart.SeriesCollection
    ' Perform desired processing on each item.
        
        ser.Format.Line.Weight = set_thickness
              
                        
    Next ser
    
End Sub

Function ChangeFormulaXRange(formula As String, dStart As Integer, dEnd As Integer, operation As String, adjustXValues As Boolean) As String

    Dim majorParts() As String
    Dim xValues() As String
    Dim yValues() As String
    Dim hasXValues As Boolean
    
    'default false
    'hasXValues = False
    
    
    majorParts = Split(formula, ",")
    
    
    'x values are only set on one series (if all series have same x range)
    'this seems to have changed with MS Excel 2013
        
                        
    xValues = Split(majorParts(1), "$")
            
        
    Dim xStart As Integer
    Dim xEnd As Integer
        
    xStart = CInt(Left(xValues(2), Len(xValues(2)) - 1))
    xEnd = CInt(xValues(4))
                      
    'perform operation (only if this is the adjustXValues is true, i.e. first series in chart)
    If (StrComp(operation, "adjust") = 0) Then
        xStart = xStart + IIf(adjustXValues, dStart, 0)
        xEnd = xEnd + IIf(adjustXValues, dEnd, 0)
    ElseIf (StrComp(operation, "set") = 0) Then
        xStart = IIf(adjustXValues, dStart, xStart)
        xEnd = IIf(adjustXValues, dEnd, xEnd)
    End If
        
    'coerce xStart to be at least 2, so user doesn't zoom past the beginning and cause an error
    If (xStart < 2) Then
        xStart = 2
    End If
    
    'coerce xEnd isn't neccessary to avoid runtime errors but just needs to examine the sheet's value at that record
    'TODO: implement xEnd coercion
    
    yValues = Split(majorParts(2), "$")
        
    Dim yStart As Integer
    Dim yEnd As Integer
        
    yStart = CInt(Left(yValues(2), Len(yValues(2)) - 1))
    yEnd = CInt(yValues(4))
        
    'perform operation
    If (StrComp(operation, "adjust") = 0) Then
        yStart = yStart + dStart
        yEnd = yEnd + dEnd
    ElseIf (StrComp(operation, "set") = 0) Then
        yStart = dStart
        yEnd = dEnd
    End If
    
    'coerce yStart to be at least 2, so user doesn't zoom past the beginning and cause an error
    If (yStart < 2) Then
        yStart = 2
    End If
    
    'check if start is greater than end, in case of user mistake (i.e zoom too much)
    'just return the current formula
    If (yStart > yEnd) Then
        ChangeFormulaXRange = formula
        Exit Function
    End If
    
    'coerce yEnd isn't neccessary to avoid runtime errors but just needs to examine the sheet's value at that record
    'TODO: implement yEnd coercion
        
    
    'put it all back together
    ChangeFormulaXRange = _
        majorParts(0) _
        & "," & xValues(0) & "$" & xValues(1) & "$" & xStart & ":$" & xValues(3) & "$" & xEnd & "," _
        & yValues(0) & "$" & yValues(1) & "$" & yStart & ":$" & yValues(3) & "$" & yEnd _
        & "," & majorParts(3)
            

End Function
