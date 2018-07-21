Sub Highlight_Selection()
    'declare vars
    Dim selectedRange As Range
    Dim highlightColor As String
    
    'initialize vars
    Set selectedRange = Selection
    With Worksheets("AutoQuery").Range("H2")
        highlightColor = CStr(Trim(.Value))
    End With
    
    'change color of selection
    With selectedRange
        If (StrComp("Yellow", highlightColor) = 0) Then
            .Interior.Color = vbYellow
        ElseIf (StrComp("Red", highlightColor) = 0) Then
            .Interior.Color = vbRed
        ElseIf (StrComp("Magenta", highlightColor) = 0) Then
            .Interior.Color = vbMagenta
        ElseIf (StrComp("Green", highlightColor) = 0) Then
            .Interior.Color = vbGreen
        ElseIf (StrComp("Cyan", highlightColor) = 0) Then
            .Interior.Color = vbCyan
        ElseIf (StrComp("Blue", highlightColor) = 0) Then
            .Interior.Color = vbBlue
            
        End If
    End With
    
    'clear the variable
    Set selectedRange = Nothing
End Sub 'end of highlight selection
Sub Clear_Selected_Highlighting()
    'declare vars
    Dim selectedRange As Range
    
    'initialize vars
    Set selectedRange = Selection
    
    'change color of selection
    With selectedRange
        .Interior.Color = vbWhite
    End With
    
    'clear the variable
    Set selectedRange = Nothing
End Sub 'end of clear selecteed highlighting
Sub Clear_All_Highlighting()
    'declare vars
    Dim Lastrow As Integer
    
    'initialize vars
    Lastrow = Worksheets("AutoQuery").Cells(Rows.Count, 1).End(xlUp).Row
    'clear highlighting
    If (Lastrow > 4) Then
        With Worksheets("AutoQuery").Rows("5:" & CStr(Lastrow))
            .Interior.Color = vbWhite
        End With
    End If
    'clear h2's highlighting
    With Worksheets("AutoQuery").Range("H2")
        .Interior.Color = vbWhite
    End With
    'scroll to the top of the sheet
    With ActiveWindow
        .ScrollRow = 1: .ScrollColumn = 1
    End With
End Sub 'end of clear all highlighting
