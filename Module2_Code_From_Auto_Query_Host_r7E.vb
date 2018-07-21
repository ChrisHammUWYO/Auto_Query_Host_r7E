Sub Clear_All_Button()
    'declare vars
    Dim Lastrow As Integer
    
    'initialize vars
    Lastrow = Worksheets("AutoQuery").Cells(Rows.Count, 1).End(xlUp).Row
    
    'With Worksheets("AutoQuery").Range("A" & CStr(5) & ":AA" & CStr(100000))
    If (Lastrow > 4) Then
        With Worksheets("AutoQuery").Rows("5:" & CStr(Lastrow))
            .ClearContents: .Interior.ColorIndex = 0
        End With
    End If
    'scroll to the top of the sheet
    With ActiveWindow
        .ScrollRow = 1: .ScrollColumn = 1
    End With
End Sub 'end of clear all button
