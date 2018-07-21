Sub Select_Values_to_Query_Column()
    With Worksheets("AutoQuery").Range("A" & CStr(5) & ":A" & CStr(100000))
        .Select
    End With
End Sub 'end of select values to query column
