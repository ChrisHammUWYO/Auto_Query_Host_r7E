Sub Remove_Non_Existing_Values_To_Query()
    'After pressing the Auto Query Host button, pressing this button will remove any rows that have a value that does not exist on host.
    'disable features to optimize macro for speed
    Call disableFeaturesForMacroOptimization
    
    'declare vars-----------------------
    Dim index As Integer
    Dim finished As Boolean
    Dim analysisValue As String
    'end of declare vars----------------
    
    'initialize vars--------------------
    index = 5 'start at the first value to query row
    finished = False
    analysisValue = ""
    'end of initialize vars-------------
    
    'while loop to remove non existing values------------
    Do While (finished = False)
        'load cell value into analysisValue
        With Worksheets("AutoQuery").Range("B" & CStr(index))
            analysisValue = Trim(.Value)
        End With
        'check to see if row is empty, if so, exit loop
        If (StrComp(vbNullString, analysisValue) = 0) Then
            finished = True
            Exit Do
        'if not empty, then check to see if exists on host is false, if so delete row
        ElseIf (StrComp("False", analysisValue) = 0) Then
            'delete the row
            Worksheets("AutoQuery").Rows(index).Delete
            'decrement index to make sure to read the correct line after deleting
            index = index - 1
        Else
            'if true, do not remove the row, do nothing
        End If 'end of if analysisValue is the Null String
        
        'increment index
        index = index + 1
    Loop 'end of while loop
    'end of while loop to remove non existing values-----
    
    're-enable features once macro is finished
    Call enableFeaturesAfterOptimizedMacroIsDone
End Sub 'end of remove non exosting value to query
Sub Remove_Existing_Values_To_Query()
    'disable features to optimize macro for speed
    Call disableFeaturesForMacroOptimization
    
    'declare vars-----------------------
    Dim index As Integer
    Dim finished As Boolean
    Dim analysisValue As String
    'end of declare vars----------------
    
    'initialize vars--------------------
    index = 5 'start at the first value to query row
    finished = False
    analysisValue = ""
    'end of initialize vars-------------
    
    'while loop to remove existing values------------
    Do While (finished = False)
        'load cell value into analysisValue
        With Worksheets("AutoQuery").Range("B" & CStr(index))
            analysisValue = Trim(.Value)
        End With
        'check to see if row is empty, if so, exit loop
        If (StrComp(vbNullString, analysisValue) = 0) Then
            finished = True
            Exit Do
        'if not empty, then check to see if exists on host is true, if so delete row
        ElseIf (StrComp("True", analysisValue) = 0) Then
            'delete the row
            Worksheets("AutoQuery").Rows(index).Delete
            'decrement index to make sure to read the correct line after deleting
            index = index - 1
        Else
            'if false, do not remove the row, do nothing
        End If 'end of if analysisValue is the Null String
        
        'increment index
        index = index + 1
    Loop 'end of while loop
    'end of while loop to remove existing values-----
    
    're-enable features once macro is finished
    Call enableFeaturesAfterOptimizedMacroIsDone
End Sub 'end of remove existing values to query
