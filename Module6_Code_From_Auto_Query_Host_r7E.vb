Sub Auto_Update_Host_Button()
    'prompt the Are you sure meessage box
    confirmationMessageBox = MsgBox("Are you sure you want to update the name/SSN/address of all the" & vbNewLine & "existing members in this list?", vbYesNo, "Are You Sure?")
    
    Select Case confirmationMessageBox
    Case vbYes
        MsgBox ("NOTICE:" & vbNewLine & "If an error occurs, a popup will display asking you to manually change the problematic SSN." & vbNewLine & vbNewLine & "IMPORTANT:" & vbNewLine & "Write down the number and manually update after the spreadsheet is finished." & vbNewLine & "Updating will resume after you click 'ok' on the Message Box.")
    Case vbNo
        Exit Sub 'quit the subroutine
    End Select
    'end of prompt the are you sure message box
    
    'declare vars
    Dim valueToQuery As String
    Dim valueType As String
    Dim index As Integer
    Dim finished As Boolean
    Dim replacementName As String
    Dim replacementSSN As String
    Dim replacementStreetAddress As String
    Dim replacementCityStateAddress As String
    Dim replacementZipAddress As String
    Dim existsOnHost As String
    
    'initialize vars
    With Worksheets("AutoQuery").Range("A2")
        valueType = Trim(.Value)
    End With
    index = 5 'start at row 5
    finished = False
    replacementName = ""
    replacementSSN = ""
    replacementStreetAddress = ""
    replacementCityStateAddress = ""
    replacementZipAddress = ""
    existsOnHost = ""

    Do While (finished = False)
         'get value to be queried
         With Worksheets("AutoQuery").Range("A" & CStr(index))
             valueToQuery = Trim(.Value)
         End With
         
         'check to see if value to query is blank
         If (StrComp(vbNullString, valueToQuery) = 0) Then 'if blank
             finished = True
             Exit Do
         End If
         
         'check to make sure that SSN/TIN or CIS valueType is being used
         If (StrComp("SSN/TIN", Trim(valueType)) = 0) Then
             'do nothing
         ElseIf (StrComp("CIS", Trim(valueType)) = 0) Then
             'do nothing
         Else
             'error out
             MsgBox "Error: only SSN/TIN and CIS valueTypes are supported"
             finished = True
             Exit Do
         End If 'end of make sure valueType is supported
         
         'check to see if member exists on host,if not , do not update
         With Worksheets("AutoQuery").Range("B" & CStr(index))
             existsOnHost = Trim(.Value)
         End With
         If (StrComp("False", existsOnHost) = 0) Then 'if not on host, move onto next member
             'do not change any values proceed to next line
         Else 'if does exist on host
            'Update Address-----------------------------------------
            'get replacement address then call function to change it
            With Worksheets("AutoQuery").Range("D" & CStr(index))
                replacementStreetAddress = Trim(.Value)
            End With
            With Worksheets("AutoQuery").Range("E" & CStr(index))
                replacementCityStateAddress = Trim(.Value)
            End With
            With Worksheets("AutoQuery").Range("F" & CStr(index))
                replacementZipAddress = Trim(.Value)
            End With
            Call changeAddress(valueType, valueToQuery, replacementStreetAddress, replacementCityStateAddress, replacementZipAddress)
            'End of Update Address----------------------------------
            
            'Update Name--------------------------------------------
            'get replacement name then call function to change it
            With Worksheets("AutoQuery").Range("C" & CStr(index))
                replacementName = Trim(.Value)
            End With
            Call changeName(valueType, valueToQuery, replacementName)
            'end of update name--------------------------------------
            
            'Update SSN-------------------------------------------
            'get replacement SSN/TIN, then call function to change it
            With Worksheets("AutoQuery").Range("I" & CStr(index))
                replacementSSN = Trim(.Value)
            End With
            Call changeSSN(valueType, valueToQuery, replacementSSN, index)
            'end of update ssn------------------------------------
         End If 'end if existsOnHOst field is false
         
         'increment index
         index = index + 1
    Loop 'end of loop
    
    MsgBox "Auto Update Has Finished. You may need to requery member(s) to see the new name/SSN/Address."

End Sub 'end of Auto Update Host

Public Function changeName(ByVal valueType As String, ByVal valueToQuery As String, ByVal replacementName As String)
    
    'declare vars--------------------------------
    Dim finished As Boolean
    Dim currentScreen As String
    Dim nameOnAccount As String 'used to store the name on account, if no name, then skip changing the name
    Dim titleOnScreen As String 'used if there is no screen name to use
    Dim recordNotFoundError As String
    'end of declare vars--------------------------
    
    'initialize vars------------------------------
    finished = False
    currentScreen = ""
    nameOnAccount = ""
    titleOnScreen = ""
    recordNotFoundError = ""
    'end of initialize vars-----------------------
    
    Call SetHost 'make sure host is set

    'reset the Host screen
    Call ResetForCommand
    
    'trim valueToQuery and valueType
    valueToQuery = Trim(valueToQuery)
    valueType = Trim(valueType)
    
    'navigate to the change name screen
    If (StrComp(valueType, "CIS") = 0) Then  'if CIS
        HostSendCommand ("rmcn;nb" & valueToQuery)
    Else 'if SSN/TIN
        HostSendCommand ("rmcn;tn" & valueToQuery)
    End If 'end of if inputValueType is CIS
    
    'RMCN fields---------------------------
    'Name Action (used to perform update) - field 7
    'SSN/TIN - Field 13
    'Name - field 18
    'P/C/D - field 20
    
    CurrentHost.Fields(7) = Trim("U")
    CurrentHost.Fields(18) = Trim(replacementName)
    
    'send the enter keystroke to host
    HostSendKeys ("@E")
    
    
    'UPDATE ALL SCREENS---------------------------------------------------
    Do While (finished = False)
    'load current screen
    currentScreen = Trim(CurrentHost.Fields(1))
    titleOnScreen = Trim(CurrentHost.Fields(3))
    
    'exit loop if rmab screen
    If (StrComp("RMAB", currentScreen) = 0) Then
        finished = True
        Exit Do
    ElseIf (StrComp("RMNC", currentScreen) = 0) Then
        'RMNC screen fields----------------------
        'Bypass name std - field 23
        'First y/n in Chg? column - field 35
        'Account Name for first y/n - field 46
        'Second y/n in Chg? column - field 52
        'Account Name for Second y/n - field 63
        '2nd Account Name for Second y/n - field 66
        '3rd Account Name for Second y/n - field 67
        'Process Complete message - field 130
        
        'Update RMNC screen---------------------
        'check to make sure name os on first account
        nameOnAccount = CurrentHost.Fields(46)
        If (StrComp("NO NAME PRESENT ON APPLICATION", Trim(nameOnAccount)) = 0) Then
            CurrentHost.Fields(35) = Trim("n")
        Else 'if there is a name
            CurrentHost.Fields(35) = Trim("y")
        End If 'end of no name on first account
        
        'check to make sure name os on second account
        nameOnAccount = CurrentHost.Fields(63)
        If (StrComp("NO NAME PRESENT ON APPLICATION", Trim(nameOnAccount)) = 0) Then
            CurrentHost.Fields(52) = Trim("n")
        Else 'if there is a name
            CurrentHost.Fields(52) = Trim("y")
        End If 'end of no name on second account
        
        'press enter to move to next screen
        HostSendKeys ("@E")
        'End of RMNC screen update--------------
    ElseIf (StrComp("AMC5", currentScreen) = 0) Then
        'AMC5 screen fields-----------------------
        'Y/N value for First prompt (change or bypass) - field 6
        'Y/N value for second prompt (update related accounts) - field 21
        
         'check for record not found error
        recordNotFoundError = Trim(CurrentHost.Fields(CurrentHost.Fields.Count))
        If recordNotFoundError Like "*REC*NOT*FND*" Then
            finished = True
            Exit Do
        End If
        'end of check for record not found error
        
        'Update AMC5 screen --------------------
        CurrentHost.Fields(6) = Trim("y")
        
        HostSendKeys ("@E") 'press enter to bring up second prompt
        CurrentHost.Fields(21) = Trim("y")
        
        'press enter to move to next screen
        HostSendKeys ("@E")
        'End of Update AMC5 screen--------------
    ElseIf (StrComp("IM02", currentScreen) = 0) Then
        'IM02 screen fields-------------------------
        'Y/N value for 'Enter Y to change account, N to Bypass Changes' - field 114
        
        'check for record not found error
        recordNotFoundError = Trim(CurrentHost.Fields(118))
        If recordNotFoundError Like "*REC*NOT*FND*" Then
            finished = True
            Exit Do
        End If
        'end of check for record not found error
        
        'Update IM02 screen ------------------------
        CurrentHost.Fields(114) = Trim("y")
        
        'press enter to move to next screen
        HostSendKeys ("@E")
        'End of Update IM02 screen------------------
    ElseIf (StrComp("ST21", currentScreen) = 0) Then
        'ST21 fields----------------------
        'Y/N value prompt (entr y to change account, n to bypass changes) - field 137
        
        'check for record not found error
        recordNotFoundError = Trim(CurrentHost.Fields(144))
        If recordNotFoundError Like "*REC*NOT*FND*" Then
            finished = True
            Exit Do
        End If
        'end of check for record not found error
        
        'Update ST21 screen---------------
        CurrentHost.Fields(137) = Trim("y")

        'press enter to move to next screen
        HostSendKeys ("@E")
        'End of updating ST21 screen------------
    ElseIf (StrComp("CUST PRIMARY NAME AND ADDR", Trim(titleOnScreen)) = 0) Then 'catches AMC5 screen that has no field to pull the screen name
        'Alternature AMC5 fields-----------------
        'Y/N value for First prompt (change or bypass) - field 6
        'Y/N value for second prompt (update related accounts) - field 21
        
        'check for record not found error
        recordNotFoundError = Trim(CurrentHost.Fields(124))
        If recordNotFoundError Like "*REC*NOT*FND*" Then
            finished = True
            Exit Do
        End If
        'end of check for record not found error
        
        'Update AMC5 Alternate screen---------------
         CurrentHost.Fields(6) = Trim("y")
        
        HostSendKeys ("@E") 'press enter to bring up second prompt
        CurrentHost.Fields(21) = Trim("y")
        
        'press enter to move to next screen
        HostSendKeys ("@E")
        'end of Update AMC5 alternate----------------
    ElseIf (StrComp("*** CREDIT CARD SYSTEM INQUIRY SCREEN ***", Trim(titleOnScreen)) = 0) Then 'catches the CCI1 screen
        'CCI1 Screen-------------------------
        'last field on screen - field 64
        
        'check for record not found error
        recordNotFoundError = Trim(CurrentHost.Fields(64))
        If recordNotFoundError Like "*REC*NOT*FND*" Then
            finished = True
            Exit Do
        End If
        'end of check for record not found error
        
        'press enter to move to next screen
        HostSendKeys ("@E")
        'End of CCI1 screen------------------
    Else 'if an unknown screen
        'ending RMCN screen-------------------------
        finished = True
        Exit Do
        'end of ending RMCN screen update---------------
    End If 'end of if rmab screen
    Loop 'end of loop
    'END OF UPDATE ALL SCREENS---------------------------------------
    
End Function 'end of changeName

Public Function changeSSN(ByVal valueType As String, ByVal valueToQuery As String, ByVal replacementSSN As String, ByVal index As Integer)
'NOTE: SINCE I CANT FIND HOW TO PRESS THE SHIFT KEY PLUS F1 IN CODE, ONLY ENTRIES ON THE FIRST PAGE CAN BE CHANGED

    'declare vars
    Dim IDCategory1 As String
    Dim IDCategory2 As String
    
    'initialize vars
    IDCategory1 = ""
    IDCategory2 = ""
    
    Call SetHost 'make sure host is set

    'reset the Host screen
    Call ResetForCommand

    'trim valueToQuery and valueType
    valueToQuery = Trim(valueToQuery)
    valueType = Trim(valueType)
    
    'navigate to the change SSN screen
    If (StrComp(valueType, "CIS") = 0) Then  'if CIS
        HostSendCommand ("rmc2;nb" & valueToQuery)
    Else 'if SSN/TIN
        HostSendCommand ("rmc2;tn" & valueToQuery)
    End If 'end of if inputValueType is CIS
     
    'RMC2 screen fields-------------------
    'action update - field 7
    'first id category - field 17  'if is TN then that is SSN/TIN
    'first id value - field 21
    'second id category - field 37  'if is TN then that is SSN/TIN
    'second id value - field 41
    'last field - field 162
    
    'Update RMC2 screen---------------------
    CurrentHost.Fields(7) = Trim("U") 'switch from I to U
    
    'depending on which id is the SSN/TIN change that one
    IDCategory1 = Trim(CurrentHost.Fields(17))
    IDCategory2 = Trim(CurrentHost.Fields(37))
    
    If (StrComp("TN", IDCategory1) = 0) Then
        CurrentHost.Fields(21) = Trim(replacementSSN)
        'if valueType is SSN/TIN update the value to query
        If (StrComp("SSN/TIN", Trim(valueType)) = 0) Then
            With Worksheets("AutoQuery").Range("A" & CStr(index))
                .Value = Trim(replacementSSN)
            End With
        End If
    ElseIf (StrComp("TN", IDCategory2) = 0) Then
        CurrentHost.Fields(41) = Trim(replacementSSN)
        'if valueType is SSN/TIN update the value to query
        If (StrComp("SSN/TIN", Trim(valueType)) = 0) Then
            With Worksheets("AutoQuery").Range("A" & CStr(index))
                .Value = Trim(replacementSSN)
            End With
        End If
    Else
        MsgBox "Error: SSN/TIN not listed on first page of RMC2. Please manually change SSN/TIN " & CStr(oldSSN) & " to new SSN/TIN " & CStr(replacementSSN)
    End If
    
    'press enter to move to next screen
        HostSendKeys ("@E")
    'end of update RMC2 screen--------------
    
    'End of RMC2 screen-------------------
    
End Function 'end of changeSSN

Public Function changeAddress(ByVal valueType As String, ByVal valueToQuery As String, ByVal replacementStreetAddress As String, ByVal replacementCityStateAddress As String, ByVal replacementZipAddress As String)
    'declare vars--------------------
    Dim replacementCityAddress As String
    Dim replacementStateAddress As String
    Dim finished As Boolean
    Dim currentScreen As String
    Dim titleOnScreen As String
    Dim addressOnAccount As String
    'end of declare vars-------------
    
    'initialize vars-----------------
    replacementCityAddress = ""
    replacementStateAddress = ""
    finished = False
    currentScreen = ""
    titleOnScreen = ""
    addressOnAccount = ""
    'end of initialize vars----------
    
    Call SetHost 'make sure host is set

    'reset the Host screen
    Call ResetForCommand
    
    'trim valueToQuery and valueType and replacementCityStateAddress
    valueToQuery = Trim(valueToQuery)
    valueType = Trim(valueType)
    replacementCityStateAddress = Trim(replacementCityStateAddress)
    
    'split replacementCityStateAddress string into city address and state address
    replacementCityAddress = Trim(Left(replacementCityStateAddress, Len(replacementCityStateAddress) - 2))
    replacementStateAddress = Trim(Right(replacementCityStateAddress, 2))
    
    'replace comma (if present) with a space
    replacementCityAddress = Trim(Replace(replacementCityAddress, ",", " "))
    
    'navigate to the change address screen
    If (StrComp(valueType, "CIS") = 0) Then  'if CIS
        HostSendCommand ("rmca;nb" & valueToQuery)
    Else 'if SSN/TIN
        HostSendCommand ("rmca;tn" & valueToQuery)
    End If 'end of if inputValueType is CIS
    
    'RMCA fields------------------------------
    'Addr Action - field 8 (switch to U from I)
    'Street Address - field 16
    'City Address - field 53
    'State Address - field 55
    'Bypass Std? (Not used) - field 58
    'Postal Code (Zip Address) - field 60
    
    CurrentHost.Fields(8) = Trim("U")
    CurrentHost.Fields(16) = Trim(replacementStreetAddress)
    CurrentHost.Fields(53) = Trim(replacementCityAddress)
    CurrentHost.Fields(55) = Trim(replacementStateAddress)
    CurrentHost.Fields(60) = Trim(replacementZipAddress)
    
    'send the enter keystroke to host
    HostSendKeys ("@E")
    
    'UPDATE ALL SCREENS---------------------------------------------------------
    Do While (finshed = False)
        'load the current screen
        currentScreen = Trim(CurrentHost.Fields(1))
        titleOnScreen = Trim(CurrentHost.Fields(3))
        
        'exit loop if rmab screen
        If (StrComp("RMAB", currentScreen) = 0) Then
            finished = True
            Exit Do
        ElseIf (StrComp("RMAC", currentScreen) = 0) Then
            'RMAC fields--------------------------
            'first y/n in chg? column - field 63
            'address of first account - field 66
            'second y/n in chg? column - field 76
            'address of second account - field 79
            'third y/n in chg? column - field 89
            'address of third account - field 92
            'fourth y/n in chg? column - field 102
            'address of the fourth account - field 105
            'error fiels - field 119
            
            'UPDATE RMAC fields------------------------
            'check to make sure first account has address
            addressOnAccount = Trim(CurrentHost.Fields(66))
            If (StrComp("NO ADDRESS PRESENT ON APPLICATION", Trim(addressOnAccount)) = 0) Or Trim(CurrentHost.Fields(119) Like "*DUPLICATE USE FOUND ON ADDR 001*") Then
                'if no address, set to N
                CurrentHost.Fields(63) = Trim("n")
            Else 'if address is present
                CurrentHost.Fields(63) = Trim("y")
            End If
            
            'check to make sure second account has address
            addressOnAccount = Trim(CurrentHost.Fields(79))
            If (StrComp("NO ADDRESS PRESENT ON APPLICATION", Trim(addressOnAccount)) = 0) Or Trim(CurrentHost.Fields(119) Like "*DUPLICATE USE FOUND ON ADDR 002*") Then
                'if no address, set to N
                CurrentHost.Fields(76) = Trim("n")
            Else 'if address is present
                CurrentHost.Fields(76) = Trim("y")
            End If
            
            'check to make sure third account has address
            addressOnAccount = Trim(CurrentHost.Fields(92))
            If (StrComp("NO ADDRESS PRESENT ON APPLICATION", Trim(addressOnAccount)) = 0) Or Trim(CurrentHost.Fields(119) Like "*DUPLICATE USE FOUND ON ADDR 003*") Then
                'if no address, set to N
                CurrentHost.Fields(89) = Trim("n")
            Else 'if address is present
                CurrentHost.Fields(89) = Trim("y")
            End If
            
            'check to make sure fourth account has address
            addressOnAccount = Trim(CurrentHost.Fields(105))
            If (StrComp("NO ADDRESS PRESENT ON APPLICATION", Trim(addressOnAccount)) = 0) Or Trim(CurrentHost.Fields(119) Like "*DUPLICATE USE FOUND ON ADDR 004*") Then
                'if no address, set to N
                CurrentHost.Fields(102) = Trim("n")
            Else 'if address is present
                CurrentHost.Fields(102) = Trim("y")
            End If
            
            'press enter to move to next screen
            HostSendKeys ("@E")
            'End of Update RMAC fields-----------------
        ElseIf (StrComp("AMC5", currentScreen) = 0) Then
            'AMC5 screen fields-----------------------
            'Y/N value for First prompt (change or bypass) - field 6
            'Y/N value for second prompt (update related accounts) - field 21
            
             'check for record not found error
            recordNotFoundError = Trim(CurrentHost.Fields(CurrentHost.Fields.Count))
            If recordNotFoundError Like "*REC*NOT*FND*" Then
                finished = True
                Exit Do
            End If
            'end of check for record not found error
            
            'Update AMC5 screen --------------------
            CurrentHost.Fields(6) = Trim("y")
            
            HostSendKeys ("@E") 'press enter to bring up second prompt
            CurrentHost.Fields(21) = Trim("y")
            
            'press enter to move to next screen
            HostSendKeys ("@E")
            'End of Update AMC5 screen--------------
        ElseIf (StrComp("IM02", currentScreen) = 0) Then
            'IM02 screen fields-------------------------
            'Y/N value for 'Enter Y to change account, N to Bypass Changes' - field 114
            
            'check for record not found error
            recordNotFoundError = Trim(CurrentHost.Fields(118))
            If recordNotFoundError Like "*REC*NOT*FND*" Then
                finished = True
                Exit Do
            End If
            'end of check for record not found error
            
            'Update IM02 screen ------------------------
            CurrentHost.Fields(114) = Trim("y")
            
            'press enter to move to next screen
            HostSendKeys ("@E")
            'End of Update IM02 screen------------------
        ElseIf (StrComp("RMCC", currentScreen) = 0) Then
            'RMCC screen Fields---------------------------
            'Y/N value for first account in Cng? column - field 57
            'address for first account - field 64
            'Y/N value for second account in Cng? column - field 65
            'address for second account - field 72
            'Y/N value for third account in Cng? column - field 73
            'address for third account - field 80
            'Y/N value for third account in Cng? column - field 81
            'address for third account - field 88
            'last field on screen - field 89/90/91 depending on # of accounts
            'TODO Missing other accounts from list
            'TODO add additional accounts to update section (below)
            
            'Update RMCC screen-------------------------
            'make sure first account has an address
            addressOnAccount = Trim(CurrentHost.Fields(64))
            If (StrComp("NO ADDRESS PRESENT ON APPLICATION", Trim(addressOnAccount)) = 0) Then
                'if no address present, set to N
                CurrentHost.Fields(57) = Trim("n")
            Else 'if address is present
                CurrentHost.Fields(57) = Trim("y")
            End If
            
            'make sure second account has an address
            addressOnAccount = Trim(CurrentHost.Fields(72))
            If (StrComp("NO ADDRESS PRESENT ON APPLICATION", Trim(addressOnAccount)) = 0) Then
                'if no address present, set to N
                CurrentHost.Fields(65) = Trim("n")
            Else 'if address is present
                CurrentHost.Fields(65) = Trim("y")
            End If
            
            'make sure third account has an address
            addressOnAccount = Trim(CurrentHost.Fields(80))
            If (StrComp("NO ADDRESS PRESENT ON APPLICATION", Trim(addressOnAccount)) = 0) Then
                'if no address present, set to N
                CurrentHost.Fields(73) = Trim("n")
            Else 'if address is present
                CurrentHost.Fields(73) = Trim("y")
            End If
            
            'make sure fourth account has an address
            addressOnAccount = Trim(CurrentHost.Fields(88))
            If (StrComp("NO ADDRESS PRESENT ON APPLICATION", Trim(addressOnAccount)) = 0) Then
                'if no address present, set to N
                CurrentHost.Fields(81) = Trim("n")
            Else 'if address is present
                CurrentHost.Fields(81) = Trim("y")
            End If
            
            'press enter to move to next screen
            HostSendKeys ("@E")
            'End of Update RMCC screen--------------------
        ElseIf (StrComp("ST21", currentScreen) = 0) Then
            'ST21 fields----------------------
            'Y/N value prompt (entr y to change account, n to bypass changes) - field 137
            
            'check for record not found error
            recordNotFoundError = Trim(CurrentHost.Fields(144))
            If recordNotFoundError Like "*REC*NOT*FND*" Then
                finished = True
                Exit Do
            End If
            'end of check for record not found error
            
            'Update ST21 screen---------------
            CurrentHost.Fields(137) = Trim("y")
    
            'press enter to move to next screen
            HostSendKeys ("@E")
            'End of updating ST21 screen------------
        ElseIf (StrComp("CUST PRIMARY NAME AND ADDR", Trim(titleOnScreen)) = 0) Then 'catches AMC5 screen that has no field to pull the screen name
            'Alternature AMC5 fields-----------------
            'Y/N value for First prompt (change or bypass) - field 6
            'Y/N value for second prompt (update related accounts) - field 21
            
            'check for record not found error
            recordNotFoundError = Trim(CurrentHost.Fields(124))
            If recordNotFoundError Like "*REC*NOT*FND*" Then
                finished = True
                Exit Do
            End If
            'end of check for record not found error
            
            'Update AMC5 Alternate screen---------------
             CurrentHost.Fields(6) = Trim("y")
            
            HostSendKeys ("@E") 'press enter to bring up second prompt
            CurrentHost.Fields(21) = Trim("y")
            
            'press enter to move to next screen
            HostSendKeys ("@E")
            'end of Update AMC5 alternate----------------
        ElseIf (StrComp("*** CREDIT CARD SYSTEM INQUIRY SCREEN ***", Trim(titleOnScreen)) = 0) Then 'catches the CCI1 screen
            'CCI1 Screen-------------------------
            'last field on screen - field 64
            
            'check for record not found error
            recordNotFoundError = Trim(CurrentHost.Fields(64))
            If recordNotFoundError Like "*REC*NOT*FND*" Then
                finished = True
                Exit Do
            End If
            'end of check for record not found error
            
            'press enter to move to next screen
            HostSendKeys ("@E")
            'End of CCI1 screen------------------
        Else
            finished = True
            Exit Do
        End If 'end of if rmab screen
    Loop 'end of while loop
    'End of UPDATE ALL SCREENS--------------------------------------------------
    
     'Dim i As Integer 'display the fields
    'For i = 1 To CurrentHost.Fields.Count - 1
        'column g is index, column h is value
    '    With Worksheets("AutoQuery").Range("G" & CStr(i + 5))
     '      .Value = CStr(i)
      '  End With
       ' With Worksheets("AutoQuery").Range("H" & CStr(i + 5))
        '    .Value = CStr(CurrentHost.Fields(i))
       'End With
    'Next i 'end of for loop
End Function 'end of changeAddress
