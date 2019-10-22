Sub internationalCenters()

'for optimization, don't refresh screen while running the macro
Application.ScreenUpdating = False

'define common variables
Dim recordCounter, innerLoopCounter, debugCode
'use the debug code to prevent data loss while testing, not to be used in production
debugCode = True

'setup worksheets for reference
Dim centersDB, exportedData As Worksheet
Set centersDB = Worksheets(2)
Set exportedData = Worksheets(1)

'setup terra dotta export data columns for reference
Dim exportedDataFirstname, exportedDataLastname, exportedData8x, exportedDataAge, exportedDataInstGPA, exportedDataOvGPA
Dim exportedDataOvHrs, exportedDataStatus, exportedDataAppDate, exportedDataProgram, exportedDataGA, exportedDataHonors
Dim exportedDataMajor1, exportedDataMajor2, exportedDataMajor3, exportedDataMinor1, exportedDataMinor2, exportedDataEmail
Dim exportedDataDegree, exportedDataLocalPhone, exportedDataLocAddress, exportedDataCriminal, exportedDataInstHrs
'setup the centers database columns
Dim centersFirstname, centersLastname, centers8x, centersAge, centersInstGPA, centersOvGPA, centersInstHrs, centersOvHrs
Dim centersStatus, centersAppDate, centersProgram, centersGA, centersHonors, centersMajor1, centersMajor2, centersMajor3
Dim centersMinor1, centersMinor2, centersEmail, centersDegree, centersLocPhone, centersLocAddress, centersCriminal

'assign column letter to each exported data field title
exportedDataFirstname = "B"
exportedDataLastname = "C"
exportedDataMiddlename = "D"
exportedData8x = "A"
exportedDataAge = "F"
exportedDataInstGPA = "G"
exportedDataOvGPA = "H"
exportedDataInstHrs = "J"
exportedDataOvHrs = "K"
exportedDataStatus = "M"
exportedDataAppDate = "N"
exportedDataProgram = "O"
exportedDataGA = "S"
exportedDataHonors = "T"
exportedDataMajor1 = "U"
exportedDataMajor2 = "V"
exportedDataMajor3 = "W"
exportedDataMinor1 = "X"
exportedDataMinor2 = "Y"
exportedDataEmail = "Z"
exportedDataNickname = "AB"
exportedDataDegree = "AH"
exportedDataLocalPhone = "S"
exportedDataLocAddress = "AS"
'assign numerical position for the centers data
centersLastname = 22
centersFirstname = 2
centersMiddleName = 3
centersStatus = 4
centersAppDate = 5
centersEmail = 6
centersAge = 7
centersGA = 8
centersMajor1 = 9
centersMajor2 = 10
centersMajor3 = 11
centersMinor1 = 12
centersMinor2 = 13
centersHonors = 14
centersInstGPA = 15
centersOvGPA = 16
centersInstHrs = 17
centersOvHrs = 18
centers8x = 1
centersProgram = 20
centersDegree = 21
centersNickname = 24
centersLocAddress = 26
centersLocPhone = 35

'modify all application dates to show month,day,year only
recordCounter = 2
Do While exportedData.Cells(recordCounter, exportedDataLastname).Value <> ""
  If exportedData.Cells(recordCounter, exportedDataAppDate).Value <> 0 Then
    exportedData.Cells(recordCounter, exportedDataAppDate).Value = Left(exportedData.Cells(recordCounter, exportedDataAppDate).Value, Len(exportedData.Cells(recordCounter, exportedDataAppDate).Value) - 4)
  End If
  
  recordCounter = recordCounter + 1
Loop

'duplicate person record check

recordCounter = 2
Do While exportedData.Cells(recordCounter, exportedDataLastname).Value <> ""
    For innerLoopCounter = 2 To 300
        If exportedData.Cells(innerLoopCounter, exportedData8x).Value = exportedData.Cells(recordCounter, exportedData8x).Value And InStr(exportedData.Cells(innerLoopCounter, exportedDataStatus).Value, "Duplicate") = 0 Then
            x = x + 1
        End If
        
        If x > 1 Then
            MsgBox (exportedData.Cells(recordCounter, exportedDataLastname).Value & vbNewLine & "Serious Error - duplicate records exist")
            If debugCode = False Then
                exportedData.UsedRange.ClearContents
                exportedData.Cells(1, 1).Value = "Copy and Paste output onto this sheet"
                End If
            Exit Sub 'serious error, macro stops all further actions
        End If
    Next innerLoopCounter
    
    x = 0
    recordCounter = recordCounter + 1
Loop

'make sure phone contains numeric values only, by checking each character one by one
'strip out alpha characters
Dim characterCounter, originalPhoneNumber, digitsOnlyPhoneNumber

recordCounter = 2
Do While exportedData.Cells(recordCounter, exportedDataLastname).Value <> ""
    originalPhoneNumber = exportedData.Cells(recordCounter, exportedDataLocalPhone).Value
    
    For characterCounter = 1 To Len(originalPhoneNumber)
        If IsNumeric(Mid(originalPhoneNumber, characterCounter, 1)) Then
            digitsOnlyPhoneNumber = digitsOnlyPhoneNumber & Mid(originalPhoneNumber, characterCounter, 1)
        End If
    Next characterCounter

    exportedData.Cells(recordCounter, exportedDataLocalPhone).Value = digitsOnlyPhoneNumber
    digitsOnlyPhoneNumber = ""
    recordCounter = recordCounter + 1
Loop

'begin data transfer
Dim exportedDataRowCounter, centersRowCounter
Dim nameChk As String
Dim firstSpace As Integer

exportedDataRowCounter = 2
Do While exportedData.Cells(exportedDataRowCounter, exportedDataLastname).Value <> ""
    For centersRowCounter = 11 To centersDB.UsedRange.Rows.Count
        'scenario one - we have a non-dup match! let's update our data! copy data and end the for loop
        If exportedData.Cells(exportedDataRowCounter, exportedData8x).Value = centersDB.Cells(centersRowCounter, centers8x).Value And InStr(exportedData.Cells(exportedDataRowCounter, exportedDataStatus).Value, "Duplicate") = 0 Then
            MsgBox ("we have a match!")
            centersDB.Cells(centersRowCounter, centersLastname).Value = exportedData.Cells(exportedDataRowCounter, exportedDataLastname).Value
            centersDB.Cells(centersRowCounter, centersFirstname).Value = exportedData.Cells(exportedDataRowCounter, exportedDataFirstname).Value
            centersDB.Cells(centersRowCounter, centersMiddleName).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMiddlename).Value
            
            'does the nickname exist??
            If exportedData.Cells(exportedDataRowCounter, exportedDataNickname).Value <> "" Then
                nameChk = exportedData.Cells(exportedDataRowCounter, exportedDataNickname).Value
                'find the position of the first space in the nickname
                If InStr(nameChk, " ") > 0 Then
                    firstSpace = InStr(nameChk, " ") - 1
                Else
                    firstSpace = Len(nameChk)
                End If
                nameChk = Left(nameChk, firstSpace)
                If exportedData.Cells(exportedDataRowCounter, exportedDataFirstname).Value <> nameChk Then
                    centersDB.Cells(centersRowCounter, centersNickname).Value = nameChk
                End If
            End If
   
            centersDB.Cells(centersRowCounter, centersAppDate).Value = exportedData.Cells(exportedDataRowCounter, exportedDataAppDate).Value
            centersDB.Cells(centersRowCounter, centersStatus).Value = exportedData.Cells(exportedDataRowCounter, exportedDataStatus).Value
            centersDB.Cells(centersRowCounter, centersAge).Value = exportedData.Cells(exportedDataRowCounter, exportedDataAge).Value
            centersDB.Cells(centersRowCounter, centersLocAddress).Value = exportedData.Cells(exportedDataRowCounter, exportedDataLocAddress).Value
            centersDB.Cells(centersRowCounter, centersLocPhone).Value = exportedData.Cells(exportedDataRowCounter, exportedDataLocalPhone).Value
            centersDB.Cells(centersRowCounter, centersEmail).Value = exportedData.Cells(exportedDataRowCounter, exportedDataEmail).Value
            centersDB.Cells(centersRowCounter, centersGA).Value = exportedData.Cells(exportedDataRowCounter, exportedDataGA).Value
            centersDB.Cells(centersRowCounter, centersMajor1).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMajor1).Value
            centersDB.Cells(centersRowCounter, centersMajor2).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMajor2).Value
            centersDB.Cells(centersRowCounter, centersMinor1).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMinor1).Value
            centersDB.Cells(centersRowCounter, centersMinor2).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMinor2).Value
            centersDB.Cells(centersRowCounter, centersInstGPA).Value = exportedData.Cells(exportedDataRowCounter, exportedDataInstGPA).Value
            centersDB.Cells(centersRowCounter, centersOvGPA).Value = exportedData.Cells(exportedDataRowCounter, exportedDataOvGPA).Value
            centersDB.Cells(centersRowCounter, centersInstHrs).Value = exportedData.Cells(exportedDataRowCounter, exportedDataInstHrs).Value
            centersDB.Cells(centersRowCounter, centersOvHrs).Value = exportedData.Cells(exportedDataRowCounter, exportedDataOvHrs).Value
            centersDB.Cells(centersRowCounter, centersHonors).Value = exportedData.Cells(exportedDataRowCounter, exportedDataHonors).Value
        Exit For
        
        'scenario two, we've hit the end of known centers applications. no other match found, migrate the rest of the unique record data exportedData
        ElseIf centersRowCounter = centersDB.UsedRange.Rows.Count And InStr(exportedData.Cells(exportedDataRowCounter, exportedDataStatus).Value, "Duplicate") = 0 Then
            centersDB.Rows(centersRowCounter).Insert Shift:=xlDown, _
            CopyOrigin:=xlFormatFromLeftOrAbove
            centersDB.Rows(centersRowCounter).Interior.ColorIndex = 0
            centersDB.Cells(centersRowCounter, centers8x).Value = exportedData.Cells(exportedDataRowCounter, exportedData8x).Value
            centersDB.Cells(centersRowCounter, centersLastname).Value = exportedData.Cells(exportedDataRowCounter, exportedDataLastname).Value
            centersDB.Cells(centersRowCounter, centersFirstname).Value = exportedData.Cells(exportedDataRowCounter, exportedDataFirstname).Value
            centersDB.Cells(centersRowCounter, centersMiddleName).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMiddlename).Value
            
            'nickname check
            If exportedData.Cells(exportedDataRowCounter, exportedDataNickname).Value <> "" Then
                nameChk = exportedData.Cells(exportedDataRowCounter, exportedDataNickname).Value
                firstSpace = InStr(nameChk, " ")
                If firstSpace > 0 Then
                    firstSpace = firstSpace - 1
                Else
                    firstSpace = Len(nameChk)
                End If
                
                nameChk = Left(nameChk, firstSpace)
                If exportedData.Cells(exportedDataRowCounter, exportedDataFirstname).Value <> nameChk Then
                    centersDB.Cells(centersRowCounter, centersNickname).Value = nameChk
                End If
            End If
  
            centersDB.Cells(centersRowCounter, centersAppDate).Value = exportedData.Cells(exportedDataRowCounter, exportedDataAppDate).Value
            centersDB.Cells(centersRowCounter, centersStatus).Value = exportedData.Cells(exportedDataRowCounter, exportedDataStatus).Value
            centersDB.Cells(centersRowCounter, centersAge).Value = exportedData.Cells(exportedDataRowCounter, exportedDataAge).Value
            centersDB.Cells(centersRowCounter, centersLocAddress).Value = exportedData.Cells(exportedDataRowCounter, exportedDataLocAddress).Value
            centersDB.Cells(centersRowCounter, centersLocPhone).Value = exportedData.Cells(exportedDataRowCounter, exportedDataLocalPhone).Value
            centersDB.Cells(centersRowCounter, centersEmail).Value = exportedData.Cells(exportedDataRowCounter, exportedDataEmail).Value
            centersDB.Cells(centersRowCounter, centersGA).Value = exportedData.Cells(exportedDataRowCounter, exportedDataGA).Value
            centersDB.Cells(centersRowCounter, centersMajor1).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMajor1).Value
            centersDB.Cells(centersRowCounter, centersMajor2).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMajor2).Value
            centersDB.Cells(centersRowCounter, centersMinor1).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMinor1).Value
            centersDB.Cells(centersRowCounter, centersMinor2).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMinor2).Value
            centersDB.Cells(centersRowCounter, centersInstGPA).Value = exportedData.Cells(exportedDataRowCounter, exportedDataInstGPA).Value
            centersDB.Cells(centersRowCounter, centersOvGPA).Value = exportedData.Cells(exportedDataRowCounter, exportedDataOvGPA).Value
            centersDB.Cells(centersRowCounter, centersInstHrs).Value = exportedData.Cells(exportedDataRowCounter, exportedDataInstHrs).Value
            centersDB.Cells(centersRowCounter, centersOvHrs).Value = exportedData.Cells(exportedDataRowCounter, exportedDataOvHrs).Value
            centersDB.Cells(centersRowCounter, centersHonors).Value = exportedData.Cells(exportedDataRowCounter, exportedDataHonors).Value
            centersRowCounter = centersRowCounter + 1
        End If
    Next centersRowCounter
    
    exportedDataRowCounter = exportedDataRowCounter + 1
Loop

'finishing moves, flawless victory
centersDB.Cells(5, 3).Value = Now
If debugCode = False Then
    exportedData.UsedRange.ClearContents
    exportedData.Cells(1, 1).Value = "Copy and Paste centers onto this sheet"
End If

MsgBox ("end")
'it's a good idea to disable the optimization and set screen updating to true
Application.ScreenUpdating = True
End Sub

'TODO - create function for nickname check
Function nicknameCheck(nickname)
    If nickname <> "" Then
                nameChk = exportedData.Cells(exportedDataRowCounter, exportedDataNickname).Value
                firstSpace = InStr(nameChk, " ")
                If firstSpace > 0 Then
                    firstSpace = firstSpace - 1
                Else
                    firstSpace = Len(nameChk)
                End If
                
                nameChk = Left(nameChk, firstSpace)
                If exportedData.Cells(exportedDataRowCounter, exportedDataFirstname).Value <> nameChk Then
                    centersDB.Cells(centersRowCounter, centersNickname).Value = nameChk
                End If
            End If

End Function
