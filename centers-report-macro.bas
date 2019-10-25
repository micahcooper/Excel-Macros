Attribute VB_Name = "Module3"
Sub FISFall2020()
'for optimization, don't refresh screen while running the macro
Application.ScreenUpdating = False

'define common variables
Dim recordCounter, innerLoopCounter, debugCode
'use the debug code to prevent data loss while testing, not to be used in production
debugCode = True

'setup worksheets for reference
Dim centersDB, exportedData As Worksheet
Set centersDB = Worksheets(1)
Set exportedData = Worksheets(2)

'setup terra dotta export data columns for reference
Dim exportedDataFirstname, exportedDataLastname, exportedDataMiddlename, exportedData8x, exportedDataAge, exportedDataInstGPA, exportedDataOvGPA
Dim exportedDataInstHrs, exportedDataOvHrs, exportedDataStatus, exportedDataAppDate, exportedDataGA, exportedDataHonors
Dim exportedDataMajor1, exportedDataMajor2, exportedDataMajor3, exportedDataMinor1, exportedDataMinor2, exportedDataEmail
'setup the centers database columns
Dim centersFirstname, centersLastname, centersMiddleName, centers8x, centersAge, centersInstGPA, centersOvGPA, centersInstHrs, centersOvHrs
Dim centersStatus, centersAppDate, centersProgram, centersGA, centersHonors, centersMajor1, centersMajor2, centersMajor3
Dim centersMinor1, centersMinor2, centersEmail

'assign column letter to each exported data field title
exportedDataLastname = "A"
exportedDataFirstname = "B"
exportedDataMiddlename = "C"
exportedDataStatus = "D"
exportedDataAppDate = "E"
exportedDataEmail = "F"
exportedDataAge = "G"
exportedDataGA = "H"
exportedDataMajor1 = "I"
exportedDataMajor2 = "J"
exportedDataMajor3 = "K"
exportedDataMinor1 = "L"
exportedDataMinor2 = "M"
exportedDataHonors = "N"
exportedDataInstGPA = "O"
exportedDataOvGPA = "P"
exportedDataInstHrs = "Q"
exportedDataOvHrs = "R"
exportedData8x = "S"

'assign numerical position for the centers data
centersLastname = 2
centersFirstname = 3
centersMiddleName = 4
centers8x = 5
centersAge = 6
centersInstGPA = 7
centersOvGPA = 8
centersInstHrs = 10
centersOvHrs = 11
centersAppDate = 14
centersGA = 19
centersHonors = 20
centersMajor1 = 21
centersMajor2 = 22
centersMajor3 = 23
centersMinor1 = 24
centersMinor2 = 25
centersEmail = 26
centersStatus = 27

centersDB.Cells(11, 1).Value = exportedData.Cells(10, 1).Value

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


'begin data transfer
Dim exportedDataRowCounter, centersRowCounter
Dim nameChk As String
Dim centersRowEnd As Integer
Dim FoundCell As Range

  Const WHAT_TO_FIND As String = "Under Review"

            Set FoundCell = centersDB.Range("L:L").Find(What:=WHAT_TO_FIND, MatchCase:=True)
            If Not FoundCell Is Nothing Then
                MsgBox (WHAT_TO_FIND & " found in row: " & FoundCell.Row)
                centersRowEnd = FoundCell.Row - 1
            Else
                MsgBox (WHAT_TO_FIND & " not found")
            End If



exportedDataRowCounter = 2
Do While exportedData.Cells(exportedDataRowCounter, exportedDataLastname).Value <> ""
    For centersRowCounter = 11 To 50
        'scenario one - we have a non-dup match! let's update our data! copy data and end the for loop
        If exportedData.Cells(exportedDataRowCounter, exportedData8x).Value = centersDB.Cells(centersRowCounter, centers8x).Value And InStr(exportedData.Cells(exportedDataRowCounter, exportedDataStatus).Value, "Duplicate") = 0 Then
            
           ' MsgBox ("we have a match in the Centers Row Counter! " & centersRowCounter & "Used Range Row Count: " & centersDB.UsedRange.Rows.Count)
            
            centersDB.Cells(centersRowCounter, centersLastname).Value = exportedData.Cells(exportedDataRowCounter, exportedDataLastname).Value
            centersDB.Cells(centersRowCounter, centersFirstname).Value = exportedData.Cells(exportedDataRowCounter, exportedDataFirstname).Value
            centersDB.Cells(centersRowCounter, centersMiddleName).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMiddlename).Value
            
            centersDB.Cells(centersRowCounter, centersAppDate).Value = exportedData.Cells(exportedDataRowCounter, exportedDataAppDate).Value
            centersDB.Cells(centersRowCounter, centersStatus).Value = exportedData.Cells(exportedDataRowCounter, exportedDataStatus).Value
            centersDB.Cells(centersRowCounter, centersAge).Value = exportedData.Cells(exportedDataRowCounter, exportedDataAge).Value
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
        
        'scenario two, add new applicant to table
         ElseIf centersDB.Cells(centersRowCounter, centers8x).Value = "" And InStr(exportedData.Cells(exportedDataRowCounter, exportedDataStatus).Value, "Duplicate") = 0 Then
            MsgBox ("New applicant found with name: " & exportedData.Cells(exportedDataRowCounter, exportedDataLastname).Value & " to be inputted on centersDB row: " & centersRowCounter)
            
            centersDB.Cells(centersRowCounter, centers8x).Value = exportedData.Cells(exportedDataRowCounter, exportedData8x).Value
            centersDB.Cells(centersRowCounter, centersLastname).Value = exportedData.Cells(exportedDataRowCounter, exportedDataLastname).Value
            centersDB.Cells(centersRowCounter, centersFirstname).Value = exportedData.Cells(exportedDataRowCounter, exportedDataFirstname).Value
            centersDB.Cells(centersRowCounter, centersMiddleName).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMiddlename).Value
            centersDB.Cells(centersRowCounter, centersAppDate).Value = exportedData.Cells(exportedDataRowCounter, exportedDataAppDate).Value
            centersDB.Cells(centersRowCounter, centersStatus).Value = exportedData.Cells(exportedDataRowCounter, exportedDataStatus).Value
            centersDB.Cells(centersRowCounter, centersAge).Value = exportedData.Cells(exportedDataRowCounter, exportedDataAge).Value
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
            
        
        'scenario three, we've hit the ebd of table, add new row and add applicant
        'TODO: centersRowCounter = centersDB.UsedRange.Rows.Count is causing data to be entered in lines ~230. Must change UsedRange to something more granular.
        'This code block re-updates data for applicants
        
        
        ElseIf centersRowCounter = centersRowEnd And InStr(exportedData.Cells(exportedDataRowCounter, exportedDataStatus).Value, "Duplicate") = 0 Then
            MsgBox ("Reached end with new applicant found with name: " & exportedData.Cells(exportedDataRowCounter, exportedDataLastname).Value & " to be inputted on centersDB row: " & centersRowCounter)
            
            centersDB.Rows(centersRowCounter).EntireRow.Insert Shift:=xlDown ', '_
            'CopyOrigin:=xlFormatFromLeftOrAbove
            centersDB.Rows(centersRowCounter).Interior.ColorIndex = 0
            centersDB.Cells(centersRowCounter, centers8x).Value = exportedData.Cells(exportedDataRowCounter, exportedData8x).Value
            centersDB.Cells(centersRowCounter, centersLastname).Value = exportedData.Cells(exportedDataRowCounter, exportedDataLastname).Value
            centersDB.Cells(centersRowCounter, centersFirstname).Value = exportedData.Cells(exportedDataRowCounter, exportedDataFirstname).Value
            centersDB.Cells(centersRowCounter, centersMiddleName).Value = exportedData.Cells(exportedDataRowCounter, exportedDataMiddlename).Value
            centersDB.Cells(centersRowCounter, centersAppDate).Value = exportedData.Cells(exportedDataRowCounter, exportedDataAppDate).Value
            centersDB.Cells(centersRowCounter, centersStatus).Value = exportedData.Cells(exportedDataRowCounter, exportedDataStatus).Value
            centersDB.Cells(centersRowCounter, centersAge).Value = exportedData.Cells(exportedDataRowCounter, exportedDataAge).Value
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
            
            'centersRowEnd = centersRowEnd + 1
        
        End If
        'centersRowCounter = 11
        exportedDataRowCounter = exportedDataRowCounter + 1
    Next centersRowCounter
    
    
Loop

'finishing moves, flawless victory
centersDB.Cells(5, 3).Value = Now
If debugCode = False Then
    exportedData.UsedRange.ClearContents
    exportedData.Cells(1, 1).Value = "Copy and Paste centers onto this sheet"
End If

'MsgBox ("end")
'it's a good idea to disable the optimization and set screen updating to true
Application.ScreenUpdating = True
End Sub
