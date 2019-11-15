Attribute VB_Name = "Module3"
'setup terra dotta export data columns for reference
Dim exportedDataFirstname, exportedDataLastname, exportedDataMiddlename, exportedData8x, exportedDataAge, exportedDataInstGPA, exportedDataOvGPA
Dim exportedDataInstHrs, exportedDataOvHrs, exportedDataStatus, exportedDataAppDate, exportedDataGA, exportedDataHonors
Dim exportedDataMajor1, exportedDataMajor2, exportedDataMajor3, exportedDataMinor1, exportedDataMinor2, exportedDataEmail
'setup the centers database columns
Dim centersFirstname, centersLastname, centersMiddleName, centers8x, centersAge, centersInstGPA, centersOvGPA, centersInstHrs, centersOvHrs
Dim centersStatus, centersAppDate, centersProgram, centersGA, centersHonors, centersMajor1, centersMajor2, centersMajor3
Dim centersMinor1, centersMinor2, centersEmail

Dim exportedDataRowCounter As Integer

Sub FISFall2020()
'for optimization, don't refresh screen while running the macro
Application.ScreenUpdating = False

'setup worksheets for reference
Dim centersDB As Worksheet, exportedData As Worksheet

Dim centersRowCounter As Integer

Set centersDB = Worksheets(1)
Set exportedData = Worksheets(2)

'define common variables
Dim recordCounter, innerLoopCounter, debugCode
'use the debug code to prevent data loss while testing, not to be used in production
debugCode = True


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
centers8x = 1
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
    For innerLoopCounter = recordCounter + 1 To 300
        If exportedData.Cells(innerLoopCounter, exportedData8x).Value = exportedData.Cells(recordCounter, exportedData8x).Value Then
            MsgBox ("Serious Error - duplicate records exist" & vbNewLine _
            & exportedData.Cells(recordCounter, exportedDataLastname) & " - Row: " & exportedData.Cells(recordCounter, exportedDataLastname).Row _
            & vbNewLine & exportedData.Cells(innerLoopCounter, exportedDataLastname) & " - Row: " & exportedData.Cells(innerLoopCounter, exportedDataLastname).Row)
            
            If debugCode = False Then
                exportedData.UsedRange.ClearContents
                exportedData.Cells(1, 1).Value = "Copy and Paste output onto this sheet"
                End If
            Exit Sub 'serious error, macro stops all further actions
        End If
    Next innerLoopCounter
    recordCounter = recordCounter + 1
Loop

'begin data transfer
Dim centersRowEnd As Integer, centersStartCounter As Integer

centersStartCounter = 11
recordCounter = centersStartCounter
exportedDataRowCounter = 2
centersRowEnd = findEndOfCentersTable(centersDB)

If centersRowEnd < centersStartCounter Then
centersRowEnd = centersStartCounter
End If
'MsgBox (centersRowEnd)

While exportedData.Cells(exportedDataRowCounter, exportedDataLastname).Value <> ""
    For recordCounter = centersStartCounter To centersRowEnd
        'MsgBox (exportedData.Cells(exportedDataRowCounter, exportedData8x).Value & " " & centersDB.Cells(recordCounter, centers8x))
        'scenario one - we have a non-dup match! let's update our data! copy data and end the for loop
        If exportedData.Cells(exportedDataRowCounter, exportedData8x).Value = centersDB.Cells(recordCounter, centers8x).Value Then
            Call TransferData(centersDB, exportedData, recordCounter)
            Exit For
            
        'scenario two, we've hit the end of section, add new row and add applicant to that row
        ElseIf recordCounter = centersRowEnd Then
            centersDB.Rows(recordCounter).EntireRow.Insert Shift:=xlDown
            Call TransferData(centersDB, exportedData, recordCounter)
            centersRowEnd = centersRowEnd + 1
            
        'scenario three, add new applicant to current row in section
        ElseIf centersDB.Cells(recordCounter, centers8x).Value = "" Then
            Call TransferData(centersDB, exportedData, recordCounter)
        End If
        
    Next recordCounter
Wend

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

'This subroutine copies applicant records line-by-line from the "Report" tab and pastes into the "3-Center Applications" tab. Data is pasted from row 11 onwards
Sub TransferData(ByVal centersDB As Worksheet, ByVal exportedData As Worksheet, ByVal centersRowCounter As Integer)
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
            
            exportedDataRowCounter = exportedDataRowCounter + 1
End Sub

'This function returns the row just before the "Under Review" row. I am making the assumption that we are only checking
'for records within the Pre Review > Complete section
Function findEndOfCentersTable(ByVal centersDB As Worksheet) As Integer
Dim FoundCell As Range

  Const WHAT_TO_FIND As String = "Under Review"

            Set FoundCell = centersDB.Range("L:L").Find(what:=WHAT_TO_FIND, MatchCase:=True)
            If Not FoundCell Is Nothing Then
                findEndOfCentersTable = FoundCell.Row
            End If
End Function



