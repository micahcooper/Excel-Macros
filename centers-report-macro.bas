Sub internationalCenters()

Application.ScreenUpdating = False
'use the debug code to prevent data loss while testing, not to be used in production
Dim debugCode As Boolean
debugCode = True

'setup worksheets for reference
Dim centersDB, exportedData As Worksheet
Set centersDB = Worksheets(2)
Set exportedData = Worksheets(1)

'setup terra dotta export data columns for reference
Dim exportedDataFirst, exportedDataLastname, exportedData8x, exportedDataAge, exportedDataInstGPA, exportedDataOvGPA, exportedDataInstHrs
Dim exportedDataOvHrs, exportedDataStatus, exportedDataAppDate, exportedDataProgram, exportedDataGA, exportedDataHonors
Dim exportedDataMajor1, exportedDataMajor2, exportedDataMajor3, exportedDataMinor1, exportedDataMinor2, exportedDataEmail
Dim exportedDataDegree, exportedDataLocalPhone, exportedDataLocAddress, exportedDataCriminal
'setup the centers database columns
Dim centersFirst, centersLast, centers8x, centersAge, centersInstGPA, centersOvGPA, centersInstHrs, centersOvHrs
Dim centersStatus, centersAppDate, centersProgram, centersGA, centersHonors, centersMajor1, centersMajor2, centersMajor3
Dim centersMinor1, centersMinor2, centersEmail, centersDegree, centersLocPhone, centersLocAddress, centersCriminal

'assign column number to each corresponding title
exportedDataFirst = "B"
exportedDataLastname = "C"
exportedDataMiddle = "D"
exportedData8x = "CX"
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

centersLast = 1
centersFirst = 2
centersMiddle = 3
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
centers8x = 19
centersProgram = 20
centersDegree = 21
centersNickname = 24
centersLocAddress = 26
centersLocPhone = 35

'modify dates to show month,day,year only
Dim k As Integer
Dim l As String
Dim n As String
Dim p As Integer
k = 2
Do While exportedData.Cells(k, exportedDataLastname).Value <> ""
  If exportedData.Cells(k, exportedDataAppDate).Value <> 0 Then
    exportedData.Cells(k, exportedDataAppDate).Value = Left(exportedData.Cells(k, exportedDataAppDate).Value, Len(exportedData.Cells(k, exportedDataAppDate).Value) - 4)
  End If
  k = k + 1
Loop

'duplicate person record check
Dim z, y, x
x = 0
z = 2
y = 2

Do While exportedData.Cells(z, exportedDataLastname).Value <> ""
  For y = 2 To 300
    If exportedData.Cells(y, exportedData8x).Value = exportedData.Cells(z, exportedData8x).Value And InStr(exportedData.Cells(y, exportedDataStatus).Value, "Duplicate") = 0 Then
      x = x + 1
    End If
    If x > 1 Then
      MsgBox (exportedData.Cells(z, exportedDataLastname).Value & vbNewLine & "Serious Error - duplicate records exist")
        If debugCode = False Then
            exportedData.UsedRange.ClearContents
            exportedData.Cells(1, 1).Value = "Copy and Paste output onto this sheet"
        End If
      Exit Sub 'serious error, macro stops all further actions
    End If
  Next y
  x = 0
  z = z + 1
Loop

'make sure phone contains numeric values only, by checking each character one by one
'strip out alpha characters
Dim recordCounter, characterCounter, originalPhoneNumber, digitsOnlyPhoneNumber

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
Dim s As Integer
Dim t As String
Dim i As Integer
Dim j As Integer
Dim m As Integer
Dim nameChk As String
Dim firstSpace As Integer
s = 2
i = 2
m = 8

Do While exportedData.Cells(i, exportedDataLastname).Value <> ""
  For j = 11 To centersDB.UsedRange.Rows.Count
    If exportedData.Cells(i, exportedData8x).Value = centersDB.Cells(j, centers8x).Value And InStr(exportedData.Cells(i, exportedDataStatus).Value, "Duplicate") = 0 Then
      centersDB.Cells(j, centersLast).Value = exportedData.Cells(i, exportedDataLastname).Value
      centersDB.Cells(j, centersFirst).Value = exportedData.Cells(i, exportedDataFirst).Value
      centersDB.Cells(j, centersMiddle).Value = exportedData.Cells(i, exportedDataMiddle).Value
      'does the nickname exist??
      If exportedData.Cells(i, exportedDataNickname).Value <> "" Then
        nameChk = exportedData.Cells(i, exportedDataNickname).Value
        firstSpace = InStr(nameChk, " ")
        If firstSpace > 0 Then
          firstSpace = firstSpace - 1
        Else
          firstSpace = Len(nameChk)
        End If
        nameChk = Left(nameChk, firstSpace)
        If exportedData.Cells(i, exportedDataFirst).Value <> nameChk Then
          centersDB.Cells(j, centersNickname).Value = nameChk
        End If
      End If
   
      centersDB.Cells(m, centersAppDate).Value = exportedData.Cells(i, exportedDataAppDate).Value
      centersDB.Cells(m, centersStatus).Value = exportedData.Cells(i, exportedDataStatus).Value
      centersDB.Cells(m, centersAge).Value = exportedData.Cells(i, exportedDataAge).Value
      centersDB.Cells(m, centersLocAddress).Value = exportedData.Cells(i, exportedDataLocAddress).Value
      centersDB.Cells(m, centersLocPhone).Value = exportedData.Cells(i, exportedDataLocalPhone).Value
      centersDB.Cells(m, centersEmail).Value = exportedData.Cells(i, exportedDataEmail).Value
      centersDB.Cells(m, centersGA).Value = exportedData.Cells(i, exportedDataGA).Value
      centersDB.Cells(m, centersMajor1).Value = exportedData.Cells(i, exportedDataMajor1).Value
      centersDB.Cells(m, centersMajor2).Value = exportedData.Cells(i, exportedDataMajor2).Value
      centersDB.Cells(m, centersMinor1).Value = exportedData.Cells(i, exportedDataMinor1).Value
      centersDB.Cells(m, centersMinor2).Value = exportedData.Cells(i, exportedDataMinor2).Value
      centersDB.Cells(m, centersInstGPA).Value = exportedData.Cells(i, exportedDataInstGPA).Value
      centersDB.Cells(m, centersOvGPA).Value = exportedData.Cells(i, exportedDataOvGPA).Value
      centersDB.Cells(m, centersInstHrs).Value = exportedData.Cells(i, exportedDataInstHrs).Value
      centersDB.Cells(m, centersOvHrs).Value = exportedData.Cells(i, exportedDataOvHrs).Value
      centersDB.Cells(m, centersHonors).Value = exportedData.Cells(i, exportedDataHonors).Value
      Exit For
    ElseIf j = centersDB.UsedRange.Rows.Count And InStr(exportedData.Cells(i, exportedDataStatus).Value, "Duplicate") = 0 Then
      centersDB.Rows(m).Insert Shift:=xlDown, _
      CopyOrigin:=xlFormatFromLeftOrAbove
      centersDB.Rows(m).Interior.ColorIndex = 0
      centersDB.Cells(m, centers810).Value = exportedData.Cells(i, exportedData8x).Value
      centersDB.Cells(m, centersLast).Value = exportedData.Cells(i, exportedDataLastname).Value
      centersDB.Cells(m, centersFirst).Value = exportedData.Cells(i, exportedDataFirst).Value
      centersDB.Cells(m, centersMiddle).Value = exportedData.Cells(i, exportedDataMiddle).Value
      'nickname check
      If exportedData.Cells(i, exportedDataNickname).Value <> "" Then
        nameChk = exportedData.Cells(i, exportedDataNickname).Value
        firstSpace = InStr(nameChk, " ")
        If firstSpace > 0 Then
          firstSpace = firstSpace - 1
        Else
          firstSpace = Len(nameChk)
        End If
        nameChk = Left(nameChk, firstSpace)
        If exportedData.Cells(i, exportedDataFirst).Value <> nameChk Then
          centersDB.Cells(m, centersNickname).Value = nameChk
        End If
      End If
  
      centersDB.Cells(m, centersAppDate).Value = exportedData.Cells(i, exportedDataAppDate).Value
      centersDB.Cells(m, centersStatus).Value = exportedData.Cells(i, exportedDataStatus).Value
      centersDB.Cells(m, centersAge).Value = exportedData.Cells(i, exportedDataAge).Value
      centersDB.Cells(m, centersLocAddress).Value = exportedData.Cells(i, exportedDataLocAddress).Value
      centersDB.Cells(m, centersLocPhone).Value = exportedData.Cells(i, exportedDataLocalPhone).Value
      centersDB.Cells(m, centersEmail).Value = exportedData.Cells(i, exportedDataEmail).Value
      centersDB.Cells(m, centersGA).Value = exportedData.Cells(i, exportedDataGA).Value
      centersDB.Cells(m, centersMajor1).Value = exportedData.Cells(i, exportedDataMajor1).Value
      centersDB.Cells(m, centersMajor2).Value = exportedData.Cells(i, exportedDataMajor2).Value
      centersDB.Cells(m, centersMinor1).Value = exportedData.Cells(i, exportedDataMinor1).Value
      centersDB.Cells(m, centersMinor2).Value = exportedData.Cells(i, exportedDataMinor2).Value
      centersDB.Cells(m, centersInstGPA).Value = exportedData.Cells(i, exportedDataInstGPA).Value
      centersDB.Cells(m, centersOvGPA).Value = exportedData.Cells(i, exportedDataOvGPA).Value
      centersDB.Cells(m, centersInstHrs).Value = exportedData.Cells(i, exportedDataInstHrs).Value
      centersDB.Cells(m, centersOvHrs).Value = exportedData.Cells(i, exportedDataOvHrs).Value
      centersDB.Cells(m, centersHonors).Value = exportedData.Cells(i, exportedDataHonors).Value
      m = m + 1
    End If
  Next j
  i = i + 1
Loop

'finishing moves, flawless victory
centersDB.Cells(5, 3).Value = Now
If debugCode = False Then
    exportedData.UsedRange.ClearContents
    exportedData.Cells(1, 1).Value = "Copy and Paste centers onto this sheet"
End If

Application.ScreenUpdating = True
End Sub
