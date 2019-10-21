Sub internationalCenters()

Application.ScreenUpdating = False
'use the debug code to prevent data loss while testing, not to be used in production
Dim debugCode As Boolean
debugCode = True

'setup worksheets for reference
Dim centersDB, exportedData As Worksheet
Set centersDB = Worksheets(2)
Set exportedData = Worksheets(1)

'setup columns for reference
Dim centersFirst, exportedDataFirst As String
Dim centersLast, exportedDataLast As String
      
Dim centers8x, exportedData8x As String
Dim centersAge, exportedDataAge As String
Dim centersInstGPA, exportedDataInstGPA As String
Dim centersOvGPA, exportedDataOvGPA As String
Dim centersInstHrs, exportedDataInstHrs As String
Dim centersOvHrs, exportedDataOvHrs As String
Dim centersStatus, exportedDataStatus As String
Dim centersAppDate, exportedDataAppDate As String
Dim centersProgram, exportedDataProgram As String
Dim centersGA, exportedDataGA As String
Dim centersHonors, exportedDataHonors As String
Dim centersMajor1, exportedDataMajor1 As String
Dim centersMajor2, exportedDataMajor2 As String
Dim centersMajor3, exportedDataMajor3 As String
Dim centersMinor1, exportedDataMinor1 As String
Dim centersMinor2, exportedDataMinor2 As String
Dim centersEmail, exportedDataEmail As String

Dim centersDegree, exportedDataDegree As String
Dim centersLocPhone, exportedDataLocPhone As String
Dim centersLocAddress, exportedDataLocAddress As String

Dim centersCriminal, exportedDataCriminal As String

'assign column number to each corresponding title
exportedDataFirst = "B"
exportedDataLast = "C"
exportedDataMiddle = "D"
exportedData8x = "CX"
exportedDataAge = 6
exportedDataInstGPA = 7
exportedDataOvGPA = 8
exportedDataInstHrs = 10
exportedDataOvHrs = 11
exportedDataStatus = 13
exportedDataAppDate = 14
exportedDataProgram = 15
exportedDataGA = 19
exportedDataHonors = 20
exportedDataMajor1 = 21
exportedDataMajor2 = 22
exportedDataMajor3 = 23
exportedDataMinor1 = 24
exportedDataMinor2 = 25
exportedDataEmail = 26
exportedDataNickname = 28
exportedDataDegree = 34
exportedDataLocPhone = 44
exportedDataLocAddress = 45

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
centersuniversityid = 19
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
Do While exportedData.Cells(k, exportedDataLast).Value <> ""
  If exportedData.Cells(k, exportedDataAppDate).Value <> 0 Then
    exportedData.Cells(k, exportedDataAppDate).Value = Left(exportedData.Cells(k, exportedDataAppDate).Value, Len(exportedData.Cells(k, exportedDataAppDate).Value) - 4)
  End If
  k = k + 1
Loop

'duplicate person record check
Dim z As Integer
z = 2
Dim y As Integer
y = 2
Dim x As Integer
x = 0

Do While exportedData.Cells(z, exportedDataLast).Value <> ""
  For y = 2 To 300
    If exportedData.Cells(y, exportedData8x).Value = exportedData.Cells(z, exportedData8x).Value And InStr(exportedData.Cells(y, exportedDataStatus).Value, "Duplicate") = 0 Then
      x = x + 1
    End If
    If x > 1 Then
      MsgBox (exportedData.Cells(z, exportedDataLast).Value & vbNewLine & "Serious Error - duplicate records exist")
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

'phone checks
Dim q As Integer
q = 2
Dim r As Integer
r = 1
Dim phoneChk As String
Dim newPhone As String

Do While exportedData.Cells(q, exportedDataLast).Value <> ""
  If exportedData.Cells(q, exportedDataLocPhone).Value <> "" Then
    phoneChk = exportedData.Cells(q, exportedDataLocPhone).Value
    For r = 1 To Len(phoneChk)
      If IsNumeric(Mid(phoneChk, r, 1)) Then
        newPhone = newPhone & Mid(phoneChk, r, 1)
      End If
    Next r
    exportedData.Cells(q, exportedDataLocPhone).Value = newPhone
    newPhone = ""
  End If
  q = q + 1
Loop

    'Begin data transfer
Dim s As Integer
s = 2
Dim t As String

Dim i As Integer
Dim j As Integer
Dim m As Integer
Dim nameChk As String
Dim firstSpace As Integer
i = 2
m = 8

Do While exportedData.Cells(i, exportedDataLast).Value <> ""
  For j = 11 To centersDB.UsedRange.Rows.Count
    If exportedData.Cells(i, exportedData8x).Value = centersDB.Cells(j, centers8x).Value And InStr(exportedData.Cells(i, exportedDataStatus).Value, "Duplicate") = 0 Then
      centersDB.Cells(j, centersLast).Value = exportedData.Cells(i, exportedDataLast).Value
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
      centersDB.Cells(m, centersLocPhone).Value = exportedData.Cells(i, exportedDataLocPhone).Value
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
      centersDB.Cells(m, centersLast).Value = exportedData.Cells(i, exportedDataLast).Value
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
      centersDB.Cells(m, centersLocPhone).Value = exportedData.Cells(i, exportedDataLocPhone).Value
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
