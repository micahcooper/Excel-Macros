Sub internationalCenters()

Application.ScreenUpdating = False

'setup worksheets for reference
Dim centersDB, exportedData As Worksheet
Set centersDB = Worksheets("International Centers")
Set exportedData = Worksheets(1)

'setup columns for reference
Dim centers8x, exportedData8x As Integer
Dim centersLast, exportedDataLast As Integer
Dim centersFirst, exportedDataFirst As Integer
Dim centersProgram, exportedDataProgram As Integer
Dim centersAppDate, exportedDataAppDate As Integer
Dim centersStatus, exportedDataStatus As Integer
Dim centersLocAddress, exportedDataLocAddress As Integer
Dim centersLocPhone, exportedDataLocPhone As Integer
Dim centersEmail, exportedDataEmail As Integer
Dim centersAge, exportedDataAge As Integer
Dim centersGA, exportedDataGA As Integer
Dim centersDegree, exportedDataDegree As Integer
Dim centersMajor1, exportedDataMajor1 As Integer
Dim centersMajor2, exportedDataMajor2 As Integer
Dim centersMajor3, exportedDataMajor3 As Integer
Dim centersMinor1, exportedDataMinor1 As Integer
Dim centersMinor2, exportedDataMinor2 As Integer
Dim centersInstGPA, exportedDataInstGPA As Integer
Dim centersOvGPA, exportedDataOvGPA As Integer
Dim centersInstHrs, exportedDataInstHrs As Integer
Dim centersOvHrs, exportedDataOvHrs As Integer
Dim centersHonors, exportedDataHonors As Integer
Dim centersCriminal, exportedDataCriminal As Integer

'assign column number to each corresponding title
exportedData8x = 5
exportedDataLast = 2
exportedDataFirst = 3
exportedDataMiddle = 4
exportedDataNickname = 28
exportedDataProgram = 15
exportedDataStatus = 13
exportedDataAppDate = 14
exportedDataLocAddress = 45
exportedDataLocPhone = 44
exportedDataEmail = 26
exportedDataAge = 6
exportedDataGA = 19
exportedDataDegree = 34
exportedDataMajor1 = 21
exportedDataMajor2 = 22
exportedDataMajor3 = 23
exportedDataMinor1 = 24
exportedDataMinor2 = 25
exportedDataInstGPA = 7
exportedDataOvGPA = 8
exportedDataInstHrs = 10
exportedDataOvHrs = 11
exportedDataHonors = 20


centersLast = 1
centersFirst = 2
centersMiddle = 3
centersProgram = 20
centersStatus = 4
centersAppDate = 5
centersEmail = 6
centersLocAddress = 26
centersLocPhone = 35
centersAge = 7
centersGA = 8
centersDegree = 21
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
centersNickname = 24

'modify dates to show month,day,year only
Dim k As Integer
Dim l As String
Dim n As String
Dim p As Integer
k = 2
Do While exportedData.Cells(k, centersLast).Value <> ""
  If exportedData.Cells(k, centersAppDate).Value <> 0 Then
    exportedData.Cells(k, centersAppDate).Value = Left(exportedData.Cells(k, centersAppDate).Value, Len(exportedData.Cells(k, centersAppDate).Value) - 4)
  End If
  k = k + 1
Loop

'
Dim z As Integer
z = 2
Dim y As Integer
y = 2
Dim x As Integer
x = 0

Do While exportedData.Cells(z, centersLast).Value <> ""
  For y = 2 To 300
    If exportedData.Cells(y, centers8x).Value = exportedData.Cells(z, centers8x).Value And InStr(exportedData.Cells(y, centersStatus).Value, "Duplicate") = 0 Then
      x = x + 1
    End If
    If x > 1 Then
      MsgBox (exportedData.Cells(z, centersLast).Value & vbNewLine & "Serious Error - duplicate records exist")
      exportedData.UsedRange.ClearContents
      exportedData.Cells(1, 1).Value = "Copy and Paste output onto this sheet"
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

Do While exportedData.Cells(q, centersLast).Value <> ""
  If exportedData.Cells(q, centersLocPhone).Value <> "" Then
    phoneChk = exportedData.Cells(q, centersLocPhone).Value
    For r = 1 To Len(phoneChk)
      If IsNumeric(Mid(phoneChk, r, 1)) Then
        newPhone = newPhone & Mid(phoneChk, r, 1)
      End If
    Next r
    exportedData.Cells(q, centersLocPhone).Value = newPhone
    newPhone = ""
  End If
  q = q + 1
Loop

'check for duplicate student records
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

Do While exportedData.Cells(i, centersLast).Value <> ""
  For j = 11 To centersDB.UsedRange.Rows.Count
    If exportedData.Cells(i, centers8x).Value = centersDB.Cells(j, centers8x).Value And InStr(centers.Cells(i, exportedDataStatus).Value, "Duplicate") = 0 Then
      centersDB.Cells(j, centersLast).Value = centers.Cells(i, exportedDataLast).Value
      centersDB.Cells(j, centersFirst).Value = centers.Cells(i, exportedDataFirst).Value
      centersDB.Cells(j, centersMiddle).Value = centers.Cells(i, exportedDataMiddle).Value
      'does the nickname exist??
      If centers.Cells(i, exportedDataNickname).Value <> "" Then
        nameChk = centers.Cells(i, exportedDataNickname).Value
        firstSpace = InStr(nameChk, " ")
        If firstSpace > 0 Then
          firstSpace = firstSpace - 1
        Else
          firstSpace = Len(nameChk)
        End If
        nameChk = Left(nameChk, firstSpace)
        If centers.Cells(i, exportedDataFirst).Value <> nameChk Then
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
exportedData.UsedRange.ClearContents
exportedData.Cells(1, 1).Value = "Copy and Paste centers onto this sheet"

Application.ScreenUpdating = True
End Sub
