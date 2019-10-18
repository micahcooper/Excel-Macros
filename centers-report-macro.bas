Attribute VB_Name = "Module1"
Sub FISD()

Application.ScreenUpdating = False

Dim db810 As Integer, td810 As Integer
Dim dbLast As Integer, tdLast As Integer
Dim dbFirst As Integer, tdFrist As Integer
Dim dbProgram As Integer, tdProgram As Integer
Dim dbAppDate As Integer, tdAppDate As Integer
Dim dbStatus As Integer, tdStatus As Integer
Dim dbLocAddress As Integer, tdLocAddress As Integer
Dim dbLocPhone As Integer, tdLocPhone As Integer
Dim dbEmail As Integer, tdEmail As Integer
Dim dbAge As Integer, tdAge As Integer
Dim dbGA As Integer, tdGA As Integer
Dim dbDegree As Integer, tdDegree As Integer
Dim dbMajor1 As Integer, tdMajor1 As Integer
Dim dbMajor2 As Integer, tdMajor2 As Integer
Dim dbMajor3 As Integer, tdMajor3 As Integer
Dim dbMinor1 As Integer, tdMinor1 As Integer
Dim dbMinor2 As Integer, tdMinor2 As Integer
Dim dbInstGPA As Integer, tdInstGPA As Integer
Dim dbOvGPA As Integer, tdOvGPA As Integer
Dim dbInstHrs As Integer, tdInstHrs As Integer
Dim dbOvHrs As Integer, tdOvHrs As Integer
Dim dbHons As Integer, tdHons As Integer
Dim dbCriminal As Integer, tdCriminal As Integer

db810 = 6
dbLast = 2
dbFirst = 3
dbMiddle = 4
dbNickname = 5
dbProgram = 13
tdStatus = 14
dbAppDate = 15
dbLocAddress = 45
dbLocPhone = 44
dbEmail = 46
dbAge = 7
dbGA = 32
dbDegree = 34
dbMajor1 = 35
dbMajor2 = 36
dbMajor3 = 37
dbMinor1 = 38
dbMinor2 = 39
dbInstGPA = 8
dbOvGPA = 9
dbInstHrs = 11
dbOvHrs = 12
dbHons = 33


tdLast = 1
tdFirst = 2
tdMiddle = 3
tdProgram = 4
tdStatus = 5
tdAppDate = 6
tdEmail = 9
tdLocAddress = 26
tdLocPhone = 35
tdAge = 10
tdGA = 11
tdDegree = 12
tdMajor1 = 13
tdMajor2 = 14
tdMajor3 = 15
tdMinor1 = 16
tdMinor2 = 17
tdHons = 18
tdInstGPA = 19
tdOvGPA = 20
tdInstHrs = 21
tdOvHrs = 22
td810 = 23
tdNickname = 24



Dim database As Worksheet
Dim tdOutput As Worksheet
Set database = Worksheets("3-Center Applications")
Set tdOutput = Worksheets("Report")

Dim k As Integer
Dim l As String
Dim n As String
Dim p As Integer
k = 2



k = 2
Do While tdOutput.Cells(k, tdLast).Value <> ""
  If tdOutput.Cells(k, tdAppDate).Value <> 0 Then
    tdOutput.Cells(k, tdAppDate).Value = Left(tdOutput.Cells(k, tdAppDate).Value, Len(tdOutput.Cells(k, tdAppDate).Value) - 4)
  End If
  k = k + 1
Loop

Dim z As Integer
z = 2
Dim y As Integer
y = 2
Dim x As Integer
x = 0

Do While tdOutput.Cells(z, tdLast).Value <> ""
  For y = 2 To 300
    If tdOutput.Cells(y, td810).Value = tdOutput.Cells(z, td810).Value And InStr(tdOutput.Cells(y, tdStatus).Value, "Duplicate") = 0 Then
      x = x + 1
    End If
    If x > 1 Then
      MsgBox (tdOutput.Cells(z, tdLast).Value & vbNewLine & "There are duplicate records in the data. Please remove duplicate applicants in TerraDotta before importing into the Excel database. Thank you.")
      tdOutput.UsedRange.ClearContents
      tdOutput.Cells(1, 1).Value = "Copy and Paste TD Output onto this sheet"
      Exit Sub
    End If
  Next y
  x = 0
  z = z + 1
Loop

Dim q As Integer
q = 2
Dim r As Integer
r = 1
Dim phoneChk As String
Dim newPhone As String

Do While tdOutput.Cells(q, tdLast).Value <> ""
  If tdOutput.Cells(q, tdLocPhone).Value <> "" Then
    phoneChk = tdOutput.Cells(q, tdLocPhone).Value
    For r = 1 To Len(phoneChk)
      If IsNumeric(Mid(phoneChk, r, 1)) Then
        newPhone = newPhone & Mid(phoneChk, r, 1)
      End If
    Next r
    tdOutput.Cells(q, tdLocPhone).Value = newPhone
    newPhone = ""
  End If
  q = q + 1
Loop

Dim s As Integer
s = 2
Dim t As String



Dim i As Integer
Dim j As Integer
Dim m As Integer
Dim nameChk As String
Dim firstSpace As Integer
i = 2
m = 12

Do While tdOutput.Cells(i, tdLast).Value <> ""
  For j = 11 To database.UsedRange.Rows.Count
    If tdOutput.Cells(i, td810).Value = database.Cells(j, db810).Value And InStr(tdOutput.Cells(i, tdStatus).Value, "Duplicate") = 0 Then
      database.Cells(j, dbLast).Value = tdOutput.Cells(i, tdLast).Value
      database.Cells(j, dbFirst).Value = tdOutput.Cells(i, tdFirst).Value
      database.Cells(j, dbMiddle).Value = tdOutput.Cells(i, tdMiddle).Value
      If tdOutput.Cells(i, tdNickname).Value <> "" Then
        nameChk = tdOutput.Cells(i, tdNickname).Value
        firstSpace = InStr(nameChk, " ")
        If firstSpace > 0 Then
          firstSpace = firstSpace - 1
        Else
          firstSpace = Len(nameChk)
        End If
        nameChk = Left(nameChk, firstSpace)
        If tdOutput.Cells(i, tdFirst).Value <> nameChk Then
          database.Cells(j, dbNickname).Value = nameChk
        End If
      End If
   
      database.Cells(j, dbAppDate).Value = tdOutput.Cells(i, tdAppDate).Value
      database.Cells(j, dbLocAddress).Value = tdOutput.Cells(i, tdLocAddress).Value
      database.Cells(j, dbLocPhone).Value = tdOutput.Cells(i, tdLocPhone).Value
      database.Cells(j, dbEmail).Value = tdOutput.Cells(i, tdEmail).Value
      database.Cells(j, dbGA).Value = tdOutput.Cells(i, tdGA).Value
      database.Cells(j, dbMajor1).Value = tdOutput.Cells(i, tdMajor1).Value
      database.Cells(j, dbMajor2).Value = tdOutput.Cells(i, tdMajor2).Value
      database.Cells(j, dbMinor1).Value = tdOutput.Cells(i, tdMinor1).Value
      database.Cells(j, dbMinor2).Value = tdOutput.Cells(i, tdMinor2).Value
      database.Cells(j, dbInstGPA).Value = tdOutput.Cells(i, tdInstGPA).Value
      database.Cells(j, dbOvGPA).Value = tdOutput.Cells(i, tdOvGPA).Value
      database.Cells(j, dbInstHrs).Value = tdOutput.Cells(i, tdInstHrs).Value
      database.Cells(j, dbOvHrs).Value = tdOutput.Cells(i, tdOvHrs).Value
      database.Cells(j, dbHons).Value = tdOutput.Cells(i, tdHons).Value
      Exit For
    ElseIf j = database.UsedRange.Rows.Count And InStr(tdOutput.Cells(i, tdStatus).Value, "Duplicate") = 0 Then
      database.Rows(m).Insert Shift:=xlDown, _
      CopyOrigin:=xlFormatFromLeftOrAbove
      database.Rows(m).Interior.ColorIndex = 0
      database.Cells(m, db810).Value = tdOutput.Cells(i, td810).Value
      database.Cells(m, dbLast).Value = tdOutput.Cells(i, tdLast).Value
      database.Cells(m, dbFirst).Value = tdOutput.Cells(i, tdFirst).Value
      database.Cells(m, dbMiddle).Value = tdOutput.Cells(i, tdMiddle).Value
      If tdOutput.Cells(i, tdNickname).Value <> "" Then
        nameChk = tdOutput.Cells(i, tdNickname).Value
        firstSpace = InStr(nameChk, " ")
        If firstSpace > 0 Then
          firstSpace = firstSpace - 1
        Else
          firstSpace = Len(nameChk)
        End If
        nameChk = Left(nameChk, firstSpace)
        If tdOutput.Cells(i, tdFirst).Value <> nameChk Then
          database.Cells(m, dbNickname).Value = nameChk
        End If
      End If
 
  
      database.Cells(m, dbAppDate).Value = tdOutput.Cells(i, tdAppDate).Value
      database.Cells(m, dbLocAddress).Value = tdOutput.Cells(i, tdLocAddress).Value
      database.Cells(m, dbLocPhone).Value = tdOutput.Cells(i, tdLocPhone).Value
      database.Cells(m, dbEmail).Value = tdOutput.Cells(i, tdEmail).Value
      database.Cells(m, dbGA).Value = tdOutput.Cells(i, tdGA).Value
      database.Cells(m, dbMajor1).Value = tdOutput.Cells(i, tdMajor1).Value
      database.Cells(m, dbMajor2).Value = tdOutput.Cells(i, tdMajor2).Value
      database.Cells(m, dbMinor1).Value = tdOutput.Cells(i, tdMinor1).Value
      database.Cells(m, dbMinor2).Value = tdOutput.Cells(i, tdMinor2).Value
      database.Cells(m, dbInstGPA).Value = tdOutput.Cells(i, tdInstGPA).Value
      database.Cells(m, dbOvGPA).Value = tdOutput.Cells(i, tdOvGPA).Value
      database.Cells(m, dbInstHrs).Value = tdOutput.Cells(i, tdInstHrs).Value
      database.Cells(m, dbOvHrs).Value = tdOutput.Cells(i, tdOvHrs).Value
      database.Cells(m, dbHons).Value = tdOutput.Cells(i, tdHons).Value
      m = m + 1
    End If
  Next j
  i = i + 1
Loop

database.Cells(5, 3).Value = Now
tdOutput.UsedRange.ClearContents
tdOutput.Cells(1, 1).Value = "Copy and Paste report onto this sheet"

Application.ScreenUpdating = True

End Sub
