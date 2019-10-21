Sub internationalCenters()

Application.ScreenUpdating = False

'setup worksheets for reference
Dim database, output As Worksheet
Set database = Worksheets("3-Center Applications")
Set output = Worksheets(1)

'setup columns for reference
Dim report8x, universityid As Integer
Dim reportLast, Last As Integer
Dim reportFirst, First As Integer
Dim reportProgram, Program As Integer
Dim reportAppDate, AppDate As Integer
Dim reportStatus, Status As Integer
Dim reportLocAddress, LocAddress As Integer
Dim reportLocPhone, LocPhone As Integer
Dim reportEmail, Email As Integer
Dim reportAge, Age As Integer
Dim reportGA, GA As Integer
Dim reporegree, Degree As Integer
Dim reportMajor1, Major1 As Integer
Dim reportMajor2, Major2 As Integer
Dim reportMajor3, Major3 As Integer
Dim reportMinor1, Minor1 As Integer
Dim reportMinor2, Minor2 As Integer
Dim reportInstGPA, InstGPA As Integer
Dim reportOvGPA, OvGPA As Integer
Dim reportInstHrs, InstHrs As Integer
Dim reportOvHrs, OvHrs As Integer
Dim reportHons, Hons As Integer
Dim reportCriminal, Criminal As Integer

'assign column number to each corresponding title
report8x = 5
reportLast = 2
reportFirst = 3
reportMiddle = 4
reportNickname = 28
reportProgram = 15
reportStatus = 13
reportAppDate = 14
reportLocAddress = 45
reportLocPhone = 44
reportEmail = 26
reportAge = 6
reportGA = 19
reporegree = 34
reportMajor1 = 21
reportMajor2 = 22
reportMajor3 = 23
reportMinor1 = 24
reportMinor2 = 25
reportInstGPA = 7
reportOvGPA = 8
reportInstHrs = 10
reportOvHrs = 11
reportHons = 20


Last = 1
First = 2
Middle = 3
Program = 20
Status = 4
AppDate = 5
Email = 6
LocAddress = 26
LocPhone = 35
Age = 7
GA = 8
Degree = 21
Major1 = 9
Major2 = 10
Major3 = 11
Minor1 = 12
Minor2 = 13
Hons = 14
InstGPA = 15
OvGPA = 16
InstHrs = 17
OvHrs = 18
universityid = 19
Nickname = 24

Dim k As Integer
Dim l As String
Dim n As String
Dim p As Integer
k = 2
Do While output.Cells(k, Last).Value <> ""
  If output.Cells(k, AppDate).Value <> 0 Then
    output.Cells(k, AppDate).Value = Left(output.Cells(k, AppDate).Value, Len(output.Cells(k, AppDate).Value) - 4)
  End If
  k = k + 1
Loop

Dim z As Integer
z = 2
Dim y As Integer
y = 2
Dim x As Integer
x = 0

Do While output.Cells(z, Last).Value <> ""
  For y = 2 To 300
    If output.Cells(y, 810).Value = output.Cells(z, 810).Value And InStr(output.Cells(y, Status).Value, "Duplicate") = 0 Then
      x = x + 1
    End If
    If x > 1 Then
      MsgBox (output.Cells(z, Last).Value & vbNewLine & "There are duplicate records in the data. Please remove duplicate applicants in TerraDotta before importing into the Excel database. Thank you.")
      output.UsedRange.ClearContents
      output.Cells(1, 1).Value = "Copy and Paste  Output onto this sheet"
      Exit Sub
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

Do While output.Cells(q, Last).Value <> ""
  If output.Cells(q, LocPhone).Value <> "" Then
    phoneChk = output.Cells(q, LocPhone).Value
    For r = 1 To Len(phoneChk)
      If IsNumeric(Mid(phoneChk, r, 1)) Then
        newPhone = newPhone & Mid(phoneChk, r, 1)
      End If
    Next r
    output.Cells(q, LocPhone).Value = newPhone
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

Do While output.Cells(i, Last).Value <> ""
  For j = 11 To database.UsedRange.Rows.Count
    If output.Cells(i, 810).Value = database.Cells(j, report810).Value And InStr(output.Cells(i, Status).Value, "Duplicate") = 0 Then
      database.Cells(j, reportLast).Value = output.Cells(i, Last).Value
      database.Cells(j, reportFirst).Value = output.Cells(i, First).Value
      database.Cells(j, reportMiddle).Value = output.Cells(i, Middle).Value
      If output.Cells(i, Nickname).Value <> "" Then
        nameChk = output.Cells(i, Nickname).Value
        firstSpace = InStr(nameChk, " ")
        If firstSpace > 0 Then
          firstSpace = firstSpace - 1
        Else
          firstSpace = Len(nameChk)
        End If
        nameChk = Left(nameChk, firstSpace)
        If output.Cells(i, First).Value <> nameChk Then
          database.Cells(j, reportNickname).Value = nameChk
        End If
      End If
   
      database.Cells(m, reportAppDate).Value = output.Cells(i, AppDate).Value
      database.Cells(m, reportStatus).Value = output.Cells(i, Status).Value
      database.Cells(m, reportAge).Value = output.Cells(i, Age).Value
      database.Cells(m, reportLocAddress).Value = output.Cells(i, LocAddress).Value
      database.Cells(m, reportLocPhone).Value = output.Cells(i, LocPhone).Value
      database.Cells(m, reportEmail).Value = output.Cells(i, Email).Value
      database.Cells(m, reportGA).Value = output.Cells(i, GA).Value
      database.Cells(m, reportMajor1).Value = output.Cells(i, Major1).Value
      database.Cells(m, reportMajor2).Value = output.Cells(i, Major2).Value
      database.Cells(m, reportMinor1).Value = output.Cells(i, Minor1).Value
      database.Cells(m, reportMinor2).Value = output.Cells(i, Minor2).Value
      database.Cells(m, reportInstGPA).Value = output.Cells(i, InstGPA).Value
      database.Cells(m, reportOvGPA).Value = output.Cells(i, OvGPA).Value
      database.Cells(m, reportInstHrs).Value = output.Cells(i, InstHrs).Value
      database.Cells(m, reportOvHrs).Value = output.Cells(i, OvHrs).Value
      database.Cells(m, reportHons).Value = output.Cells(i, Hons).Value
      Exit For
    ElseIf j = database.UsedRange.Rows.Count And InStr(output.Cells(i, Status).Value, "Duplicate") = 0 Then
      database.Rows(m).Insert Shift:=xlDown, _
      CopyOrigin:=xlFormatFromLeftOrAbove
      database.Rows(m).Interior.ColorIndex = 0
      database.Cells(m, report810).Value = output.Cells(i, 810).Value
      database.Cells(m, reportLast).Value = output.Cells(i, Last).Value
      database.Cells(m, reportFirst).Value = output.Cells(i, First).Value
      database.Cells(m, reportMiddle).Value = output.Cells(i, Middle).Value
      If output.Cells(i, Nickname).Value <> "" Then
        nameChk = output.Cells(i, Nickname).Value
        firstSpace = InStr(nameChk, " ")
        If firstSpace > 0 Then
          firstSpace = firstSpace - 1
        Else
          firstSpace = Len(nameChk)
        End If
        nameChk = Left(nameChk, firstSpace)
        If output.Cells(i, First).Value <> nameChk Then
          database.Cells(m, reportNickname).Value = nameChk
        End If
      End If
  
      database.Cells(m, reportAppDate).Value = output.Cells(i, AppDate).Value
      database.Cells(m, reportStatus).Value = output.Cells(i, Status).Value
      database.Cells(m, reportAge).Value = output.Cells(i, Age).Value
      database.Cells(m, reportLocAddress).Value = output.Cells(i, LocAddress).Value
      database.Cells(m, reportLocPhone).Value = output.Cells(i, LocPhone).Value
      database.Cells(m, reportEmail).Value = output.Cells(i, Email).Value
      database.Cells(m, reportGA).Value = output.Cells(i, GA).Value
      database.Cells(m, reportMajor1).Value = output.Cells(i, Major1).Value
      database.Cells(m, reportMajor2).Value = output.Cells(i, Major2).Value
      database.Cells(m, reportMinor1).Value = output.Cells(i, Minor1).Value
      database.Cells(m, reportMinor2).Value = output.Cells(i, Minor2).Value
      database.Cells(m, reportInstGPA).Value = output.Cells(i, InstGPA).Value
      database.Cells(m, reportOvGPA).Value = output.Cells(i, OvGPA).Value
      database.Cells(m, reportInstHrs).Value = output.Cells(i, InstHrs).Value
      database.Cells(m, reportOvHrs).Value = output.Cells(i, OvHrs).Value
      database.Cells(m, reportHons).Value = output.Cells(i, Hons).Value
      m = m + 1
    End If
  Next j
  i = i + 1
Loop

'finishing moves, flawless victory
database.Cells(5, 3).Value = Now
output.UsedRange.ClearContents
output.Cells(1, 1).Value = "Copy and Paste report onto this sheet"

Application.ScreenUpdating = True

End Sub
