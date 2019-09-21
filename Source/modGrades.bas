Attribute VB_Name = "modGrades"
Option Explicit

'loop variables
Global lp1, lp2, lp3 As Integer

'count variables
Global cnt1, cnt2 As Integer

'input arrays
Global gstrInputSubjectCode(1 To 2200) As String
Global gstrInputForename(1 To 2200) As String
Global gstrInputSurname(1 To 2200) As String
Global gstrInputCandidateNumber(1 To 2200) As String
Global gstrInputGrade(1 To 2200) As String
Global gstrInputEffort(1 To 2200) As String
Global gstrInputSubjectName(1 To 2200) As String

'output arrays
Global gstrOutputName(1 To 220) As String
Global gstrOutputSurname(1 To 220) As String
Global gstrOutputForename(1 To 220) As String
Global gstrOutputCandidateNumber(1 To 220) As String
Global gstrOutputSubject(1 To 10, 1 To 220) As String
Global gstrOutputGrade(1 To 10, 1 To 220) As String
Global gstrOutputEffort(1 To 10, 1 To 220) As String

'location variables
Global gstrSettingsLocation As String
Global gstrWordLocation As String
Global gstrOutputDocumentLocation As String
Global gstrInputDatabaseLocation As String
Global gstrOutputDatabaseLocation As String
Global gstrSubjectListLocation As String
Global gstrBasicDatabaseLocation As String

'pointer variables
Global gintCurrentStudent As Integer
Global gintSubjectModifier As Integer

'miscellaneous
Global gintCurrentTab As Integer
Global gintResponse As String
Global gstrWantedSubjectCode As String
Global gstrWantedStudent As String
Global gstrLargeSubjectList(1 To 102) As String
Global gstrSelectedSubject As String
Global mintErrorAction As Integer
Global gblnListClear As Boolean
Global gstrAppPath As String
Global gintNumberOfRecords As Integer
Global gintNumberOfFields As Integer

Public Function CaseSubjectName(val As String) As String

'decode subject abbreviations
Select Case val
    Case "En"
        CaseSubjectName = "English"
    Case "Fr"
        CaseSubjectName = "French"
    Case "Gg"
        CaseSubjectName = "Geography"
    Case "Ma"
        CaseSubjectName = "Maths"
    Case "Ar"
        CaseSubjectName = "Art"
    Case "Pe"
        CaseSubjectName = "Physical Education"
    Case "Sb"
        CaseSubjectName = "Double Science"
    Case "Sa"
        CaseSubjectName = "Single Science"
    Case "Hi"
        CaseSubjectName = "History"
    Case "Te"
        CaseSubjectName = "Technology"
    Case "Rs"
        CaseSubjectName = "Religious Studies"
    Case "Pt"
        CaseSubjectName = "Pottery"
    Case "Ya"
        CaseSubjectName = "Youth Award Scheme"
    Case "Me"
        CaseSubjectName = "Media Studies"
    Case "Sp"
        CaseSubjectName = "Spanish"
    Case "Sp"
        CaseSubjectName = "Japanese"
    Case "Ge"
        CaseSubjectName = "German"
    Case "Jp"
        CaseSubjectName = "Japanese"
    Case "Dr"
        CaseSubjectName = "Drama"
    Case Else
        CaseSubjectName = "Unknown (" & val & ")"
End Select

End Function

Public Sub CreateStudentList(val As String)

'create list of student details:
Open val For Output As #1
    For lp1 = 1 To 2200 Step 10
        Print #1, gstrInputCandidateNumber(lp1); ","; gstrInputForename(lp1); ","; gstrInputSurname(lp1)
    Next lp1
Close #1

End Sub

Public Sub LoadSettings()

Dim intSettingsFileLength

'create settings file path
gstrSettingsLocation = gstrAppPath & "grades.ini"

'check length
Open gstrSettingsLocation For Binary As #1
    intSettingsFileLength = LOF(1)
Close

'check settings location is correct
If Exists(gstrSettingsLocation) = False Or intSettingsFileLength < 5 Then

    gstrOutputDocumentLocation = gstrAppPath & "grades.doc"
    gstrInputDatabaseLocation = gstrAppPath & "input.csv"
    gstrOutputDatabaseLocation = gstrAppPath & "output.csv"
    gstrSubjectListLocation = gstrAppPath & "subjects.csv"
    gstrWordLocation = WordLocation()
    
    'create settings file
    Open gstrSettingsLocation For Output As #1
        Print #1, gstrOutputDocumentLocation
        Print #1, gstrInputDatabaseLocation
        Print #1, gstrOutputDatabaseLocation
        Print #1, gstrSubjectListLocation
        Print #1, gstrWordLocation
    Close #1
    
Else

    'import settings from file
    Open gstrSettingsLocation For Input As #1
        Input #1, gstrOutputDocumentLocation
        Input #1, gstrInputDatabaseLocation
        Input #1, gstrOutputDatabaseLocation
        Input #1, gstrSubjectListLocation
        Input #1, gstrWordLocation
    Close #1

End If
 
End Sub

Public Sub CheckSettings()

    'check word path is correct
    If Exists(gstrWordLocation) = False Then
        gstrWordLocation = WordLocation
        If Exists(gstrWordLocation) = False Then
            MsgBox "Microsoft Word could not be found.", vbCritical, "Year 10 Predicted Grades"
            End
        End If
    End If

    'check subject list location is correct
    If Exists(gstrSubjectListLocation) = False Then
        gstrSubjectListLocation = gstrAppPath & "subjects.csv"
        If Exists(gstrSubjectListLocation) = False Then
            MsgBox "Subject list could not be found.", vbCritical, "Year 10 Predicted Grades"
            End
        End If
    End If

    'check input database location is correct
    If Exists(gstrInputDatabaseLocation) = False Then
        gstrInputDatabaseLocation = gstrAppPath & "input.csv"
        If Exists(gstrInputDatabaseLocation) = False Then
            MsgBox "Input database could not be found.", vbCritical, "Year 10 Predicted Grades"
            End
        End If
    End If

    'check output document location is correct
    If Exists(gstrOutputDocumentLocation) = False Then
        gstrOutputDocumentLocation = gstrAppPath & "grades.doc"
        If Exists(gstrOutputDocumentLocation) = False Then
            MsgBox "Output document could not be found.", vbCritical, "Year 10 Predicted Grades"
            End
        End If
    End If
    
End Sub

Public Sub SaveSettings()

'save new values direct from boxes
Open gstrSettingsLocation For Output As #1
    Print #1, frmOptions.txtDocumentLocation.Text
    Print #1, frmOptions.txtInputDatabaseLocation.Text
    Print #1, frmOptions.txtOutputDatabaseLocation.Text
    Print #1, gstrSubjectListLocation
    Print #1, gstrWordLocation
Close #1

End Sub

Public Sub LoadData(val As String)

'display status
frmMain.staMain.SimpleText = "Importing data from " & val & "..."

'open file
Open val For Input As #1
    For cnt1 = 1 To 2200
        Input #1, gstrInputSubjectCode(cnt1), gstrInputSurname(cnt1), gstrInputForename(cnt1), gstrInputCandidateNumber(cnt1), gstrInputGrade(cnt1), gstrInputEffort(cnt1)
    Next cnt1
Close #1

'list all students
Call ListStudents

'move to first record and update fields
gintCurrentStudent = 1
gintSubjectModifier = 0
Call UpdateFields(gintCurrentStudent, gintSubjectModifier)

'display status
frmMain.staMain.SimpleText = "Ready"

End Sub

Public Sub ListStudents()

frmMain.lstStudent(0).Clear

'display list of all students
For lp1 = 1 To 2200 Step 10
    frmMain.lstStudent(0).AddItem gstrInputCandidateNumber(lp1) & " - " & gstrInputSurname(lp1) & ", " & gstrInputForename(lp1)
Next lp1

End Sub

Public Sub SaveData(val As String)

'display status
frmMain.staMain.SimpleText = "Saving " & val & "..."

'open specified file to export data
Open val For Output As #1
    For cnt1 = 1 To 2200
        Print #1, gstrInputSubjectCode(cnt1); ","; gstrInputSurname(cnt1); ","; gstrInputForename(cnt1); ","; gstrInputCandidateNumber(cnt1); ","; gstrInputGrade(cnt1); ","; gstrInputEffort(cnt1) 'write data from arrays to disk
    Next cnt1
Close #1

'update status
frmMain.staMain.SimpleText = "Ready"

End Sub

Public Sub UpdateFields(val1 As Integer, val2 As Integer)

'display student details
frmMain.lblStudent(1).Caption = gstrInputForename(val1) & " " & gstrInputSurname(val1)
frmMain.lblCandidateNumber(1).Caption = gstrInputCandidateNumber(val1 + val2)

'display subject details
frmMain.lblSubject(1).Caption = CaseSubjectName(Left(gstrInputSubjectCode(val1 + val2), 2))
frmMain.lblSubjectCode(1).Caption = gstrInputSubjectCode(val1 + val2)

'show current grade
Select Case gstrInputGrade(val1 + val2)
    Case "A* to C"
        frmMain.optGrade(0).Value = True
    Case "C to E"
        frmMain.optGrade(1).Value = True
    Case "E to U"
        frmMain.optGrade(2).Value = True
    Case Else
        For lp1 = 1 To 3
            frmMain.optGrade(lp1 - 1).Value = False
        Next lp1
End Select

'display current effort
Select Case gstrInputEffort(val1 + val2)
    Case "Excellent"
        frmMain.optEffort(0).Value = True
    Case "Good"
        frmMain.optEffort(1).Value = True
    Case "Average"
        frmMain.optEffort(2).Value = True
    Case "Satisfactory"
        frmMain.optEffort(3).Value = True
    Case "Poor"
        frmMain.optEffort(4).Value = True
    Case Else

End Select

End Sub

Public Sub WantedStudent(val As String)

'search input database for candidate number
For lp1 = 1 To 2200 Step 10
    Select Case gstrInputCandidateNumber(lp1)
        Case val
            gintCurrentStudent = lp1
    End Select
Next lp1

End Sub

Public Sub ListAllSubjects(val As String)

'open data file containing subject list
Open val For Input As #1
    For lp1 = 1 To 102
        Input #1, gstrLargeSubjectList(lp1)
        frmMain.lstSubject(1).AddItem gstrLargeSubjectList(lp1)
    Next lp1
Close #1

End Sub

Public Sub ExportData(val1 As String)

'show status
frmMain.staMain.SimpleText = "Exporting to " & val1 & "..."

'set count value
cnt2 = 1

'enter the main details
For lp3 = 1 To 2200 Step 10

gstrOutputCandidateNumber(cnt2) = gstrInputCandidateNumber(lp3)
gstrOutputForename(cnt2) = gstrInputForename(lp3)
gstrOutputSurname(cnt2) = gstrInputSurname(lp3)

For lp2 = 1 To 10
gstrInputSubjectName(lp3 + (lp2 - 1)) = CaseSubjectName(Left(gstrInputSubjectCode(lp3 + (lp2 - 1)), 2))
gstrOutputSubject(lp2, cnt2) = gstrInputSubjectName(lp3 + (lp2 - 1))
Next lp2

For lp2 = 1 To 10
gstrOutputGrade(lp2, cnt2) = gstrInputGrade(lp3 + (lp2 - 1))
Next lp2

For lp2 = 1 To 10
gstrOutputEffort(lp2, cnt2) = gstrInputEffort(lp3 + (lp2 - 1))
Next lp2

cnt2 = cnt2 + 1
Next lp3

'write to file
Open val1 For Output As #1
    
    'export column headings
    Print #1, "Candidate Number"; ","; "Name"; ","; "Subject 1"; ","; "Grade 1"; ","; "Effort 1"; ","; "Subject 2"; ","; "Grade 2"; ","; "Effort 2"; ","; "Subject 3"; ","; "Grade 3"; ","; "Effort 3"; ","; "Subject 4"; ","; "Grade 4"; ","; "Effort 4"; ","; "Subject 5"; ","; "Grade 5"; ","; "Effort 5"; ","; "Subject 6"; ","; "Grade 6"; ","; "Effort 6"; ","; "Subject 7"; ","; "Grade 7"; ","; "Effort 7"; ","; "Subject 8"; ","; "Grade 8"; ","; "Effort 8"; ","; "Subject 9"; ","; "Grade 9"; ","; "Effort 9"; ","; "Subject 10"; ","; "Grade 10"; ","; "Effort 10"
    
    'export data
    For lp3 = 1 To 220
        gstrOutputName(lp3) = gstrOutputForename(lp3) & " " & gstrOutputSurname(lp3)
        Print #1, gstrOutputCandidateNumber(lp3); ","; gstrOutputName(lp3); ","; gstrOutputSubject(1, lp3); ","; gstrOutputGrade(1, lp3); ","; gstrOutputEffort(1, lp3); ","; gstrOutputSubject(2, lp3); ","; gstrOutputGrade(2, lp3); ","; gstrOutputEffort(2, lp3); ","; gstrOutputSubject(3, lp3); ","; gstrOutputGrade(3, lp3); ","; gstrOutputEffort(3, lp3); ","; gstrOutputSubject(4, lp3); ","; gstrOutputGrade(4, lp3); ","; gstrOutputEffort(4, lp3); ","; gstrOutputSubject(5, lp3); ","; gstrOutputGrade(5, lp3); ","; gstrOutputEffort(5, lp3); ","; gstrOutputSubject(6, lp3); ","; gstrOutputGrade(6, lp3); ","; gstrOutputEffort(6, lp3); ","; gstrOutputSubject(7, lp3); ","; gstrOutputGrade(7, lp3); ","; gstrOutputEffort(7, lp3); ","; gstrOutputSubject(8, lp3); ","; gstrOutputGrade(8, lp3); ","; gstrOutputEffort(8, lp3); ","; gstrOutputSubject(9, lp3); ","; gstrOutputGrade(9, lp3); ","; gstrOutputEffort(9, lp3); ","; gstrOutputSubject(10, lp3); ","; gstrOutputGrade(10, lp3); ","; gstrOutputEffort(10, lp3)
    Next lp3

Close #1

'show status
frmMain.staMain.SimpleText = "Ready"

End Sub

Public Sub ClearRadios()

For lp1 = 0 To 2
    frmMain.optGrade(lp1).Value = False
Next lp1

For lp1 = 0 To 4
    frmMain.optEffort(lp1).Value = False
Next lp1

End Sub

Public Sub ViewRadios(val As Boolean)

Select Case val

    Case False

        For lp1 = 0 To 2
            frmMain.optGrade(lp1).Value = False
            frmMain.optGrade(lp1).Enabled = False
        Next lp1

        For lp1 = 0 To 4
            frmMain.optEffort(lp1).Enabled = False
            frmMain.optEffort(lp1).Value = False
        Next lp1
    
    Case True
    
        For lp1 = 0 To 2
            frmMain.optGrade(lp1).Enabled = True
        Next lp1

        For lp1 = 0 To 4
            frmMain.optEffort(lp1).Enabled = True
        Next lp1
    
End Select

End Sub

Public Sub ClearAll()

'clear data
frmMain.optGrade(0).Value = False
frmMain.optGrade(1).Value = False
frmMain.optGrade(2).Value = False
frmMain.optEffort(0).Value = False
frmMain.optEffort(1).Value = False
frmMain.optEffort(2).Value = False
frmMain.optEffort(3).Value = False
frmMain.optEffort(4).Value = False
frmMain.lblStudent(1).Caption = ""
frmMain.lblSubject(1).Caption = ""
frmMain.lblSubjectCode(1).Caption = ""
frmMain.lblCandidateNumber(1).Caption = ""

End Sub

Public Function WordLocation() As String

'find Microsoft Word path in registry
WordLocation = GetSettingString(HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Extensions", "doc")
WordLocation = Left(WordLocation, (Len(WordLocation) - 6))

End Function

Public Function Exists(filename As String) As Boolean

On Error GoTo ErrorHandler

'open filename and check length
Open filename For Binary As #1
    If LOF(1) <> 0 Then Close: Exists = True: Exit Function
Close

Kill filename
Exists = False

Exit Function

ErrorHandler:
    If Err.Number = 76 Then
        Exists = False
    End If
    
End Function

Public Sub DatabaseSize(val As String)

Dim varNonsense As Variant
   
'count number of records
Open val For Input As 1
    cnt1 = 0
    While Not EOF(1)
        Line Input #1, varNonsense
        cnt1 = cnt1 + 1
    Wend
Close #1
gintNumberOfRecords = cnt1

'count number of fields
Open val For Input As 1
    cnt1 = 0
    While Not EOF(1)
        Input #1, varNonsense
        cnt1 = cnt1 + 1
    Wend
Close #1
gintNumberOfFields = cnt1 / gintNumberOfRecords

End Sub

Public Sub Add2UnknownFields(val As String)

Dim varValue As Variant

'append extra fields and create input database
Open val For Input As 1
    Open gstrInputDatabaseLocation For Output As 2
        For lp1 = 1 To gintNumberOfRecords
            Line Input #1, varValue
            varValue = varValue & ",Unknown,Unknown"
            Print #2, varValue
        Next lp1
    Close #2
Close #1

End Sub

