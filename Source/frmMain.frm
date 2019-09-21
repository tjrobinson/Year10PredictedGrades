VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year 10 Predicted Grades - Edit Data"
   ClientHeight    =   4935
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   7275
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMenu 
      Caption         =   "&Main Menu"
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      ToolTipText     =   "Return to Main Menu"
      Top             =   180
      Width           =   1455
   End
   Begin VB.Frame fraEffort 
      Caption         =   "&Effort"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   5640
      TabIndex        =   7
      Top             =   2580
      Width           =   1455
      Begin VB.OptionButton optEffort 
         Caption         =   "Poor"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Current effort"
         Top             =   1500
         Width           =   1215
      End
      Begin VB.OptionButton optEffort 
         Caption         =   "Satisfactory"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Current effort"
         Top             =   1200
         Width           =   1275
      End
      Begin VB.OptionButton optEffort 
         Caption         =   "Average"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Current effort"
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton optEffort 
         Caption         =   "Good"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Current effort"
         Top             =   600
         Width           =   1275
      End
      Begin VB.OptionButton optEffort 
         Caption         =   "Excellent"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Current effort"
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame fraGrade 
      Caption         =   "&Grade"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5640
      TabIndex        =   3
      Top             =   1020
      Width           =   1455
      Begin VB.OptionButton optGrade 
         Caption         =   "E to U"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Anticipated grade"
         Top             =   900
         Width           =   1155
      End
      Begin VB.OptionButton optGrade 
         Caption         =   "C to E"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Anticipated grade"
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optGrade 
         Caption         =   "A* to C"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Anticipated grade"
         Top             =   300
         Width           =   1155
      End
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   3675
      Left            =   180
      TabIndex        =   1
      Top             =   780
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6482
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   564
      ShowFocusRect   =   0   'False
      MouseIcon       =   "frmMain.frx":0442
      TabCaption(0)   =   "Input by stu&dent"
      TabPicture(0)   =   "frmMain.frx":045E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblGroup"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblStudent(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblStudent(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lstSubject(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lstStudent(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtGroup"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Input by su&bject"
      TabPicture(1)   =   "frmMain.frx":047A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstStudent(1)"
      Tab(1).Control(1)=   "lstSubject(1)"
      Tab(1).Control(2)=   "lblStudent(4)"
      Tab(1).Control(3)=   "lblSubject(2)"
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtGroup 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   4560
         TabIndex        =   26
         Top             =   3060
         Width           =   495
      End
      Begin VB.ListBox lstStudent 
         BackColor       =   &H00C0FFFF&
         Height          =   2400
         Index           =   0
         ItemData        =   "frmMain.frx":0496
         Left            =   240
         List            =   "frmMain.frx":049D
         TabIndex        =   23
         Top             =   960
         Width           =   2595
      End
      Begin VB.ListBox lstStudent 
         BackColor       =   &H00C0FFFF&
         Height          =   2400
         Index           =   1
         Left            =   -72840
         TabIndex        =   22
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox lstSubject 
         BackColor       =   &H00C0FFC0&
         Height          =   2010
         Index           =   0
         Left            =   2940
         TabIndex        =   17
         Top             =   960
         Width           =   2115
      End
      Begin VB.ListBox lstSubject 
         BackColor       =   &H00C0FFC0&
         Height          =   2400
         Index           =   1
         ItemData        =   "frmMain.frx":04AD
         Left            =   -74760
         List            =   "frmMain.frx":04B4
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblStudent 
         Caption         =   "Student:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -72840
         TabIndex        =   30
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblStudent 
         Caption         =   "Subject:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2940
         TabIndex        =   29
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblStudent 
         Caption         =   "Student:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblSubject 
         Caption         =   "Subject Code:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   27
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Group:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3660
         TabIndex        =   25
         Top             =   3120
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4620
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   "Ready"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12779
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgOutputDestination 
      Left            =   1440
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.csv"
      DialogTitle     =   "Year 10 Predicted Grades - Export"
      FileName        =   "output.csv"
      Filter          =   "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
      Flags           =   34822
   End
   Begin MSComDlg.CommonDialog dlgInputSource 
      Left            =   780
      Top             =   4980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.csv"
      DialogTitle     =   "Year 10 Predicted Grades - Import"
      FileName        =   "input.csv"
      Filter          =   "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
      Flags           =   38916
      InitDir         =   "c:\work\computing\final"
   End
   Begin VB.Label lblSubjectCode 
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   21
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label lblCandidateNumber 
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   20
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label lblSubjectCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Subject Code:"
      Height          =   195
      Index           =   0
      Left            =   3300
      TabIndex        =   19
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label lblCandidateNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Candidate:"
      Height          =   195
      Index           =   0
      Left            =   3300
      TabIndex        =   18
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label lblStudent 
      Alignment       =   1  'Right Justify
      Caption         =   "Student:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   180
      Width           =   675
   End
   Begin VB.Label lblSubject 
      Height          =   195
      Index           =   1
      Left            =   900
      TabIndex        =   15
      Top             =   420
      Width           =   2295
   End
   Begin VB.Label lblStudent 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   900
      TabIndex        =   14
      Top             =   180
      Width           =   2295
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  'Right Justify
      Caption         =   "Subject:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   420
      Width           =   675
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMenu_Click()

'save changes
Call SaveData(gstrInputDatabaseLocation)
Call ExportData(gstrOutputDatabaseLocation)

frmMain.Hide
frmMenu.Show

End Sub

Private Sub Form_Load()

tabMain.Tab = 0

'clear lists
lstSubject(0).Clear
lstSubject(1).Clear
lstStudent(0).Clear
lstStudent(1).Clear

Call ListStudents
Call ViewRadios(False)
Call ClearAll

End Sub

Private Sub Form_Unload(Cancel As Integer)

'show response dialog
gintResponse = MsgBox("Save changes?", 36, "Year 10 Predicted Grades")

'decide on action to take
If gintResponse = vbYes Then
    Call SaveData(gstrInputDatabaseLocation)
    Call ExportData(gstrOutputDatabaseLocation)
    frmMenu.Show
Else
    frmMenu.Show
End If

End Sub

Private Sub lstStudent_Click(Index As Integer)

Call ClearRadios

'select action depending on what is pressed
Select Case Index

    'first tab
    Case 0
        
        'find wanted student
        Call WantedStudent(Left(lstStudent(0).Text, 4))

        'update fields
        Call UpdateFields(gintCurrentStudent, gintSubjectModifier)

        lstSubject(0).Clear
        For lp1 = 0 To 9
            lstSubject(0).AddItem CaseSubjectName(Left(gstrInputSubjectCode(gintCurrentStudent + lp1), 2))
        Next lp1
        
        lblSubject(1).Caption = ""
        lblSubjectCode(1).Caption = ""
        Call ViewRadios(False)

    'second tab
    Case 1
        
        'find wanted student
        Call WantedStudent(Left(lstStudent(1).Text, 4))
    
        'search for selected subject
        gstrSelectedSubject = lstSubject(1).Text
        For lp1 = 0 To 9
            Select Case gstrSelectedSubject
                Case gstrInputSubjectCode(gintCurrentStudent + lp1)
                    gintSubjectModifier = lp1
            End Select
        Next lp1
    
    'update with new data
    Call UpdateFields(gintCurrentStudent, gintSubjectModifier)

    Call ViewRadios(True)
    
End Select

End Sub

Private Sub lstSubject_Click(Index As Integer)

Call ClearRadios

'select action depening on tab
Select Case Index

    'first tab
    Case 0

        gintSubjectModifier = lstSubject(0).ListIndex
        gblnListClear = True
        Call UpdateFields(gintCurrentStudent, gintSubjectModifier)
        Call ViewRadios(True)
        
        'show group
        Select Case Mid(gstrInputSubjectCode(gintCurrentStudent + gintSubjectModifier), 7, 2)
        Case "__"
            txtGroup.Text = "N/A"
        Case Else
            txtGroup.Text = Mid(gstrInputSubjectCode(gintCurrentStudent + gintSubjectModifier), 7, 2)
        End Select

    'second tab
    Case 1
              
        lstStudent(1).Clear
        
        'find wanted subject code
        gstrWantedSubjectCode = lstSubject(1).Text

        'begin loop
        For lp1 = 1 To 2200
            Select Case gstrInputSubjectCode(lp1)
                Case gstrWantedSubjectCode
                    lstStudent(1).AddItem gstrInputCandidateNumber(lp1) & " - " & gstrInputSurname(lp1) & ", " & gstrInputForename(lp1)
            End Select
        Next lp1
        
        Call UpdateFields(gintCurrentStudent, gintSubjectModifier)
        Call ClearAll
        Call ViewRadios(False)
        
        'show subject data
        lblSubjectCode(1).Caption = lstSubject(1).Text
        lblSubject(1).Caption = CaseSubjectName(Left(lstSubject(1).Text, 2))
        
End Select

End Sub

Private Sub optEffort_Click(Index As Integer)

'update effort data
Select Case Index
    Case 0
        gstrInputEffort(gintCurrentStudent + gintSubjectModifier) = "Excellent"
    Case 1
        gstrInputEffort(gintCurrentStudent + gintSubjectModifier) = "Good"
    Case 2
        gstrInputEffort(gintCurrentStudent + gintSubjectModifier) = "Average"
    Case 3
        gstrInputEffort(gintCurrentStudent + gintSubjectModifier) = "Satisfactory"
    Case 4
        gstrInputEffort(gintCurrentStudent + gintSubjectModifier) = "Poor"
End Select

'display selected effort
optEffort(Index).Value = True

Call UpdateFields(gintCurrentStudent, gintSubjectModifier)

End Sub

Private Sub optGrade_Click(Index As Integer)
       
'update grades data
Select Case Index
    Case 0
        gstrInputGrade(gintCurrentStudent + gintSubjectModifier) = "A* to C"
    Case 1
        gstrInputGrade(gintCurrentStudent + gintSubjectModifier) = "C to E"
    Case 2
        gstrInputGrade(gintCurrentStudent + gintSubjectModifier) = "E to U"
End Select

'display selected grade
optGrade(Index).Value = True

Call UpdateFields(gintCurrentStudent, gintSubjectModifier)
       
End Sub


Private Sub tabMain_Click(PreviousTab As Integer)

'clear subject list box
lstSubject(1).Clear

'update gintCurrentTab variable and display
Select Case PreviousTab
    Case 0
        gintCurrentTab = 1
    Case 1
        gintCurrentTab = 0
    End Select

'open list of subjects
Call ListAllSubjects(gstrSubjectListLocation)

Call ClearAll

End Sub
