VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year 10 Predicted Grades - View Data"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgLoad 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6165
      _Version        =   393216
      Appearance      =   1
      Min             =   1
      Max             =   2200
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Main Menu"
      Height          =   375
      Left            =   2580
      Picture         =   "frmView.frx":0442
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDisplayOutput 
      Caption         =   "View &Output Database"
      Height          =   375
      Left            =   4140
      TabIndex        =   1
      Top             =   120
      Width           =   2235
   End
   Begin MSFlexGridLib.MSFlexGrid flexView 
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6165
      _Version        =   393216
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdDisplayInput 
      Caption         =   "View &Input Database"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

'return to menu
frmView.Hide
frmMenu.Show

End Sub

Private Sub cmdDisplayInput_Click()

'set up form
flexView.Visible = False
cmdDisplayOutput.Enabled = False
cmdDisplayInput.Enabled = False
cmdBack.Enabled = False
flexView.Visible = False
prgLoad.Visible = True

'set properties of grid
With flexView
    .Clear
    .Cols = 6
    .Rows = 2201
    .FixedCols = 0
End With

prgLoad.Max = flexView.Rows

'create column headers
flexView.Row = 0
flexView.Col = 0
flexView.Text = "Subject Code"
flexView.Col = 1
flexView.Text = "Surname"
flexView.Col = 2
flexView.Text = "Forename"
flexView.Col = 3
flexView.Text = "Candidate Number"
flexView.Col = 4
flexView.Text = "Grade"
flexView.Col = 5
flexView.Text = "Effort"

'display data on grid
For lp1 = 1 To 2200

    flexView.Row = lp1
    flexView.Col = 0
    flexView.Text = gstrInputSubjectCode(lp1)
    flexView.Col = 1
    flexView.Text = gstrInputSurname(lp1)
    flexView.Col = 2
    flexView.Text = gstrInputForename(lp1)
    flexView.Col = 3
    flexView.Text = gstrInputCandidateNumber(lp1)
    flexView.Col = 4
    flexView.Text = gstrInputGrade(lp1)
    flexView.Col = 5
    flexView.Text = gstrInputEffort(lp1)
    
    prgLoad.Value = lp1
    
Next lp1

'set up form
flexView.Visible = True
prgLoad.Visible = False
cmdDisplayInput.Enabled = False
cmdBack.Enabled = True
cmdDisplayOutput.Enabled = True

End Sub

Private Sub cmdDisplayOutput_Click()

'set up form
cmdDisplayOutput.Enabled = False
cmdDisplayInput.Enabled = False
cmdBack.Enabled = False
flexView.Visible = True
prgLoad.Visible = True

'set grid properties
With flexView
    .Clear
    .Cols = 32
    .Rows = 221
    .FixedCols = 1
End With

prgLoad.Max = flexView.Rows

'create column headers
flexView.Row = 0
flexView.Col = 0
flexView.Text = "Name"
flexView.Col = 1
flexView.Text = "Candidate Number"
flexView.Col = 2
flexView.Text = "Subject (1)"
flexView.Col = 3
flexView.Text = "Grade (1)"
flexView.Col = 4
flexView.Text = "Effort (1)"
flexView.Col = 5
flexView.Text = "Subject (2)"
flexView.Col = 6
flexView.Text = "Grade (2)"
flexView.Col = 7
flexView.Text = "Effort (2)"
flexView.Col = 8
flexView.Text = "Subject (3)"
flexView.Col = 9
flexView.Text = "Grade (3)"
flexView.Col = 10
flexView.Text = "Effort (3)"
flexView.Col = 11
flexView.Text = "Subject (4)"
flexView.Col = 12
flexView.Text = "Grade (4)"
flexView.Col = 13
flexView.Text = "Effort (4)"
flexView.Col = 14
flexView.Text = "Subject (5)"
flexView.Col = 15
flexView.Text = "Grade (5)"
flexView.Col = 16
flexView.Text = "Effort (5)"
flexView.Col = 17
flexView.Text = "Subject (6)"
flexView.Col = 18
flexView.Text = "Grade (6)"
flexView.Col = 19
flexView.Text = "Effort (6)"
flexView.Col = 20
flexView.Text = "Subject (7)"
flexView.Col = 21
flexView.Text = "Grade (7)"
flexView.Col = 22
flexView.Text = "Effort (7)"
flexView.Col = 23
flexView.Text = "Subject (8)"
flexView.Col = 24
flexView.Text = "Grade (8)"
flexView.Col = 25
flexView.Text = "Effort (8)"
flexView.Col = 26
flexView.Text = "Subject (9)"
flexView.Col = 27
flexView.Text = "Grade (9)"
flexView.Col = 28
flexView.Text = "Effort (9)"
flexView.Col = 29
flexView.Text = "Subject (10)"
flexView.Col = 30
flexView.Text = "Grade (10)"
flexView.Col = 31
flexView.Text = "Effort (10)"

'display data on grid
For lp1 = 1 To 220
    flexView.Row = lp1
    flexView.Col = 0
    flexView.Text = gstrOutputName(lp1)
    flexView.Col = 1
    flexView.Text = gstrOutputCandidateNumber(lp1)
    flexView.Col = 2
    flexView.Text = gstrOutputSubject(1, lp1)
    flexView.Col = 3
    flexView.Text = gstrOutputGrade(1, lp1)
    flexView.Col = 4
    flexView.Text = gstrOutputEffort(1, lp1)
    flexView.Col = 5
    flexView.Text = gstrOutputSubject(2, lp1)
    flexView.Col = 6
    flexView.Text = gstrOutputGrade(2, lp1)
    flexView.Col = 7
    flexView.Text = gstrOutputEffort(2, lp1)
    flexView.Col = 8
    flexView.Text = gstrOutputSubject(3, lp1)
    flexView.Col = 9
    flexView.Text = gstrOutputGrade(3, lp1)
    flexView.Col = 10
    flexView.Text = gstrOutputEffort(3, lp1)
    flexView.Col = 11
    flexView.Text = gstrOutputSubject(4, lp1)
    flexView.Col = 12
    flexView.Text = gstrOutputGrade(4, lp1)
    flexView.Col = 13
    flexView.Text = gstrOutputEffort(4, lp1)
    flexView.Col = 14
    flexView.Text = gstrOutputSubject(5, lp1)
    flexView.Col = 15
    flexView.Text = gstrOutputGrade(5, lp1)
    flexView.Col = 16
    flexView.Text = gstrOutputEffort(5, lp1)
    flexView.Col = 17
    flexView.Text = gstrOutputSubject(6, lp1)
    flexView.Col = 18
    flexView.Text = gstrOutputGrade(6, lp1)
    flexView.Col = 19
    flexView.Text = gstrOutputEffort(6, lp1)
    flexView.Col = 20
    flexView.Text = gstrOutputSubject(7, lp1)
    flexView.Col = 21
    flexView.Text = gstrOutputGrade(7, lp1)
    flexView.Col = 22
    flexView.Text = gstrOutputEffort(7, lp1)
    flexView.Col = 23
    flexView.Text = gstrOutputSubject(8, lp1)
    flexView.Col = 24
    flexView.Text = gstrOutputGrade(8, lp1)
    flexView.Col = 25
    flexView.Text = gstrOutputEffort(8, lp1)
    flexView.Col = 26
    flexView.Text = gstrOutputSubject(9, lp1)
    flexView.Col = 27
    flexView.Text = gstrOutputGrade(9, lp1)
    flexView.Col = 28
    flexView.Text = gstrOutputEffort(9, lp1)
    flexView.Col = 29
    flexView.Text = gstrOutputSubject(10, lp1)
    flexView.Col = 30
    flexView.Text = gstrOutputGrade(10, lp1)
    flexView.Col = 31
    flexView.Text = gstrOutputEffort(10, lp1)

    prgLoad.Value = lp1
    
Next lp1

'set up form
flexView.Visible = True
prgLoad.Visible = False
cmdDisplayInput.Enabled = True
cmdBack.Enabled = True

End Sub

Private Sub Form_Load()

'find output variables
cnt2 = 1

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

For lp3 = 1 To 220
    gstrOutputName(lp3) = gstrOutputForename(lp3) & " " & gstrOutputSurname(lp3)
Next lp3

End Sub

Private Sub Form_Unload(Cancel As Integer)

frmMenu.Show

End Sub
