VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year 10 Predicted Grades - Menu"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Height          =   855
      Left            =   300
      Picture         =   "frmMenu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Letters"
      Top             =   2160
      Width           =   915
   End
   Begin VB.CommandButton cmdView 
      Height          =   855
      Left            =   300
      Picture         =   "frmMenu.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "View Data"
      Top             =   1200
      Width           =   915
   End
   Begin VB.CommandButton cmdMain 
      Height          =   855
      Left            =   300
      Picture         =   "frmMenu.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Edit Data"
      Top             =   240
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Height          =   855
      Left            =   300
      Picture         =   "frmMenu.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Save and Exit"
      Top             =   4080
      Width           =   915
   End
   Begin VB.CommandButton cmdOptions 
      Height          =   855
      Left            =   300
      Picture         =   "frmMenu.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "File Locations"
      Top             =   3120
      Width           =   915
   End
   Begin MSComDlg.CommonDialog dlgBasicDatabaseLocation 
      Left            =   3660
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.csv"
      DialogTitle     =   "Student Database Location"
      FileName        =   "studentdata.csv"
      Filter          =   "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
      Flags           =   38916
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   60
      TabIndex        =   15
      Top             =   5160
      Width           =   4035
   End
   Begin VB.Label lblPrint 
      Caption         =   "Using Microsoft Word merge tools."
      Height          =   195
      Index           =   1
      Left            =   1380
      TabIndex        =   14
      Top             =   2640
      Width           =   2475
   End
   Begin VB.Label lblPrint 
      Caption         =   "Print Letters"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1380
      TabIndex        =   13
      Top             =   2280
      Width           =   2475
   End
   Begin VB.Label lblView 
      Caption         =   "View all the data at a glance."
      Height          =   195
      Index           =   1
      Left            =   1380
      TabIndex        =   12
      Top             =   1680
      Width           =   2475
   End
   Begin VB.Label lblView 
      Caption         =   "View Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1380
      TabIndex        =   11
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Label lblExit 
      Caption         =   "For when you have finished."
      Height          =   195
      Index           =   1
      Left            =   1380
      TabIndex        =   10
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label lblExit 
      Caption         =   "Save and Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1380
      TabIndex        =   9
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label lblOptions 
      Caption         =   "Use this to locate necessary files."
      Height          =   195
      Index           =   1
      Left            =   1380
      TabIndex        =   8
      Top             =   3600
      Width           =   2475
   End
   Begin VB.Label lblOptions 
      Caption         =   "File Locations"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1380
      TabIndex        =   7
      Top             =   3240
      Width           =   2475
   End
   Begin VB.Label lblMain 
      Caption         =   "Edit, print, save and export."
      Height          =   195
      Index           =   1
      Left            =   1380
      TabIndex        =   6
      Top             =   720
      Width           =   2475
   End
   Begin VB.Label lblMain 
      Caption         =   "Edit Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1380
      TabIndex        =   5
      Top             =   360
      Width           =   2475
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

'show response dialog
gintResponse = MsgBox("Are you sure you want to exit?", 36, "Year 10 Predicted Grades")

'decide on action to take
If gintResponse = vbYes Then
    Call SaveData(gstrInputDatabaseLocation)
    Call ExportData(gstrOutputDatabaseLocation)
    End
Else
    Exit Sub
End If

End Sub

Private Sub cmdMain_Click()

'display data editing form
frmMenu.Hide
frmMain.Show

End Sub

Private Sub cmdOptions_Click()

'display options form
frmMenu.Hide
frmOptions.Show

End Sub

Private Sub cmdPrint_Click()

'append filename to end
Dim strFullWordPath As String
strFullWordPath = gstrWordLocation & " " & Chr(34) & gstrOutputDocumentLocation & Chr(34)

'open letter.doc with Microsoft Word
Call Shell(strFullWordPath, vbMaximizedFocus)

End Sub

Private Sub cmdView_Click()

'display viewing form
frmMenu.Hide
frmView.Show

End Sub

Private Sub Form_Load()


'append extra character if running from the root of a drive
If Len(App.Path) < 4 Then
    gstrAppPath = App.Path & ""
Else
    gstrAppPath = App.Path & "\"
End If


Call LoadSettings

If Exists(gstrInputDatabaseLocation) = False Or Exists(gstrBasicDatabaseLocation) = False Then
    
    MsgBox "You appear to be running this program for the first time. Please find the student database file you wish to use using the dialog box provided.", vbInformation, "Year 10 Predicted Grades"
    
    On Error GoTo CancelError
    
    dlgBasicDatabaseLocation.ShowOpen
    gstrBasicDatabaseLocation = dlgBasicDatabaseLocation.filename
    
    On Error Resume Next
    
    Call DatabaseSize(gstrBasicDatabaseLocation)
    
    Select Case gintNumberOfFields
    Case 4
        Call Add2UnknownFields(gstrBasicDatabaseLocation)
    Case Else
        MsgBox "Student database is corrupt.", vbCritical, "Year 10 Predicted Grades"
        End
    End Select

End If

Call DatabaseSize(gstrInputDatabaseLocation)

Call CheckSettings

'create settings file
Open gstrSettingsLocation For Output As #1
    Print #1, gstrOutputDocumentLocation
    Print #1, gstrInputDatabaseLocation
    Print #1, gstrOutputDatabaseLocation
    Print #1, gstrSubjectListLocation
    Print #1, gstrWordLocation
    Print #1, gstrBasicDatabaseLocation
Close #1

'resize input arrays
ReDim gstrInputSubjectCode(1 To gintNumberOfRecords) As String
ReDim gstrInputForename(1 To gintNumberOfRecords) As String
ReDim gstrInputSurname(1 To gintNumberOfRecords) As String
ReDim gstrInputCandidateNumber(1 To gintNumberOfRecords) As String
ReDim gstrInputGrade(1 To gintNumberOfRecords) As String
ReDim gstrInputEffort(1 To gintNumberOfRecords) As String
ReDim gstrInputSubjectName(1 To gintNumberOfRecords) As String

'resize output arrays
ReDim gstrOutputName(1 To (gintNumberOfRecords / 10)) As String
ReDim gstrOutputSurname(1 To (gintNumberOfRecords / 10)) As String
ReDim gstrOutputForename(1 To (gintNumberOfRecords / 10)) As String
ReDim gstrOutputCandidateNumber(1 To (gintNumberOfRecords / 10)) As String
ReDim gstrOutputSubject(1 To 10, 1 To (gintNumberOfRecords / 10)) As String
ReDim gstrOutputGrade(1 To 10, 1 To (gintNumberOfRecords / 10)) As String
ReDim gstrOutputEffort(1 To 10, 1 To (gintNumberOfRecords / 10)) As String

'select data source and import
Call LoadData(gstrInputDatabaseLocation)

'display title and version number
lblVersion.Caption = "(c) Tom Robinson 2000, v" & App.Major & "." & App.Minor & "." & App.Revision

Exit Sub

CancelError:
    End

End Sub

Private Sub Form_Unload(Cancel As Integer)

'show response dialog
gintResponse = MsgBox("Save changes?", 36, "Year 10 Predicted Grades")

'decide on action to take
If gintResponse = vbYes Then
    Call SaveData(gstrInputDatabaseLocation)
    Call ExportData(gstrOutputDatabaseLocation)
    End
Else
    End
End If

End Sub
