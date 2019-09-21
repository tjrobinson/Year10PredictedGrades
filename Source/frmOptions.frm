VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year 10 Predicted Grades - File Locations"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDocumentLocation 
      Caption         =   "Output Document Location"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   5235
      Begin VB.CommandButton cmbBrowseForDocument 
         Caption         =   "Bro&wse..."
         Height          =   315
         Left            =   3900
         TabIndex        =   4
         ToolTipText     =   "Locate the Output document"
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtDocumentLocation 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   300
         Width           =   3675
      End
   End
   Begin VB.Frame fraInputDatabaseLocation 
      Caption         =   "Input Database Location"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5235
      Begin VB.TextBox txtInputDatabaseLocation 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   300
         Width           =   3675
      End
      Begin VB.CommandButton cmdBrowseForInputDatabase 
         Caption         =   "B&rowse..."
         Height          =   315
         Left            =   3900
         TabIndex        =   2
         ToolTipText     =   "Locate the Input Database"
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame fraOutputDatabaseLocation 
      Caption         =   "Output Database Location"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   5235
      Begin VB.CommandButton cmbBrowseOutputDatabase 
         Caption         =   "Br&owse..."
         Height          =   315
         Left            =   3900
         TabIndex        =   3
         ToolTipText     =   "Locate the Output Database"
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtOutputDatabaseLocation 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   300
         Width           =   3675
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4020
      TabIndex        =   1
      Top             =   2700
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   2700
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgLocateDocument 
      Left            =   660
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.doc"
      DialogTitle     =   "Output Document Location"
      FileName        =   "grades.doc"
      Filter          =   "Word Documents (*.doc)|*.doc|All Files (*.*)|*.*"
      Flags           =   4228
   End
   Begin MSComDlg.CommonDialog dlgInputDatabaseLocation 
      Left            =   1140
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.csv"
      DialogTitle     =   "Input Database Location"
      FileName        =   "input.csv"
      Filter          =   "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
      Flags           =   38916
   End
   Begin MSComDlg.CommonDialog dlgOutputDatabaseLocation 
      Left            =   1620
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.csv"
      DialogTitle     =   "Output Database Location"
      FileName        =   "output.csv"
      Filter          =   "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
      Flags           =   34822
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbBrowseForDocument_Click()

On Error GoTo CancelError

'show common dialog box
dlgLocateDocument.ShowOpen

'assign to variable and update text box
gstrOutputDocumentLocation = dlgLocateDocument.filename
txtDocumentLocation.Text = gstrOutputDocumentLocation

Exit Sub

CancelError:
    Exit Sub

End Sub

Private Sub cmbBrowseOutputDatabase_Click()

On Error GoTo CancelError

'show common dialog box
dlgOutputDatabaseLocation.ShowOpen

'assign to variable and update text box
gstrOutputDatabaseLocation = dlgOutputDatabaseLocation.filename
txtOutputDatabaseLocation.Text = gstrOutputDatabaseLocation

Exit Sub

CancelError:
    Exit Sub

End Sub

Private Sub cmdBrowseForInputDatabase_Click()

On Error GoTo CancelError

'show common dialog box
dlgInputDatabaseLocation.ShowOpen

'assign to variable and update text box
gstrInputDatabaseLocation = dlgInputDatabaseLocation.filename
txtInputDatabaseLocation.Text = gstrInputDatabaseLocation

Exit Sub

CancelError:
    Exit Sub

End Sub

Private Sub cmdCancel_Click()

'hide options form without saving changes
frmOptions.Hide
frmMenu.Show

End Sub

Private Sub cmdOK_Click()

'save settings
Call SaveSettings

'hide options form
frmOptions.Hide
frmMenu.Show

End Sub

Private Sub Form_Load()

'set common dialog default to application path
dlgInputDatabaseLocation.InitDir = App.Path
dlgOutputDatabaseLocation.InitDir = App.Path
dlgLocateDocument.InitDir = App.Path

'load location settings
Call LoadSettings

'update text boxes
txtDocumentLocation.Text = gstrOutputDocumentLocation
txtInputDatabaseLocation.Text = gstrInputDatabaseLocation
txtOutputDatabaseLocation.Text = gstrOutputDatabaseLocation

End Sub

Private Sub Form_Unload(Cancel As Integer)

frmMenu.Show

End Sub
