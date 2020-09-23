VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRMGEN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Readme Generator // Jaime Muscatelli"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8775
   Icon            =   "FRMGEN.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4200
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrcaption 
      Interval        =   10000
      Left            =   3840
      Top             =   2760
   End
   Begin VB.Frame FRAAditional 
      Caption         =   "&Additional Arguments"
      Height          =   2295
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   8535
      Begin VB.TextBox txtAdditionalArguments 
         Height          =   1935
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame FRASTANDARD 
      Caption         =   "Standard Arguments"
      Height          =   3615
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton CMDABOUT 
         Caption         =   "&About"
         Height          =   195
         Left            =   6240
         TabIndex        =   19
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton CMDPREVIEW 
         Caption         =   "&Preview Project"
         Height          =   375
         Left            =   6240
         TabIndex        =   18
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "&Save Project"
         Height          =   375
         Left            =   6240
         TabIndex        =   17
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Frame FRAAPPPROPS 
         Caption         =   "Programming Information"
         Height          =   1695
         Left            =   5760
         TabIndex        =   23
         Top             =   240
         Width           =   2655
         Begin VB.TextBox txtOS_MadeFor 
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Text            =   "Compatible Operating Systems"
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox txtOS_MadeIn 
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Text            =   "Operating System Designed In"
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox txtLanguage 
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Text            =   "Language Written In"
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame fraAuthor 
         Caption         =   "Author Information"
         Height          =   2655
         Left            =   3000
         TabIndex        =   22
         Top             =   240
         Width           =   2655
         Begin VB.TextBox txtWebsite 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Text            =   "Website"
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtScreenName 
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Text            =   "Screen Name"
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox txtCompany 
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Text            =   "Company"
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox txtEmail 
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Text            =   "Email"
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox txtAuthor 
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Text            =   "Author"
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame FRAAPP 
         Caption         =   "File Properties"
         Height          =   3135
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2655
         Begin VB.TextBox txtCopyright 
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Text            =   "© (Copyright)"
            Top             =   2640
            Width           =   2415
         End
         Begin VB.TextBox txtFileSize 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Text            =   "File Size"
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtVersionNumber 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Text            =   "Version Number"
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox txtDescription 
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Text            =   "Description"
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox txtFileName 
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Text            =   "File Name"
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox txtTitle 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Text            =   "Title"
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmdTEMPSAVE 
         Caption         =   "&Temp Save"
         Height          =   375
         Left            =   6240
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilenew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnufileline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufilesave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnufileline2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "&Exit"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "FRMGEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public README As New clsReadme

Private Sub CMDABOUT_Click()
MsgBox "Author: Jaime Muscatelli " & vbCrLf & "Website: www.jprogs.cjb.net" & vbCrLf & "Email: webmaster@jaimemuscatelli.zzn.com" & vbCrLf & "AOL SN: Jaime141974" & vbCrLf & "Language: Visual Basic", vbOKOnly, Me.Caption
End Sub

Private Sub CMDPREVIEW_Click()
README.PreviewProject
End Sub

Private Sub CMDSAVE_Click()
README.UpdateProject
README.SaveProject
End Sub

Private Sub cmdTEMPSAVE_Click()
README.UpdateProject
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sResult As String
Cancel = True

sResult = MsgBox("Are you sure you wish to exit?", vbYesNo + vbQuestion + vbSystemModal, Me.Caption)

If sResult = vbNo Then
Exit Sub
ElseIf sResult = vbYes Then
Set README = Nothing
End
End If
End Sub

Private Sub mnufileclose_Click()

End Sub

Private Sub mnufileexit_Click()
Call Form_Unload(True)
End Sub

Private Sub mnufilenew_Click()
README.NewProject
End Sub

Private Sub mnufilesave_Click()
README.SaveProject
End Sub

Private Sub tmrcaption_Timer()
Me.Caption = "Readme Generator // Jaime Muscatelli"
End Sub

Private Sub txtFileSize_KeyPress(KeyAscii As Integer)
Select Case KeyAscii

Case 48 To 57

Case 8

Case 32

Case Else
MsgBox "Please enter a number value", vbOKOnly + vbExclamation, Me.Caption
txtFileSize.Text = vbNullString
KeyAscii = 0
End Select
End Sub

Private Sub txtVersionNumber_KeyPress(KeyAscii As Integer)
Select Case KeyAscii

Case 48 To 57

Case 8

Case 32

Case Else
MsgBox "Please enter a number value", vbOKOnly + vbExclamation, Me.Caption
txtFileSize.Text = vbNullString
KeyAscii = 0
End Select
End Sub
