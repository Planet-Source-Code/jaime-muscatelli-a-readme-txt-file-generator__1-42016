VERSION 5.00
Begin VB.Form FRMPREVIEW 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Readme Previewer"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   7575
      Left            =   120
      ScaleHeight     =   7515
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.TextBox txtPreview 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   7335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   120
         Width           =   9375
      End
   End
End
Attribute VB_Name = "FRMPREVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
FRMGEN.Show
Unload Me
End Sub
