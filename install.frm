VERSION 5.00
Begin VB.Form frmInst 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pStatus 
      AutoRedraw      =   -1  'True
      Height          =   330
      Left            =   675
      ScaleHeight     =   270
      ScaleWidth      =   4680
      TabIndex        =   1
      Top             =   90
      Width           =   4740
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   360
      Top             =   495
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   90
      Picture         =   "install.frx":0000
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait while initializing the setup process..."
      Height          =   195
      Left            =   675
      TabIndex        =   0
      Top             =   450
      Width           =   4740
   End
End
Attribute VB_Name = "frmInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Max As Long
Public Total As Integer
Public File As String
Dim TEMP As String

Private Sub Timer1_Timer()
SetFonts Me
Total = ReadINIFile("Setup", "TotalFiles", 0)
TEMP = ReadINIFile("Setup", "ShortcutFile", "File" & Total)
File = ReadINIFile(TEMP, "Source", "")
InstallFiles
UpdateStatus pStatus, Max
Me.Hide
frmOK.Show vbModal
EndSetup True
End Sub

