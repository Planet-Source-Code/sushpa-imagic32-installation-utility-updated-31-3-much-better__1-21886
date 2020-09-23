VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Removal"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmQuit 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   395
      Left            =   3375
      TabIndex        =   1
      Top             =   990
      Width           =   870
   End
   Begin VB.PictureBox pS 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   330
      Left            =   225
      ScaleHeight     =   270
      ScaleWidth      =   3960
      TabIndex        =   0
      Top             =   585
      Width           =   4020
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      MouseIcon       =   "main.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "main.frx":0614
      Top             =   990
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Setup is now removing the files that were copied to your system during installation. Please wait."
      Height          =   420
      Left            =   225
      TabIndex        =   2
      Top             =   90
      Width           =   4020
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DoneValue As Integer

Private Sub cmQuit_Click()
End
End Sub

Private Sub Form_Click()
DoneValue = DoneValue + 1
UpdateStatus pS, DoneValue
End Sub

Private Sub Form_Load()
If MsgBox("Are you sure you want to remove " & ReadINIFile("Uninstall", "Title", "this program") & "?", vbQuestion + vbYesNo, "Remove") = vbNo Then End
Me.Caption = ReadINIFile("Uninstall", "Title", "Program") & " Removal"
End Sub

Private Sub Image1_Click()
MsgBox "InstallMagic Utility By Sushant Pandurangi" & vbCrLf & "Copyright Sushant Pandurangi, 2000-2001." & vbCrLf & "Please send mails to sushant@phreaker.net." & vbCrLf & "Or visit http://sushantshome.tripod.com." & vbCrLf & vbCrLf & "Distributed as non-commercial freeware.", vbInformation, "About Setup"
End Sub
