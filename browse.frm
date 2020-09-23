VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Smart Installer - (Application Title)"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   Icon            =   "browse.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmOK 
      Caption         =   "N&ext"
      Default         =   -1  'True
      Height          =   375
      Left            =   3015
      TabIndex        =   5
      Top             =   1845
      Width           =   960
   End
   Begin VB.CommandButton cmX 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4005
      TabIndex        =   4
      Top             =   1845
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   90
      TabIndex        =   3
      Top             =   1665
      Width           =   4875
   End
   Begin VB.Image Command1 
      Height          =   480
      Left            =   180
      MouseIcon       =   "browse.frx":000C
      MousePointer    =   99  'Custom
      Picture         =   "browse.frx":0316
      ToolTipText     =   "About Setup"
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Click 'N&ext' to continue with the Installation process."
      Height          =   240
      Left            =   900
      TabIndex        =   2
      Top             =   1395
      Width           =   4020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"browse.frx":0620
      Height          =   825
      Left            =   900
      TabIndex        =   1
      Top             =   450
      Width           =   3975
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Application Title) Installation"
      Height          =   195
      Left            =   900
      TabIndex        =   0
      Top             =   135
      Width           =   2010
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "browse.frx":06D9
      Top             =   585
      Width           =   480
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmOK_Click()
Me.Hide
If ShowReadme = True Then frmBegin.Show vbModal, frmBack Else EndSetup
End Sub

Private Sub cmX_Click()
Unload Me
End Sub

Sub Command1_Click()
MsgBox "InstallMagic Utility By Sushant Pandurangi" & vbCrLf & "Copyright Sushant Pandurangi, 2000-2001." & vbCrLf & "Please send mails to sushant@phreaker.net." & vbCrLf & "Or visit http://sushantshome.tripod.com." & vbCrLf & vbCrLf & "Distributed as non-commercial freeware.", vbInformation, "About Setup"
End Sub

Private Sub Form_Load()
SetFonts Me
Me.Caption = "Install Magic - " & ReadINIFile("Setup", "Title", "Program")
lbTitle.Caption = ReadINIFile("Setup", "Title", "Program") & " Installation"
Label2.Caption = Replace(Label2.Caption, "copy", "remove")
Label2.Caption = Replace(Label2.Caption, "to", "from")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If LastRites() = vbYes Then
EndSetup
Else
Cancel = True
End If
End Sub
