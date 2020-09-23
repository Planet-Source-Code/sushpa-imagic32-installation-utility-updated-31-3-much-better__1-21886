VERSION 5.00
Begin VB.Form frmBrowse 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse for Folder"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   Icon            =   "welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Can&cel"
      Height          =   555
      Left            =   2250
      Picture         =   "welcome.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1710
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do&ne"
      Default         =   -1  'True
      Height          =   600
      Left            =   2250
      Picture         =   "welcome.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   825
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   105
      TabIndex        =   3
      Top             =   810
      Width           =   2040
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   225
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2430
      MouseIcon       =   "welcome.frx":0733
      MousePointer    =   99  'Custom
      Picture         =   "welcome.frx":0A3D
      ToolTipText     =   "About Setup"
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   630
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drives:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   45
      Width           =   495
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
bPath = UCase(Left(Drive1.Drive, 2)) & "\" & Right(Dir1.Path, Len(Dir1.Path) - 3)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
On Error GoTo hell
Dir1.Path = Drive1.Drive
Exit Sub
hell:
If MsgBox("Device cannot be accessed" & vbNewLine & "Try again?", vbRetryCancel + vbExclamation, "Setup") = vbRetry Then Drive1_Change 'GoTo start
End Sub

Private Sub Form_Load()
SetFonts Me
End Sub

Private Sub Image1_Click()
frmWelcome.Command1_Click
End Sub
