VERSION 5.00
Begin VB.Form frmRead 
   AutoRedraw      =   -1  'True
   Caption         =   "Terms and conditions"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   Icon            =   "readme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox pStat 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   6150
      TabIndex        =   4
      Top             =   3615
      Visible         =   0   'False
      Width           =   6150
   End
   Begin VB.CommandButton cmYes 
      Caption         =   "I Acc&ept"
      Default         =   -1  'True
      Height          =   375
      Left            =   4050
      TabIndex        =   3
      Top             =   3195
      Width           =   1050
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "I Decli&ne"
      Height          =   375
      Left            =   5130
      TabIndex        =   2
      Top             =   3195
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Height          =   125
      Left            =   45
      TabIndex        =   1
      Top             =   3060
      Width           =   6045
   End
   Begin VB.TextBox txData 
      BackColor       =   &H80000004&
      Height          =   2985
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6045
   End
   Begin VB.Label Label1 
      Caption         =   "CONTROLS ARE RESIZED AT RUNTIME."
      Height          =   240
      Left            =   45
      TabIndex        =   5
      Top             =   3285
      Visible         =   0   'False
      Width           =   3840
   End
End
Attribute VB_Name = "frmRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmNo_Click()
IfReadmeShown = False
Unload Me
End Sub

Private Sub cmYes_Click()
IfReadmeShown = True
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo hell
Dim sText As String
sText = " Please resize this window as needed and read the terms in full before proceeding."
SetFonts Me
pStat.Print sText
ChDir App.Path
Open ReadINIFile("Setup", "Readme", "") For Input As #1
txData.Text = Input(LOF(1), 1)
Close #1
Me.Caption = ReadINIFile("Setup", "ReadmeTitle", "Terms and conditions for " & ReadINIFile("Setup", "Title", "This Program"))
Exit Sub
hell:
MsgBox Error & ".", vbExclamation, "Setup"
Close #1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If IfReadmeShown = True Then Cancel = False: Exit Sub
If MsgBox("Are you sure you do not want to" & vbNewLine & "accept the terms and conditions?" & vbNewLine & vbNewLine & "Setup will close if you decline.", vbOKCancel + vbExclamation, "Readme") = vbOK Then
IfReadmeShown = False
Cancel = False
Else
Cancel = True
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
txData.Width = ScaleWidth
txData.Height = ScaleHeight - Frame1.Height - cmYes.Height - 135 - pStat.Height
Frame1.Width = ScaleWidth
Frame1.Top = txData.Height + 45
cmYes.Top = Frame1.Top + Frame1.Height + 45
cmNo.Top = cmYes.Top
cmNo.Left = ScaleWidth - cmNo.Width - 45
cmYes.Left = cmNo.Left - cmYes.Width - 45
End Sub
