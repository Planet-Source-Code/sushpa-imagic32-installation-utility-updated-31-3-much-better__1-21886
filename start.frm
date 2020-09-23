VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "start.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Height          =   125
      Left            =   45
      TabIndex        =   2
      Top             =   405
      Width           =   2445
   End
   Begin VB.CommandButton cmYes 
      Caption         =   "Y&es"
      Default         =   -1  'True
      Height          =   375
      Left            =   2565
      TabIndex        =   1
      Top             =   315
      Width           =   825
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "N&o"
      Height          =   375
      Left            =   3420
      TabIndex        =   0
      Top             =   315
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Begin installation?"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   90
      Width           =   1275
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmno_Click()
EndSetup False
End Sub

Private Sub CmYes_Click()
Me.Hide
frmBack.Show
End Sub

Private Sub Form_Load()
Label1.Caption = "Begin Installation of " & ReadINIFile("Setup", "Title", "Program") & "?"
End Sub
