VERSION 5.00
Begin VB.Form frmEnd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm Exit"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "exit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "R&esume"
      Height          =   385
      Left            =   4140
      TabIndex        =   3
      Top             =   855
      Width           =   915
   End
   Begin VB.CommandButton cmQ 
      Caption         =   "&Quit"
      Default         =   -1  'True
      Height          =   385
      Left            =   3285
      TabIndex        =   2
      Top             =   855
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   $"exit.frx":000C
      Height          =   825
      Left            =   630
      TabIndex        =   1
      Top             =   90
      Width           =   4425
   End
   Begin VB.Label lblT 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   540
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "exit.frx":00DA
      Top             =   225
      Width           =   480
   End
End
Attribute VB_Name = "frmEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmQ_Click()
ReturnedExit = vbYes
Unload Me
End Sub

Private Sub Command1_Click()
ReturnedExit = vbNo
Unload Me
End Sub

Private Sub Form_Load()
SetFonts Me
lblT.Caption = ReadINIFile("Setup", "Title", "Installation")
End Sub
