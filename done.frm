VERSION 5.00
Begin VB.Form frmOK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finished"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "done.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   135
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&lose"
      Default         =   -1  'True
      Height          =   375
      Left            =   3150
      TabIndex        =   3
      Top             =   945
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   135
      Picture         =   "done.frx":000C
      ScaleHeight     =   435
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   765
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "done.frx":0115
      Top             =   225
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "The installation has been completed. Thank you for using InstallMagic. $App$ has been properly set up on your computer. "
      Height          =   780
      Left            =   765
      TabIndex        =   0
      Top             =   225
      Width           =   3525
   End
   Begin VB.Label lbTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   135
      Visible         =   0   'False
      Width           =   3525
   End
End
Attribute VB_Name = "frmOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetFonts Me
lbTitle.Caption = ReadINIFile("Setup", "Title", "Unknown Program")
Label1.Caption = Replace(Label1.Caption, "$App$", ReadINIFile("Setup", "Title", "The Program"))
lbTitle.Font.Name = ReadINIFile("Setup", "TitleFont", "Tahoma")
lbTitle.Font.Bold = CBool(ReadINIFile("Setup", "TitleBold", 0))
lbTitle.Font.Italic = CBool(ReadINIFile("Setup", "TitleItalic", 0))
End Sub

