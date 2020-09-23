VERSION 5.00
Begin VB.Form frmBack 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Install Magic for Windows 95/98"
   ClientHeight    =   8595
   ClientLeft      =   -4065
   ClientTop       =   -1905
   ClientWidth     =   11880
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tGrade 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2925
      Top             =   2745
   End
   Begin VB.Timer tmDraw 
      Interval        =   1
      Left            =   5940
      Top             =   225
   End
   Begin VB.Timer tmGo 
      Interval        =   500
      Left            =   6525
      Top             =   225
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "We'll be using the app's name to find whether this app is being used for removing or installing. "
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   45
      TabIndex        =   5
      Top             =   4050
      Visible         =   0   'False
      Width           =   6675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"main.frx":0442
      ForeColor       =   &H00C0FFFF&
      Height          =   870
      Index           =   1
      Left            =   3555
      TabIndex        =   4
      Top             =   2385
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"main.frx":04CC
      ForeColor       =   &H00C0FFFF&
      Height          =   1635
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   2385
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "main.frx":05C3
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbCR 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright Information goes here."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   45
      TabIndex        =   2
      Top             =   630
      Width           =   2280
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"main.frx":08CD
      ForeColor       =   &H00C0E0FF&
      Height          =   1455
      Left            =   45
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   6150
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(Application Title) Setup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   630
      TabIndex        =   0
      Top             =   90
      Width           =   4350
   End
End
Attribute VB_Name = "frmBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GradeColour As Integer

Private Sub Form_Load()
SetFonts Me
GradeForm Me, , , 1000
lbTitle.FontBold = CBool(ReadINIFile("Setup", "TitleBold", 0))
lbTitle.FontItalic = CBool(ReadINIFile("Setup", "TitleItalic", 1))
lbTitle.Caption = ReadINIFile("Setup", "Title", "Program") & " Setup"
lbTitle.Font.Name = Trim(ReadINIFile("Setup", "TitleFont", "Arial"))
lbCR.Caption = Trim(ReadINIFile("Setup", "Copyright", ""))
End Sub

Private Sub tGrade_Timer()
'The below if statememt actually regulates how
'many colours to grade through. if you only want
'red green and blue then change the 11 to 2.
If GradeColour = 11 Then GradeColour = -1
'-1 is used because we need a 0 as well (red) so
'(-1) + (1) will give 0 which we need to draw red.
GradeColour = GradeColour + 1 'Increment
GradeForm Me, GradeColour, 1, 1000
Me.Refresh
End Sub

Private Sub tmDraw_Timer()
Screen.MousePointer = 11
GradeForm Me
Me.Refresh
Screen.MousePointer = 0
tmDraw.Enabled = False
End Sub

Private Sub TmGo_Timer()
'Enables to show frmBegin after a delay
'Creates some *professional* feel
frmWelcome.Show vbModal, Me
tmGo.Enabled = False
End Sub
