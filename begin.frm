VERSION 5.00
Begin VB.Form frmBegin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Beginning Installation"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "begin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmStart 
      Caption         =   "S&tart"
      Default         =   -1  'True
      Height          =   825
      Left            =   4500
      Picture         =   "begin.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   870
   End
   Begin VB.CommandButton cmX 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   825
      Left            =   5400
      Picture         =   "begin.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   825
   End
   Begin VB.Frame Frame3 
      Caption         =   "Installation"
      Height          =   125
      Left            =   135
      TabIndex        =   9
      Top             =   1845
      Width           =   4110
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Shortcut"
      Height          =   1140
      Left            =   3375
      TabIndex        =   6
      Top             =   135
      Width           =   2760
      Begin VB.TextBox txName 
         Height          =   315
         Left            =   630
         TabIndex        =   8
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1485
         MouseIcon       =   "begin.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Type in the 'Nam&e' of the start menu and desktop shortcut."
         Height          =   420
         Left            =   540
         TabIndex        =   12
         Top             =   675
         Width           =   2130
      End
      Begin VB.Image Image2 
         Height          =   435
         Left            =   60
         Picture         =   "begin.frx":0C28
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Nam&e:"
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cr&eate shortcuts in the start menu and on the desktop"
      Height          =   375
      Left            =   3375
      TabIndex        =   5
      Top             =   1350
      UseMaskColor    =   -1  'True
      Value           =   1  'Checked
      Width           =   2670
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pa&th"
      Height          =   1500
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   3165
      Begin VB.CommandButton cmBrowse 
         Caption         =   "B&rowse..."
         Height          =   600
         Left            =   2160
         Picture         =   "begin.frx":0D31
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   765
         Width           =   890
      End
      Begin VB.TextBox txPath 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   385
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1350
         MouseIcon       =   "begin.frx":0ED8
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   1215
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   60
         Picture         =   "begin.frx":11E2
         Top             =   855
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Specify the path to Install to, or choose it by clicking 'B&rowse'."
         Height          =   645
         Left            =   585
         TabIndex        =   4
         Top             =   810
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inst&allation Path:"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   180
         Width           =   1170
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   150
      Left            =   1530
      MouseIcon       =   "begin.frx":12EB
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   2385
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Installation"
      Height          =   195
      Left            =   270
      TabIndex        =   14
      Top             =   1845
      Width           =   750
   End
   Begin VB.Label Label5 
      Caption         =   "When you have completed making your selections, simply click the 'S&tart' button on the right."
      Height          =   465
      Left            =   405
      TabIndex        =   13
      Top             =   2160
      Width           =   3840
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   90
      Picture         =   "begin.frx":15F5
      Top             =   2115
      Width           =   210
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
'The below line makes the entire frame
'disappear if checkbox is unchecked.
'-----------------------------------------------
'Frame2.Visible = CBool(Check1.Value)
'-----------------------------------------------
txName.Enabled = CBool(Check1.Value)
If txName.Enabled = False Then txName.BackColor = &H8000000F Else txName.BackColor = vbWhite
End Sub

Private Sub cmBrowse_Click()
Dim P  As String
P = BrowsePath(txPath.Text)
If P <> "" Then txPath.Text = P
End Sub

Private Sub cmStart_Click()
sFolderName = txPath.Text
If PathExists(txPath.Text) = False Then
    If GetDriveType(Left(txPath.Text, 3)) <> 5 Then
        If CreateDir() = False Then
            Exit Sub
        End If
    Else
        MsgBox "Cannot Install to CD-ROM." & vbNewLine & "CDFS Not supported yet.", vbExclamation, "CDFS"
            Exit Sub
    End If
End If
Me.Hide
frmInst.Show vbModal, frmBack
End Sub

Private Sub cmX_Click()
Unload Me
End Sub

Private Sub Form_Load()
SetFonts Me
Check1.Value = ReadINIFile("Setup", "Shortcut", 0)
txName.Text = Trim(ReadINIFile("Setup", "ShortcutName", ""))
sShortCutFile = ReadINIFile("Setup", "ShortcutFile", "")
txPath.Text = Trim(ReadINIFile("Setup", "DefaultPath", ""))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If LastRites = vbYes Then
EndSetup
Else
Cancel = True
End If
End Sub

Private Sub Label6_Click()
'Easter egg
ShellExecute Me.hwnd, "open", WindowsDir & "\notepad.exe", App.Path & "\setup.inf", App.Path, 10
End Sub

Private Sub Label7_Click()
cmStart_Click
End Sub

Private Sub Label8_Click()
On Error Resume Next
txName.SetFocus
End Sub

Private Sub Label9_Click()
cmBrowse_Click
End Sub

Function CreateDir() As Boolean
If MsgBox("The path does not exist." & vbNewLine & "Do you want to create it?", vbYesNo + vbInformation, "Path") = vbYes Then
MkDir txPath.Text
CreateDir = True
Else
CreateDir = False
End If
End Function
