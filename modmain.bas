Attribute VB_Name = "mdSetup"
'-----------------------------------------
'Setup Demonstration By Sushant Pandurangi
'Installation Magic Utility - iMagic32.vbp
'sushant@phreaker.net / http://sushantshome.tripod.com
'You are free to use this sourcecode in your
'projects but if you appreciate the work put into
'this please include my name in the credits.
'Thanks.
'-----------------------------------------
Option Explicit
Public ReturnedExit As VbMsgBoxResult
'-----------------------------------------
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Public bPath As String 'Returned from BrowsePath()
Public sFileContents As String 'File contents of 'setup.inf'
Public sFolderName As String 'Where to install the files
Public sShortCutName As String 'Shortcut in Start menu
Public sShortCutFile As String 'Target for the shortcut
Public IfReadmeShown As Boolean 'If readme accepted
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Function BrowsePath(Optional InitialPath As String) As String
On Error Resume Next
Load frmBrowse
frmBrowse.Drive1.Drive = Left(InitialPath, 3)
frmBrowse.Dir1.Path = InitialPath
frmBrowse.Show vbModal, frmBack
BrowsePath = bPath
End Function

Sub UpdateStatus(Pic As Object, ByVal sngPercent As Single, Optional ByVal fBorderCase As Boolean = False)
On Error Resume Next
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer
    'For this to work well, we need a white background and any color foreground (blue)
    Const colBackground = &HFFFFFF ' white
    Const colForeground = 0 'Black
    Pic.ForeColor = colForeground
    Pic.BackColor = colBackground
    'Format percentage and get attributes of text
    Dim intPercent
    intPercent = sngPercent ' Int(100 * sngPercent + 0.5)
    'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    'prevents, for instance, the status bar from reaching 100% until we are entirely done.
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    strPercent = Format$(intPercent) & "%"
    intWidth = Pic.TextWidth(strPercent)
    intHeight = Pic.TextHeight(strPercent)
    'Now set intX and intY to the starting location for printing the percentage
    intX = Pic.Width / 2 - intWidth / 2
    intY = (Pic.Height / 2 - intHeight / 2) - 20
    'Need to draw a filled box with the pics background color to wipe out previous
    'percentage display (if any)
    Pic.DrawMode = 13 ' Copy Pen
    Pic.Line (intX, intY)-Step(intWidth, intHeight), Pic.BackColor, BF
    'Back to the center print position and print the text
    Pic.CurrentX = intX
    Pic.CurrentY = intY
    Pic.Print strPercent
    'Now fill in the box with the ribbon color to the desired percentage
    'If percentage is 0, fill the whole box with the background color to clear it
    'Use the "Not XOR" pen so that we change the color of the text to white
    'wherever we touch it, and change the color of the background to blue
    'wherever we touch it.
    Pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
        Pic.Line (0, 0)-(Pic.Width * sngPercent, Pic.Height), Pic.ForeColor, BF
    Else
        Pic.Line (0, 0)-(Pic.Width, Pic.Height), Pic.BackColor, BF
    End If
    Pic.Refresh
End Sub

Sub InstallFiles()
ChDir App.Path
'This procedure will read the INF file and begin installing the listed files.
Dim i As Integer, strS1 As String, strS2 As String, strS3 As String, strD1 As String, strD2 As String, strD3 As String
    For i = 1 To CInt(ReadINIFile("Setup", "TotalFiles"))
    strS1 = Replace(ReadINIFile("File" & i, "Source"), "<WinPath>", WindowsDir)
    strS2 = Replace(strS1, "<SysPath>", SystemDir)
    strS3 = Replace(strS2, "<AppPath>", App.Path)
    strS3 = Replace(strS3, "\\", "\")
    strD1 = Replace(ReadINIFile("File" & i, "Destination"), "<WinPath>", WindowsDir)
    strD2 = Replace(strD1, "<SysPath>", SystemDir)
    strD3 = Replace(strD2, "<AppPath>", sFolderName)
    strD3 = Replace(strD3, "\\", "\")
        DoEvents
            FileCopy strS3, strD3
        DoEvents
    Next i
FileCopy App.Path & "\" & App.EXEName & ".exe", Replace(sFolderName & "\" & "remove.exe", "\\", "\")
End Sub

Public Function ReadINIFile(Section As String, Key As String, Optional sDefault As String)
    Dim sRet As String
    sRet = String(255, Chr(0))
    ' Get data from INI file
    ReadINIFile = Left(sRet, GetPrivateProfileString(Section, Key, sDefault, sRet, Len(sRet), App.Path & "\setup.inf"))
End Function

Public Sub WriteINIFile(Section As String, Key As String, Value As String, Filename As String)
    WritePrivateProfileString Section, Key, Value, Filename
End Sub

Sub SetFonts(pForm As Form)
On Error Resume Next
    Dim pControl As Control
        For Each pControl In pForm.Controls
            If IsValidFont(Trim(ReadINIFile("Setup", "TextFont", "Tahoma"))) = True Then
                pControl.Font.Name = Trim(ReadINIFile("Setup", "TextFont", "Tahoma"))
            Else
                pControl.Font.Name = "MS Sans Serif"
            End If
        Next pControl
End Sub

Function WindowsDir() As String
    Dim strBuf As String
    strBuf = Space$(255)
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    If GetWindowsDirectory(strBuf, 255) > 0 Then
        WindowsDir = Left(strBuf, GetWindowsDirectory(strBuf, 255))
    Else
        WindowsDir = vbNullString
    End If
End Function

Function SystemDir() As String
    Dim strBuf As String
    strBuf = Space$(255)
    'Get the system directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    If GetSystemDirectory(strBuf, 255) > 0 Then
        'strBuf = StripTerminator(strBuf)
        'AddDirSep strBuf
       SystemDir = Left(strBuf, GetSystemDirectory(strBuf, 255))
    Else
        SystemDir = vbNullString
    End If
End Function

Function IsValidFont(Font As String) As Boolean
Dim i As Integer
For i = 0 To Screen.FontCount - 1
If Screen.Fonts(i) = Font Then
IsValidFont = True
Exit Function
Else
IsValidFont = False
End If
Next i
End Function

Function ShowReadme() As Boolean
If ReadINIFile("Setup", "Readme", "") = "" Then ShowReadme = True: Exit Function
frmRead.Show vbModal
ShowReadme = IfReadmeShown
End Function

Sub Shortcut(sPath As String, sName As String, sTarget As String, Optional sArguments As String)
Dim lReturN As Long
lReturN = fCreateShellLink(sPath, sName, sTarget, sArguments)
End Sub

Sub EndSetup(Optional Done As Boolean = False)
Dim Msg As String
If Done = True Then Msg = "" Else Msg = vbNewLine & vbNewLine & "Installation was cancelled."
If Msg <> "" Then MsgOK "Installation has not been completed." & " Thank you for using setup. You can run setup at a later time to complete the Installation." & Msg, vbInformation, "Done"
End
End Sub

Sub UnInstLog(pFileName As String, pNumber As Integer)
WriteINIFile "Uninstall", "Title", ReadINIFile("Setup", "Title", ""), sFolderName & "\uninst.isu"
WriteINIFile "Uninstall", "TotalFiles", ReadINIFile("Setup", "TotalFiles", 0), sFolderName & "\uninst.isu"
WriteINIFile "Uninstall", "File" & pNumber, pFileName, sFolderName & "\uninst.isu"
End Sub

Function PathExists(PathName As String) As Boolean
PathExists = CBool(PathFileExists(PathName))
End Function

Sub MsgOK(Message As String, Optional Dummy As VbMsgBoxStyle, Optional Title As String = "Finished")
Load frmOK
frmOK.Label1.Caption = Message
frmOK.Show vbModal, frmBack
End Sub

Function LastRites() As VbMsgBoxResult
Load frmEnd
frmEnd.Show vbModal
LastRites = ReturnedExit
End Function


Sub GradeForm(pObject As Object, Optional Colour As Integer = 2, Optional Orientation As Integer = 0, Optional Range As Integer = 700)
pObject.AutoRedraw = True
Dim intY As Integer, sColour
pObject.Scale (0, 0)-(Range, Range)
For intY = 0 To Range
Select Case Colour
Case 0
sColour = RGB(CInt((intY / Range) * 255), 0, 0)
Case 1
sColour = RGB(0, CInt((intY / Range) * 255), 0)
Case 2
sColour = RGB(0, 0, CInt((intY / Range) * 255))
Case 3
sColour = RGB(0, 128, CInt((intY / Range) * 255))
Case 4
sColour = RGB(128, 0, CInt((intY / Range) * 255))
Case 5
sColour = RGB(CInt((intY / Range) * 255), 0, 128)
Case 6
sColour = RGB(CInt((intY / Range) * 255), 128, 0)
Case 7
sColour = RGB(128, CInt((intY / Range) * 255), 0)
Case 8
sColour = RGB(0, CInt((intY / Range) * 255), 128)
Case 9
sColour = RGB(0, CInt((intY / Range) * 255), CInt((intY / Range) * 255))
Case 10
sColour = RGB(CInt((intY / Range) * 255), CInt((intY / Range) * 255), 0)
Case 11
sColour = RGB(CInt((intY / Range) * 255), 0, CInt((intY / Range) * 255))
End Select
If Orientation = 0 Then
pObject.Line (0, intY)-(Range, intY), sColour
Else
pObject.Line (intY, 0)-(intY, Range), sColour
End If
Next intY
End Sub
