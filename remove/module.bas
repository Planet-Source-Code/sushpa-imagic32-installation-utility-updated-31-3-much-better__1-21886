Attribute VB_Name = "mdGeneral"
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadINIFile(Section As String, Key As String, Optional sDefault As String)
    Dim sRet As String
    sRet = String(255, Chr(0))
    ' Get data from INI file
    ReadINIFile = Left(sRet, GetPrivateProfileString(Section, Key, sDefault, sRet, Len(sRet), App.Path & "\uninst.isu"))
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

Sub UpdateStatus(Pic As Object, ByVal sngPercent As Integer, Optional ByVal fBorderCase As Boolean = False)
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
    Dim intPercent As Integer
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

