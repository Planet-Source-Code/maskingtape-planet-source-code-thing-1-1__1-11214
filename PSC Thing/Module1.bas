Attribute VB_Name = "Module1"
Global WasDelete As Boolean
Global tabvb As Integer
Global tabc As Integer
Global tabj As Integer
Global tabp As Integer
Global tabasp As Integer
Global tabdelphi As Integer
Global tabsql As Integer

Public Type iSense
    sOut As String * 50
End Type

Public nid As NOTIFYICONDATA
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
    End Type
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    Public Const WM_MOUSEMOVE = &H200

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Sub InitializeTrayIcon()

      With nid
        .cbSize = Len(nid)
        .hwnd = Form1.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .szTip = "Planet Source Code Thing" & vbNullChar
        .hIcon = Form1.Icon
    End With
    
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Public Function IntelliSense(tBox As TextBox, AddRecord As Boolean) As String
    Dim iChannel As Integer, iActive As Integer, iLength As Integer, i As Integer
    Dim iFile As String
    Dim iSense As iSense
    Dim Done As Boolean
    
    iFile = App.Path & "\history.dat"
    iLength = Len(iSense)
    iChannel = FreeFile
    Open iFile For Random As iChannel Len = iLength
    Close iChannel
    
    iActive = FileLen(iFile) / iLength
    iChannel = FreeFile
    Open iFile For Random As iChannel Len = iLength
        If AddRecord Then
            iSense.sOut = tBox.Text
            Put iChannel, iActive + 1, iSense
        Else
            Do While Not EOF(iChannel) And Done = False
                i = i + 1
                Get iChannel, i, iSense
                If tBox.Text = Mid(RTrim(iSense.sOut), 1, Len(tBox.Text)) Then
                    IntelliSense = RTrim(iSense.sOut)
                End If
            Loop
        End If
    Close iChannel
End Function

Public Sub iSenseChange(tBox As TextBox)
    Dim iStart As Integer
    Dim iSense As String
    
    iStart = tBox.SelStart
    iSense = IntelliSense(tBox, False)
    If iSense <> "" And Not WasDelete Then
        tBox.Text = iSense
        tBox.SelStart = iStart
        tBox.SelLength = Len(tBox.Text) - iStart
    End If
End Sub

Public Sub iSenseKeyPress(tBox As TextBox, KeyAscii As Integer)
    If KeyAscii = 13 And tBox.Text <> "" Then
        IntelliSense tBox, True
    ElseIf KeyAscii = 8 Then
        WasDelete = True
    Else
        WasDelete = False
    End If
End Sub

Public Sub savecfg()
If Form1.Timer1.Enabled = True Then temp1 = "true" Else temp1 = "false"

SaveSetting "PSCT", "Settings", "Left", Form1.Left
SaveSetting "PSCT", "Settings", "Top", Form1.Top
SaveSetting "PSCT", "Settings", "AutoR", temp1
SaveSetting "PSCT", "Settings", "FirstRun", "1"

End Sub

Public Sub loadcfg()

If GetSetting("PSCT", "Settings", "AutoR") = "true" Then Form1.Timer1.Enabled = True: Form1.ar.Checked = True
If GetSetting("PSCT", "Settings", "AutoR") = "false" Then Form1.Timer1.Enabled = False: Form1.ar.Checked = False

Form1.Left = GetSetting("PSCT", "Settings", "Left", "0")
Form1.Top = GetSetting("PSCT", "Settings", "Top", "0")
Form1.SSTab1.TabVisible(0) = GetSetting("PSCT", "Settings", "VB", "1")
Form1.SSTab1.TabVisible(1) = GetSetting("PSCT", "Settings", "C++", "1")
Form1.SSTab1.TabVisible(2) = GetSetting("PSCT", "Settings", "Java", "1")
Form1.SSTab1.TabVisible(3) = GetSetting("PSCT", "Settings", "Perl", "1")
Form1.SSTab1.TabVisible(4) = GetSetting("PSCT", "Settings", "ASP", "1")
Form1.SSTab1.TabVisible(5) = GetSetting("PSCT", "Settings", "Delphi", "1")
Form1.SSTab1.TabVisible(6) = GetSetting("PSCT", "Settings", "SQL", "1")

End Sub
