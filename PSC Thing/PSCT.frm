VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Planet Source Code Thing!"
   ClientHeight    =   6030
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   2775
   ControlBox      =   0   'False
   Icon            =   "PSCT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   5775
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "Version 1.1"
            TextSave        =   "Version 1.1"
            Object.ToolTipText     =   "Version Number "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   442
            MinWidth        =   442
            Text            =   "_"
            TextSave        =   "_"
            Object.ToolTipText     =   "Minimize"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   442
            MinWidth        =   442
            Text            =   "X"
            TextSave        =   "X"
            Object.ToolTipText     =   "Close"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   600
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Planet Source Code - VB"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   2775
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         ToolTipText     =   "Select Other Search Engines"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Go"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7858
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "VB"
      TabPicture(0)   =   "PSCT.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "WebBrowser1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ProgressBar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "C++"
      TabPicture(1)   =   "PSCT.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ProgressBar2"
      Tab(1).Control(1)=   "WebBrowser2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Java"
      TabPicture(2)   =   "PSCT.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ProgressBar3"
      Tab(2).Control(1)=   "WebBrowser3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Perl"
      TabPicture(3)   =   "PSCT.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ProgressBar4"
      Tab(3).Control(1)=   "WebBrowser4"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "ASP"
      TabPicture(4)   =   "PSCT.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ProgressBar5"
      Tab(4).Control(1)=   "WebBrowser5"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Delphi"
      TabPicture(5)   =   "PSCT.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ProgressBar6"
      Tab(5).Control(1)=   "WebBrowser6"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "SQL"
      TabPicture(6)   =   "PSCT.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ProgressBar7"
      Tab(6).Control(1)=   "WebBrowser7"
      Tab(6).ControlCount=   2
      Begin MSComctlLib.ProgressBar ProgressBar7 
         Height          =   735
         Left            =   -72480
         TabIndex        =   18
         Top             =   3720
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar6 
         Height          =   735
         Left            =   -72480
         TabIndex        =   17
         Top             =   3720
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar5 
         Height          =   735
         Left            =   -72480
         TabIndex        =   16
         Top             =   3720
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar4 
         Height          =   735
         Left            =   -72480
         TabIndex        =   15
         Top             =   3720
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   735
         Left            =   -72480
         TabIndex        =   14
         Top             =   3720
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   735
         Left            =   -72480
         TabIndex        =   13
         Top             =   3720
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   735
         Left            =   2520
         TabIndex        =   12
         Top             =   3720
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser2 
         Height          =   4400
         Left            =   -75000
         TabIndex        =   5
         Top             =   0
         Width           =   2415
         ExtentX         =   4260
         ExtentY         =   7761
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   4400
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2415
         ExtentX         =   4260
         ExtentY         =   7761
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser3 
         Height          =   4400
         Left            =   -75000
         TabIndex        =   7
         Top             =   0
         Width           =   2415
         ExtentX         =   4260
         ExtentY         =   7761
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser4 
         Height          =   4400
         Left            =   -75000
         TabIndex        =   8
         Top             =   0
         Width           =   2415
         ExtentX         =   4260
         ExtentY         =   7761
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser5 
         Height          =   4400
         Left            =   -75000
         TabIndex        =   9
         Top             =   0
         Width           =   2415
         ExtentX         =   4260
         ExtentY         =   7761
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser6 
         Height          =   4400
         Left            =   -75000
         TabIndex        =   10
         Top             =   0
         Width           =   2415
         ExtentX         =   4260
         ExtentY         =   7761
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser7 
         Height          =   4400
         Left            =   -75000
         TabIndex        =   11
         Top             =   0
         Width           =   2415
         ExtentX         =   4260
         ExtentY         =   7761
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Line Line3 
      X1              =   2085
      X2              =   2085
      Y1              =   4695
      Y2              =   4965
   End
   Begin VB.Line Line2 
      X1              =   1335
      X2              =   1335
      Y1              =   4695
      Y2              =   4965
   End
   Begin VB.Image Image7 
      Height          =   300
      Left            =   2130
      Picture         =   "PSCT.frx":0506
      ToolTipText     =   "Show/Hide Search Bar"
      Top             =   4680
      Width           =   300
   End
   Begin VB.Image Image6 
      Height          =   225
      Left            =   1800
      Picture         =   "PSCT.frx":058E
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image5 
      Height          =   225
      Left            =   120
      Picture         =   "PSCT.frx":06DD
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   1305
      Picture         =   "PSCT.frx":08C1
      ToolTipText     =   "Goto..."
      Top             =   4680
      Width           =   825
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   855
      Picture         =   "PSCT.frx":0A0E
      ToolTipText     =   "Refresh All Tabs..."
      Top             =   4680
      Width           =   450
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   30
      Picture         =   "PSCT.frx":0AD2
      ToolTipText     =   "Refresh Selected Tab..."
      Top             =   4680
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "PSCT.frx":0CDD
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   3000
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu ar 
         Caption         =   "Auto Refresh"
         Checked         =   -1  'True
      End
      Begin VB.Menu clearac 
         Caption         =   "Clear Autocomplete"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu selecttabs 
         Caption         =   "Select Visible Tabs"
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
      Visible         =   0   'False
      Begin VB.Menu about2 
         Caption         =   "About"
      End
      Begin VB.Menu viewcode 
         Caption         =   "View my other code"
      End
   End
   Begin VB.Menu searche 
      Caption         =   "se"
      Visible         =   0   'False
      Begin VB.Menu psc 
         Caption         =   "Planet Source Code"
         Begin VB.Menu pscvb 
            Caption         =   "Visual Basic"
            Checked         =   -1  'True
         End
         Begin VB.Menu pscc 
            Caption         =   "C++"
         End
         Begin VB.Menu pscj 
            Caption         =   "Java"
         End
         Begin VB.Menu pscp 
            Caption         =   "Perl"
         End
         Begin VB.Menu pscasp 
            Caption         =   "ASP"
         End
         Begin VB.Menu pscdelphi 
            Caption         =   "Delphi"
         End
         Begin VB.Menu pscsql 
            Caption         =   "SQL"
         End
      End
      Begin VB.Menu wse 
         Caption         =   "Web Search Engines"
         Begin VB.Menu altavista 
            Caption         =   "Altavista"
         End
         Begin VB.Menu excite 
            Caption         =   "Excite"
         End
         Begin VB.Menu hotbot 
            Caption         =   "HotBot"
         End
         Begin VB.Menu webcrawler 
            Caption         =   "Webcrawler"
         End
         Begin VB.Menu yahoo 
            Caption         =   "Yahoo"
         End
      End
   End
   Begin VB.Menu sysmenu 
      Caption         =   "sysmenu"
      Visible         =   0   'False
      Begin VB.Menu minimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu refresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mengoto 
      Caption         =   "goto"
      Visible         =   0   'False
      Begin VB.Menu newvbcode 
         Caption         =   "Newest VB Code"
      End
      Begin VB.Menu newestccode 
         Caption         =   "Newest C++ Code"
      End
      Begin VB.Menu newestjcode 
         Caption         =   "Newest Java Code"
      End
      Begin VB.Menu newestpcode 
         Caption         =   "Newest Perl Code"
      End
      Begin VB.Menu newestaspcode 
         Caption         =   "Newest ASP Code"
      End
      Begin VB.Menu newestdcode 
         Caption         =   "Newest Delphi Code"
      End
      Begin VB.Menu newestsqlcode 
         Caption         =   "Newest SQL Code"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timervar As Integer
Private Sub about2_Click()
MsgBox "Made by MaskingTape in the year 2000. Amazing, huh?"
End Sub

Private Sub altavista_Click()
Call uncheck
altavista.Checked = True
Frame1.Caption = "Search Altavista"
End Sub

Private Sub ar_Click()
If ar.Checked = True Then
    ar.Checked = False
    Timer1.Enabled = False
ElseIf ar.Checked = False Then
    ar.Checked = True
    Timer1.Enabled = True
End If
End Sub

Private Sub clearac_Click()
filenum = FreeFile
Open App.Path & "\history.dat" For Output As #filenum
Write #filenum, ""
Close #filenum
End Sub

Private Sub Command2_Click()
searchvar = LCase(Replace(Text1, " ", "+"))

If pscvb.Checked = True Then
    Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=" & searchvar & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=1&optSort=Alphabetical"), vbHide
ElseIf excite.Checked = True Then
    Shell ("start http://search.excite.com/search.gw?search=" & searchvar), vbHide
ElseIf yahoo.Checked = True Then
    Shell ("start http://search.yahoo.com/bin/search?p=" & searchvar), vbHide
ElseIf hotbot.Checked = True Then
    Shell ("start http://hotbot.lycos.com/?MT=" & searchvar), vbHide
ElseIf webcrawler.Checked = True Then
    Shell ("start http://www.webcrawler.com/cgi-bin/WebQuery?mode=compact&maxHits=25&searchText=" & searchvar), vbHide
ElseIf altavista.Checked = True Then
    Shell ("start http://www.altavista.com/cgi-bin/query?q=" & searchvar & "&kl=XX&pg=q&Translate=on"), vbHide
ElseIf pscc.Checked = True Then
    Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=" & searchvar & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=3&optSort=Alphabetical"), vbHide
ElseIf pscj.Checked = True Then
    Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=" & searchvar & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=2&B1=Go&optSort=Alphabetical"), vbHide
ElseIf pscp.Checked = True Then
    Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=" & searchvar & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=6&B1=Go&optSort=Alphabetical"), vbHide
ElseIf pscasp.Checked = True Then
    Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=" & searchvar & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=4&B1=Go&optSort=Alphabetical"), vbHide
ElseIf pscdelphi.Checked = True Then
    Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=" & searchvar & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=7&B1=Go&optSort=Alphabetical"), vbHide
ElseIf pscsql.Checked = True Then
    Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=" & serachvar & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=5&B1=Go&optSort=Alphabetical"), vbHide
End If
End Sub

Private Sub Command3_Click()
PopupMenu searche
End Sub

Private Sub excite_Click()
Call uncheck
excite.Checked = True
Frame1.Caption = "Search Excite"
End Sub

Private Sub exit_Click()
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub Form_Load()
If GetSetting("PSCT", "Settings", "FirstRun", "0") = 0 Then
    rc = MsgBox("Since this is your first time running PSCT, would you like to configure the tabs?", vbYesNo, "First Run")
    If rc = vbYes Then Form2.Show
Else
End If

Call loadcfg

Me.Width = 2865
Me.Height = 5715
InitializeTrayIcon

WebBrowser1.Navigate "http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp?lngWId=1"
WebBrowser2.Navigate "http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp?lngWId=3"
WebBrowser3.Navigate "http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp?lngWId=2"
WebBrowser4.Navigate "http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp?lngWId=6"
WebBrowser5.Navigate "http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp?lngWId=4"
WebBrowser6.Navigate "http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp?lngWId=7"
WebBrowser7.Navigate "http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp?lngWId=5"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
Dim msg As Long

If Me.ScaleMode = vbPixels Then
    msg = X
Else
    msg = X / Screen.TwipsPerPixelX
End If
    
Select Case msg
    Case 517
    Me.PopupMenu sysmenu
    Case 514
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    minimize.Caption = "Minimize"
End Select

End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid
Call savecfg
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
Call savecfg
End
End Sub

Private Sub hotbot_Click()
Call uncheck
hotbot.Checked = True
Frame1.Caption = "Search HotBot"
End Sub

Private Sub Image2_Click()
If SSTab1.Tab = 0 Then WebBrowser1.refresh
If SSTab1.Tab = 1 Then WebBrowser2.refresh
If SSTab1.Tab = 2 Then WebBrowser3.refresh
If SSTab1.Tab = 3 Then WebBrowser4.refresh
If SSTab1.Tab = 4 Then WebBrowser5.refresh
If SSTab1.Tab = 5 Then WebBrowser6.refresh
If SSTab1.Tab = 6 Then WebBrowser7.refresh
End Sub

Private Sub Image3_Click()
WebBrowser1.refresh
WebBrowser2.refresh
WebBrowser3.refresh
WebBrowser4.refresh
WebBrowser5.refresh
WebBrowser6.refresh
WebBrowser7.refresh

End Sub

Private Sub Image4_Click()
Me.PopupMenu mengoto

End Sub

Private Sub Image5_Click()
Me.PopupMenu options
End Sub

Private Sub Image6_Click()
Me.PopupMenu about
End Sub

Private Sub Image7_Click()
If Me.Height = 5715 Then
    Me.Height = 6405
Else
    Me.Height = 5715
End If

End Sub

Private Sub minimize_Click()
If minimize.Caption = "Minimize" Then
    Me.Hide
    minimize.Caption = "Maximize"
ElseIf minimize.Caption = "Maximize" Then
    Me.Show
    minimize.Caption = "Minimize"
End If
End Sub

Private Sub newestaspcode_Click()
Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWid=4"), vbHide
End Sub

Private Sub newestccode_Click()
Shell ("start http://planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWid=3"), vbHide
End Sub

Private Sub newestdcode_Click()
Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWid=7"), vbHide
End Sub

Private Sub newestjcode_Click()
Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWid=2"), vbHide
End Sub

Private Sub newestpcode_Click()
Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWid=6"), vbHide
End Sub

Private Sub newestsqlcode_Click()
Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWid=5"), vbHide
End Sub

Private Sub newvbcode_Click()
Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWid=1"), vbHide
End Sub

Private Sub pscasp_Click()
Call uncheck
pscasp.Checked = True
Frame1.Caption = "Search Planet Source Code - ASP"
End Sub

Private Sub pscc_Click()
Call uncheck
pscc.Checked = True
Frame1.Caption = "Search Planet Source Code - C++"
End Sub

Private Sub pscdelphi_Click()
Call uncheck
pscdelphi.Checked = True
Frame1.Caption = "Search Planet Source Code - Delphi"
End Sub

Private Sub pscj_Click()
Call uncheck
pscj.Checked = True
Frame1.Caption = "Search Planet Source Code - Java"
End Sub

Private Sub pscp_Click()
Call uncheck
pscp.Checked = True
Frame1.Caption = "Search Planet Source Code - Perl"
End Sub

Private Sub pscsql_Click()
Call uncheck
pscsql.Checked = True
Frame1.Caption = "Search Planet Source Code - SQL"
End Sub

Private Sub pscvb_Click()
Call uncheck
pscvb.Checked = True
Frame1.Caption = "Search Planet Source Code - VB"
End Sub

Private Sub refresh_Click()
Image2_Click
End Sub

Private Sub selecttabs_Click()
Form2.Show
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
If Panel.Index = 2 Then
    Me.Hide
    minimize.Caption = "Maximize"
ElseIf Panel.Index = 3 Then
    Shell_NotifyIcon NIM_DELETE, nid
    Unload Me
End If
End Sub

Private Sub Text1_Change()
iSenseChange Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
iSenseKeyPress Text1, KeyAscii

If KeyAscii = 13 Then Command2_Click

End Sub

Private Sub Timer1_Timer()
timervar = timervar + 1

If timervar = 5 Then
WebBrowser1.refresh
WebBrowser2.refresh
WebBrowser3.refresh
WebBrowser4.refresh
WebBrowser5.refresh
WebBrowser6.refresh
WebBrowser7.refresh

timervar = 0
End If
End Sub

Private Sub viewcode_Click()
Shell ("start http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=33657&strAuthorName=MaskingTape&txtMaxNumberOfEntriesPerPage=25"), vbHide
End Sub

Public Sub uncheck()
pscvb.Checked = False
excite.Checked = False
yahoo.Checked = False
hotbot.Checked = False
webcrawler.Checked = False
altavista.Checked = False
pscc.Checked = False
pscj.Checked = False
pscp.Checked = False
pscasp.Checked = False
pscdelphi.Checked = False
pscsql.Checked = False
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo errorhandler
ProgressBar1.Max = ProgressMax
ProgressBar1.Value = Progress

errorhandler:
    Select Case Err
    Case Is = 380
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    End Select
End Sub

Private Sub WebBrowser2_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo errorhandler
ProgressBar2.Max = ProgressMax
ProgressBar2.Value = Progress

errorhandler:
    Select Case Err
    Case Is = 380
    ProgressBar2.Max = 1
    ProgressBar2.Value = 0
    End Select
End Sub

Private Sub WebBrowser3_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo errorhandler
ProgressBar3.Max = ProgressMax
ProgressBar3.Value = Progress

errorhandler:
    Select Case Err
    Case Is = 380
    ProgressBar3.Max = 1
    ProgressBar3.Value = 0
    End Select
End Sub

Private Sub WebBrowser4_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo errorhandler
ProgressBar4.Max = ProgressMax
ProgressBar4.Value = Progress

errorhandler:
    Select Case Err
    Case Is = 380
    ProgressBar4.Max = 1
    ProgressBar4.Value = 0
    End Select
End Sub

Private Sub WebBrowser5_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo errorhandler
ProgressBar5.Max = ProgressMax
ProgressBar5.Value = Progress

errorhandler:
    Select Case Err
    Case Is = 380
    ProgressBar5.Max = 1
    ProgressBar5.Value = 0
    End Select
End Sub

Private Sub WebBrowser6_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo errorhandler
ProgressBar6.Max = ProgressMax
ProgressBar6.Value = Progress

errorhandler:
    Select Case Err
    Case Is = 380
    ProgressBar6.Max = 1
    ProgressBar6.Value = 0
    End Select
End Sub

Private Sub WebBrowser7_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo errorhandler
ProgressBar7.Max = ProgressMax
ProgressBar7.Value = Progress

errorhandler:
    Select Case Err
    Case Is = 380
    ProgressBar7.Max = 1
    ProgressBar7.Value = 0
    End Select
End Sub

Private Sub webcrawler_Click()
Call uncheck
webcrawler.Checked = True
Frame1.Caption = "Search Webcrawler"
End Sub

Private Sub yahoo_Click()
Call uncheck
yahoo.Checked = True
Frame1.Caption = "Search Yahoo"
End Sub
