VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabs"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2220
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   2220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check7 
      Caption         =   "SQL"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   480
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Delphi"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox Check5 
      Caption         =   "ASP"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   0
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Perl"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Java"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "C++"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Visual Basic"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Value           =   1  'Checked
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = Checked Then
    Form1.SSTab1.TabVisible(0) = True
Else
    Form1.SSTab1.TabVisible(0) = False
End If

SaveSetting "PSCT", "Settings", "VB", Check1.Value
End Sub

Private Sub Check2_Click()
If Check2.Value = Checked Then
    Form1.SSTab1.TabVisible(1) = True
Else
    Form1.SSTab1.TabVisible(1) = False
End If

SaveSetting "PSCT", "Settings", "C++", Check2.Value
End Sub

Private Sub Check3_Click()
If Check3.Value = Checked Then
    Form1.SSTab1.TabVisible(2) = True
Else
    Form1.SSTab1.TabVisible(2) = False
End If

SaveSetting "PSCT", "Settings", "Java", Check3.Value
End Sub

Private Sub Check4_Click()
If Check4.Value = Checked Then
    Form1.SSTab1.TabVisible(3) = True
Else
    Form1.SSTab1.TabVisible(3) = False
End If

SaveSetting "PSCT", "Settings", "Perl", Check4.Value
End Sub

Private Sub Check5_Click()
If Check5.Value = Checked Then
    Form1.SSTab1.TabVisible(4) = True
Else
    Form1.SSTab1.TabVisible(4) = False
End If

SaveSetting "PSCT", "Settings", "ASP", Check5.Value
End Sub

Private Sub Check6_Click()
If Check6.Value = Checked Then
    Form1.SSTab1.TabVisible(5) = True
Else
    Form1.SSTab1.TabVisible(5) = False
End If

SaveSetting "PSCT", "Settings", "Delphi", Check6.Value
End Sub

Private Sub Check7_Click()
If Check7.Value = Checked Then
    Form1.SSTab1.TabVisible(6) = True
Else
    Form1.SSTab1.TabVisible(6) = False
End If

SaveSetting "PSCT", "Settings", "SQL", Form2.Check7.Value
End Sub

Private Sub Form_Load()
Check1.Value = GetSetting("PSCT", "Settings", "VB", "1")
Check2.Value = GetSetting("PSCT", "Settings", "C++", "1")
Check3.Value = GetSetting("PSCT", "Settings", "Java", "1")
Check4.Value = GetSetting("PSCT", "Settings", "Perl", "1")
Check5.Value = GetSetting("PSCT", "Settings", "ASP", "1")
Check6.Value = GetSetting("PSCT", "Settings", "Delphi", "1")
Check7.Value = GetSetting("PSCT", "Settings", "SQL", "1")
End Sub
