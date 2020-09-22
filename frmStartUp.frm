VERSION 5.00
Begin VB.Form frmStartUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Death By Tank"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   ClipControls    =   0   'False
   Icon            =   "frmStartUp.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdView 
      Caption         =   "View Top 10"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Width           =   2655
   End
   Begin VB.OptionButton optEasy 
      Caption         =   "Wimp"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton optHard 
      Caption         =   "Game Over"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optMedium 
      Caption         =   "Quick Death"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1015
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Begin!"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "How Do You Want To Die?"
      Height          =   2460
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   735
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optSingle 
         Caption         =   "Single Player"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optMultiNet 
         Caption         =   "Multi-Player Internet"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
      End
      Begin VB.OptionButton optMultiKey 
         Caption         =   "Multi-Player Keyboard"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ScoreType
    Name As String
    Points As Long
End Type

Private Sub cmdStart_Click()
Unload frmScores
If optSingle.Value = True Then
    GameType = "Single"
    If optEasy.Value = True Then
        frmPlay.Timer2.Interval = 500
    ElseIf optMedium.Value = True Then
        frmPlay.Timer2.Interval = 375
    Else
        frmPlay.Timer2.Interval = 1
    End If
    frmPlay.Show
ElseIf optMultiKey.Value = True Then
    GameType = "Key"
    frmPlay.Show
ElseIf optMultiNet.Value = True Then
    GameType = "Net"
    frmPlay.Show
End If
Unload Me
End Sub

Private Sub cmdView_Click()
    frmScores.Show
End Sub

Private Sub CmdExit_Click()
    End
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
If Dir(App.Path & "\Ts.ini") = "" Then
    Dim file As String
    Dim Scores(1 To 10) As ScoreType
    Scores(1).Name = "Chris"
    Scores(1).Points = 1000000
    Scores(2).Name = "Allison"
    Scores(2).Points = 975000
    Scores(3).Name = "Jeff"
    Scores(3).Points = 932000
    Scores(4).Name = "Chad"
    Scores(4).Points = 905000
    Scores(5).Name = "Jeremy"
    Scores(5).Points = 890000
    Scores(6).Name = "Aaron"
    Scores(6).Points = 800000
    Scores(7).Name = "Deeds"
    Scores(7).Points = 755000
    Scores(8).Name = "Anna"
    Scores(8).Points = 700000
    Scores(9).Name = "Coolio"
    Scores(9).Points = 675000
    Scores(10).Name = "J. Benson"
    Scores(10).Points = 10
    file = App.Path & "\Ts.ini"
    Open file For Random As #1
    For X = 1 To 10
        Put 1, X, Scores(X)
    Next X
    Close #1
End If
End Sub


Private Sub optMultiKey_Click()
optEasy.Enabled = False
optMedium.Enabled = False
optHard.Enabled = False
End Sub

Private Sub optMultiNet_Click()
optEasy.Enabled = False
optMedium.Enabled = False
optHard.Enabled = False
End Sub

Private Sub optSingle_Click()
optEasy.Enabled = True
optMedium.Enabled = True
optHard.Enabled = True
End Sub
