VERSION 5.00
Begin VB.Form frmScores 
   BackColor       =   &H00800080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Top Ten Scores"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2670
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000E7EDA&
   Icon            =   "frmScores.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblTopScores 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000E7EDA&
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ScoreType
    Name As String
    Points As Long
End Type
Dim Score(1 To 10) As ScoreType

Private Sub cmdLeave_Click()
    Unload frmScores
End Sub
Private Sub PrintToScreen()
For Y = 0 To 9
    For z = 1 To 14 - Len(Score(Y + 1).Name)
        Spaces = Spaces & " "
    Next z
    lblTopScores.Caption = lblTopScores.Caption & Score(Y + 1).Name & Spaces & Score(Y + 1).Points & Chr(10)
    Spaces = ""
Next Y
End Sub
Private Sub Form_Load()
Dim file As String
Dim Spaces As String
On Error Resume Next
file = App.Path & "\Ts.ini"
Open file For Random As #1
For w = 1 To 10
    Get 1, w, Score(w)
Next w
Close #1
Call PrintToScreen
End Sub
