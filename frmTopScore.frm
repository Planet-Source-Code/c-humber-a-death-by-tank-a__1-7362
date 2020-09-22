VERSION 5.00
Begin VB.Form frmTopScore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Top Score - Count"
   ClientHeight    =   3195
   ClientLeft      =   3675
   ClientTop       =   2760
   ClientWidth     =   4095
   Icon            =   "frmTopScore.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3555
      Top             =   90
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1710
      MaxLength       =   12
      TabIndex        =   13
      Top             =   90
      Width           =   1305
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Top 10"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Enter Your First Name:"
      Height          =   255
      Left            =   75
      TabIndex        =   12
      Top             =   120
      Width           =   1620
   End
   Begin VB.Label lblTotalCap 
      Alignment       =   1  'Right Justify
      Caption         =   "TOTAL"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   2280
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   2640
      Y1              =   720
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   3960
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblGreatAlive 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblTimeLeft 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblWinBonus 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblPDiff 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblWinBonusCap 
      Caption         =   "Win bonus                           (1000)"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblTimeLeftCap 
      Caption         =   "Time under 5 mins.       (1s x 2500)"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblGreatAliveCap 
      Caption         =   "Greatest Alive Time      (1s x 5000)"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblPDiffCap 
      Caption         =   "Points Diference              (x10000)"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmTopScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ScoreType
    Name As String
    Points As Long
End Type
Dim ttlUnder As Integer
Dim topScore As Long
Dim timerCounter As Integer
Dim Scores(1 To 10) As ScoreType
Private Sub CheckTopScore()
Dim file As String
Dim PlayerScore As Long
    file = App.Path & "\" & "Ts.ini"
    Open file For Random As #1
    For X = 1 To 10
          Get 1, X, Scores(X)
    Next X
    Close #1
    If topScore >= Scores(10).Points Then
        Call sndPlaySound(App.Path & "\Sound\Win.wav", &H1)
        MsgBox "You Made The Top Score List."
        txtName.SetFocus
    Else
        MsgBox "You Did not Make it on the Top Score List"
        frmStartUp.Show
        Unload frmTopScore
    End If
End Sub
Private Sub Sort()
    For X = 1 To 9
        For Y = (X + 1) To 10
            If Scores(X).Points < Scores(Y).Points Then
                Dim PlaceHolder As ScoreType
                PlaceHolder = Scores(X)
                Scores(X) = Scores(Y)
                Scores(Y) = PlaceHolder
            End If
        Next Y
    Next X
    Kill App.Path & "\ts.ini"
    Open App.Path & "\ts.ini" For Random As #1
    For w = 1 To 10
        Put 1, w, Scores(w)
    Next w
    Close #1
End Sub

Private Sub cmdMain_Click()
    Unload frmTopScore
End Sub

Private Sub cmdOK_Click()
    If topScore >= Scores(10).Points And txtName.Text <> "" Then
        Scores(10).Name = txtName.Text
        Scores(10).Points = topScore
        Call Sort
        Open App.Path & "\ts.ini" For Random As #1
            For a = 1 To 10
                Get 1, a, Scores(a)
            Next a
        Close #1
        cmdOK.Enabled = False
        cmdView.Enabled = True
        cmdMain.Enabled = True
    Else
        MsgBox "You did not enter a name, Please do so."
        txtName.SetFocus
    End If
End Sub

Private Sub cmdView_Click()
frmScores.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
timerCounter = 0
Dim Minutes As String
Move (Screen.Width - Width) / 2, ((Screen.Height - Height) / 2) + 500
lblPDiff.Caption = ScoreP1 - ScoreP2
Minutes = CStr(GreatAliveP1 / 60) & "."
ttlUnder = 300 - CurrentTime
If GreatAliveP1 < 10 Then
    lblGreatAlive.Caption = "0:0" & GreatAliveP1
Else
    If GreatAliveP1 - (CInt(Mid((GreatAliveP1 / 60), 1, InStr(1, Minutes, "."))) * 60) < 10 Then
        lblGreatAlive.Caption = CInt(Mid((GreatAliveP1 / 60), 1, InStr(1, Minutes, "."))) & ":0" & GreatAliveP1 - (CInt(Mid((GreatAliveP1 / 60), 1, InStr(1, Minutes, "."))) * 60)
    Else
        lblGreatAlive.Caption = CInt(Mid((GreatAliveP1 / 60), 1, InStr(1, Minutes, "."))) & ":" & GreatAliveP1 - (CInt(Mid((GreatAliveP1 / 60), 1, InStr(1, Minutes, "."))) * 60)
    End If
End If
Minutes = CStr(ttlUnder / 60) & "."
If ttlUnder > 0 Then
    If ttlUnder < 10 Then
        lblTimeLeft.Caption = "0:0" & ttlUnder
    Else
        If ttlUnder - (CInt(Mid((ttlUnder / 60), 1, InStr(1, Minutes, "."))) * 60) < 10 Then
            lblTimeLeft.Caption = CInt(Mid((ttlUnder / 60), 1, InStr(1, Minutes, "."))) & ":0" & ttlUnder - (CInt(Mid((ttlUnder / 60), 1, InStr(1, Minutes, "."))) * 60)
        Else
            lblTimeLeft.Caption = CInt(Mid((ttlUnder / 60), 1, InStr(1, Minutes, "."))) & ":" & ttlUnder - (CInt(Mid((ttlUnder / 60), 1, InStr(1, Minutes, "."))) * 60)
        End If
    End If
Else
    lblTimeLeft.Caption = "> 5 mins."
End If
If ScoreP2 = 0 Then
    lblWinBonus = "1000"
Else
    lblWinBonus.Caption = "1"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmdView.Enabled = True Then
    frmStartUp.Show
Else
    Cancel = True
End If
End Sub

Private Sub Timer1_Timer()
Select Case timerCounter
    Case 0
        lblPDiff.Caption = "(" & lblPDiff.Caption & ")" & "  " & CStr(ScoreP1 * 10000)
        topScore = (ScoreP1 * 10000)
    Case 1
        lblGreatAlive.Caption = "(" & lblGreatAlive.Caption & ")" & "  " & CStr(GreatAliveP1 * 5000)
        topScore = topScore + (GreatAliveP1 * 5000)
    Case 2
        If (300 - CurrentTime) > 0 Then
            lblTimeLeft.Caption = "(" & lblTimeLeft.Caption & ")" & "  " & CStr((300 - CurrentTime) * 2500)
            topScore = topScore + CLng((300 - CurrentTime) * 2500)
        Else
            lblTimeLeft.Caption = "0"
        End If
    Case 3
        topScore = topScore + CLng(lblWinBonus.Caption * 1000)
        lblWinBonus.Caption = "(" & lblWinBonus.Caption & ")" & "  " & CStr(lblWinBonus.Caption * 1000)
    Case 4
        lblTotal.Caption = topScore
        txtName.Enabled = True
        txtName.BackColor = &H80000005
        Call CheckTopScore
    Case 5
        Timer1.Enabled = False
End Select
timerCounter = timerCounter + 1
End Sub
