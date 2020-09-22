VERSION 5.00
Begin VB.Form frmPoints 
   BackColor       =   &H00008000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Points"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2055
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6885
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   6480
   End
   Begin VB.Label lblSpeedP2 
      BackColor       =   &H00008000&
      Caption         =   "Speed:  1"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblSpeedP1 
      BackColor       =   &H00008000&
      Caption         =   "Speed:  1"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblP2Alive 
      BackColor       =   &H00008000&
      Caption         =   "Alive Time:  0:00"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblP1Alive 
      BackColor       =   &H00008000&
      Caption         =   "Alive Time:  0:00"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblP2PpM 
      BackColor       =   &H00008000&
      Caption         =   "Points/Min:  0"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblP1PpM 
      BackColor       =   &H00008000&
      Caption         =   "Points/Min:  0"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00008000&
      Caption         =   "Time-->  0:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   600
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblP2Points 
      BackColor       =   &H00008000&
      Caption         =   "Points:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblP1Points 
      BackColor       =   &H00008000&
      Caption         =   "Points:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   600
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblPlayer2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblPlayer1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub TankPicture()
    ScaleMode = 3
    Call BitBlt(frmPoints.hDC, 35, 165, frmImages.Picture1.Width - 2, frmImages.Picture1.Height - 2, frmImages.Picture1.hDC, 0, 0, vbSrcAnd)
    Call BitBlt(frmPoints.hDC, 35, 165, frmImages.Picture2.Width - 2, frmImages.Picture2.Height - 2, frmImages.Picture2.hDC, 0, 0, vbSrcPaint)
    Call BitBlt(frmPoints.hDC, 35, 385, frmImages.Picture15.Width - 2, frmImages.Picture15.Height - 2, frmImages.Picture15.hDC, 0, 0, vbSrcAnd)
    Call BitBlt(frmPoints.hDC, 35, 385, frmImages.Picture16.Width - 2, frmImages.Picture16.Height - 2, frmImages.Picture16.hDC, 0, 0, vbSrcPaint)
    ScaleMode = 1
End Sub
Private Sub BulletHandle()
        If RightKey = True Then
            BltDir = 4
            locBulletX = locTankX + 57 + Speed
            locBulletY = locTankY + (frmImages.Picture5.Height / 2) - 7
        End If
        If LeftKey = True Then
            BltDir = 3
            locBulletX = locTankX - 16 - Speed
            locBulletY = locTankY + (frmImages.Picture7.Height / 2) - 7
        End If
        If DownKey = True Then
            BltDir = 2
            locBulletX = locTankX + (frmImages.Picture1.Width / 2) - 7
            locBulletY = locTankY + 57 - Speed
        End If
        If UpKey = True Then
            BltDir = 1
            locBulletX = locTankX + (frmImages.Picture1.Width / 2) - 7
            locBulletY = locTankY - 16 - Speed
        End If
End Sub

Private Sub BulletHandle2()
        If RightKey2 = True Then
            BltDir2 = 4
            locBulletX2 = locTankX2 + 57 + Speed2
            locBulletY2 = locTankY2 + (frmImages.Picture5.Height / 2) - 7
        End If
        If LeftKey2 = True Then
            BltDir2 = 3
            locBulletX2 = locTankX2 - 16 - Speed2
            locBulletY2 = locTankY2 + (frmImages.Picture7.Height / 2) - 7
        End If
        If DownKey2 = True Then
            BltDir2 = 2
            locBulletX2 = locTankX2 + (frmImages.Picture1.Width / 2) - 7
            locBulletY2 = locTankY2 + 57 - Speed2
        End If
        If UpKey2 = True Then
            BltDir2 = 1
            locBulletX2 = locTankX2 + (frmImages.Picture1.Width / 2) - 7
            locBulletY2 = locTankY2 - 16 - Speed2
        End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Sent = True And GameType = "Net" Or GameType <> net Then
If KeyCode > 36 And KeyCode < 41 Then
    LeftKey = False
    UpKey = False
    DownKey = False
    RightKey = False
End If
If KeyCode = 70 Or KeyCode = 84 Or KeyCode = 72 Or KeyCode = 71 Then
    If GameType = "Key" Then
        LeftKey2 = False
        UpKey2 = False
        DownKey2 = False
        RightKey2 = False
    End If
End If
Select Case KeyCode
    Case 37
        LeftKey = True
    Case 38
        UpKey = True
    Case 39
        RightKey = True
    Case 40
        DownKey = True
    Case 17
        If Shoot = False Then
            Shoot = True
            Call BulletHandle
        End If
    Case 107
        If Speed < 4 Then
            Speed = Speed + 1
            frmPoints.lblSpeedP1.Caption = "Speed:  " & Speed
        End If
    Case 109
        If Speed > 1 Then
            Speed = Speed - 1
            frmPoints.lblSpeedP1.Caption = "Speed:  " & Speed
        End If
        
    Case 70
        If GameType = "Key" Then
            LeftKey2 = True
        End If
    Case 84
        If GameType = "Key" Then
            UpKey2 = True
        End If
    Case 72
        If GameType = "Key" Then
            RightKey2 = True
        End If
    Case 71
        If GameType = "Key" Then
            DownKey2 = True
        End If
    Case 16
        If Shoot2 = False And GameType = "Key" Then
            Shoot2 = True
            Call BulletHandle2
        End If
    Case 9
        If Speed2 < 4 And GameType = "Key" Then
            Speed2 = Speed2 + 1
            frmPoints.lblSpeedP2.Caption = "Speed:  " & Speed2
        End If
    Case 192
        If Speed2 > 1 And GameType = "Key" Then
            Speed2 = Speed2 - 1
            frmPoints.lblSpeedP2.Caption = "Speed:  " & Speed2
        End If
    Case 123
        Unload frmPlay
        frmStartUp.Show
    End Select
    End If
End Sub
Private Sub Form_Load()
    frmPoints.Width = Screen.Width - frmPlay.Width - 100
    frmPoints.Height = frmPlay.Height
    Move frmPlay.Left - frmPoints.Width, frmPlay.Top
    Line1.X1 = 0
    Line1.X2 = frmPoints.Width
    Line1.Y1 = frmPoints.Height / 2
    Line1.Y2 = frmPoints.Height / 2
    lblPlayer1.Top = 500
    Line2.X1 = 0
    Line2.X2 = frmPoints.Width
    Line2.Y1 = lblPlayer1.Top - 60
    Line2.Y2 = lblPlayer1.Top - 60
    lblPlayer2.Top = frmPoints.Height / 2 + 60
    lblP1Points.Top = lblPlayer1.Top + 400
    lblP2Points.Top = lblPlayer2.Top + 400
    lblTime.Top = 50
    lblTime.Left = (frmPoints.Width / 2) - (lblTime.Width / 2)
    lblP1Points.Caption = "Points: " & ScoreP1
    lblP2Points.Caption = "Points: " & ScoreP2
    lblP1PpM.Left = lblPlayer1.Left
    lblP2PpM.Left = lblPlayer2.Left
    lblP1PpM.Top = lblP1Points.Top + 400
    lblP2PpM.Top = lblP2Points.Top + 400
    lblP1Alive.Left = lblP1PpM.Left
    lblP2Alive.Left = lblP2PpM.Left
    lblP1Alive.Top = lblP1PpM.Top + 400
    lblP2Alive.Top = lblP2PpM.Top + 400
    lblSpeedP1.Top = lblP1Alive.Top + 400
    lblSpeedP2.Top = lblP2Alive.Top + 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ScoreP1 <> 25 And ScoreP1 <> 25 And ClosePlay = False Then
        Cancel = True
    End If
End Sub

Private Sub Timer1_Timer()
    Dim Minutes As String
    Dim AliveP1Minutes As String
    Dim AliveP2Minutes As String
    CurrentTime = CurrentTime + 1
    AliveP1 = AliveP1 + 1
    AliveP2 = AliveP2 + 1
    Minutes = CStr(CurrentTime / 60) & "."
    AliveP1Minutes = CStr(AliveP1 / 60) & "."
    AliveP2Minutes = CStr(AliveP2 / 60) & "."
    If CurrentTime < 10 Then
        lblTime.Caption = "Time-->  0:0" & CurrentTime
    Else
        If CurrentTime - (CInt(Mid((CurrentTime / 60), 1, InStr(1, Minutes, "."))) * 60) < 10 Then
            lblTime.Caption = "Time-->  " & CInt(Mid((CurrentTime / 60), 1, InStr(1, Minutes, "."))) & ":0" & CurrentTime - (CInt(Mid((CurrentTime / 60), 1, InStr(1, Minutes, "."))) * 60)
        Else
            lblTime.Caption = "Time-->  " & CInt(Mid((CurrentTime / 60), 1, InStr(1, Minutes, "."))) & ":" & CurrentTime - (CInt(Mid((CurrentTime / 60), 1, InStr(1, Minutes, "."))) * 60)
        End If
    End If
    If AliveP1 < 10 Then
        lblP1Alive.Caption = "Alive Time:  0:0" & AliveP1
    Else
        If AliveP1 - (CInt(Mid((AliveP1 / 60), 1, InStr(1, AliveP1Minutes, "."))) * 60) < 10 Then
            lblP1Alive.Caption = "Alive Time:  " & CInt(Mid((AliveP1 / 60), 1, InStr(1, AliveP1Minutes, "."))) & ":0" & AliveP1 - (CInt(Mid((AliveP1 / 60), 1, InStr(1, AliveP1Minutes, "."))) * 60)
        Else
            lblP1Alive.Caption = "Alive Time:  " & CInt(Mid((AliveP1 / 60), 1, InStr(1, AliveP1Minutes, "."))) & ":" & AliveP1 - (CInt(Mid((AliveP1 / 60), 1, InStr(1, AliveP1Minutes, "."))) * 60)
        End If
    End If
    If AliveP2 < 10 Then
        lblP2Alive.Caption = "Alive Time:  0:0" & AliveP2
    Else
        If AliveP2 - (CInt(Mid((AliveP2 / 60), 1, InStr(1, AliveP2Minutes, "."))) * 60) < 10 Then
            lblP2Alive.Caption = "Alive Time:  " & CInt(Mid((AliveP2 / 60), 1, InStr(1, AliveP2Minutes, "."))) & ":0" & AliveP2 - (CInt(Mid((AliveP2 / 60), 1, InStr(1, AliveP2Minutes, "."))) * 60)
        Else
            lblP2Alive.Caption = "Alive Time:  " & CInt(Mid((AliveP2 / 60), 1, InStr(1, AliveP2Minutes, "."))) & ":" & AliveP2 - (CInt(Mid((AliveP2 / 60), 1, InStr(1, AliveP2Minutes, "."))) * 60)
        End If
    End If
    frmPoints.lblSpeedP1.Caption = "Speed:  " & Speed
    frmPoints.lblSpeedP2.Caption = "Speed:  " & Speed2
    lblP1PpM.Caption = "Points/Min:  " & Round(ScoreP1 / (CurrentTime / 60))
    lblP2PpM.Caption = "Points/Min:  " & Round(ScoreP2 / (CurrentTime / 60))
    Call TankPicture
End Sub
