VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmPlay 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Death By Tank - Chris Humber"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   ForeColor       =   &H80000008&
   Icon            =   "frmPlay.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1440
      Top             =   0
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3720
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.OptionButton optServer 
      Caption         =   "Server"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton optClient 
      Caption         =   "Client"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdListenConnect 
      Caption         =   "Connect"
      Height          =   300
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "5001"
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblPort 
      Alignment       =   1  'Right Justify
      Caption         =   "Port:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim Sent As Boolean
    Dim DataSend As String
    
Private Sub AssignInfo()
DataSend = ""
If LeftKey = True Then
    DataSend = "L"
ElseIf RightKey = True Then
    DataSend = "R"
ElseIf UpKey = True Then
    DataSend = "U"
ElseIf DownKey = True Then
    DataSend = "D"
End If
DataSend = DataSend & "," & Speed
If Shoot = True Then
    DataSend = DataSend & ",S" & "," & BltDir & "," & locBulletX & "," & locBulletY
Else
    DataSend = DataSend & ",s" & "," & BltDir & "," & locBulletX & "," & locBulletY
End If
DataSend = DataSend & "," & locTankX & "," & locTankY & ","
End Sub

Private Sub SendInfo()
    AssignInfo
    Winsock1.SendData DataSend
    Sent = True
End Sub

Private Sub HandlePoints(Player1 As Boolean)
If Player1 = True Then
    ScoreP2 = ScoreP2 + 1
    frmPoints.lblP2Points.Caption = "Points: " & ScoreP2
    If ScoreP2 = 25 Then
        MsgBox "Player 2 Wins!!! " & Chr(10) & ScoreP2 & " to " & ScoreP1
        If GreatAliveP2 < AliveP2 And Timer2.Interval = 1 Then
            GreatAliveP2 = AliveP2
        End If
        Unload frmPlay
    End If
    If GreatAliveP1 < AliveP1 Then
        GreatAliveP1 = AliveP1
    End If
    AliveP1 = 0
Else
    ScoreP1 = ScoreP1 + 1
    frmPoints.lblP1Points.Caption = "Points: " & ScoreP1
    If ScoreP1 = 25 Then
        MsgBox "Player 1 Wins!!! " & Chr(10) & ScoreP1 & " to " & ScoreP2
        If GreatAliveP1 < AliveP1 Then
            GreatAliveP1 = AliveP1
        End If
        If GameType = "Single" And frmPlay.Timer2.Interval = 1 Then
            frmTopScore.Show
        End If
        Unload frmPlay
    End If
    If GreatAliveP2 < AliveP2 Then
        GreatAliveP2 = AliveP2
    End If
    AliveP2 = 0
End If
End Sub

Private Sub TankExplode(Player1 As Boolean)
Call sndPlaySound(App.Path & "\Sound\Explode.wav", &H1)
Randomize
If Player1 = True Then
    Call BitBlt(frmPlay.hDC, locBulletX2, locBulletY2, frmImages.Picture14.ScaleWidth, frmImages.Picture14.ScaleHeight, frmImages.Picture14.hDC, 0, 0, vbSrcCopy)
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmImages.Picture23.ScaleWidth, frmImages.Picture23.ScaleHeight, frmImages.Picture23.hDC, 0, 0, vbSrcCopy)
    For X = 1 To 5000000
    Next X
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmImages.Picture9.ScaleWidth, frmImages.Picture9.ScaleHeight, frmImages.Picture9.hDC, 0, 0, vbSrcCopy)
    locTankX = Int(Rnd * 605) - 13
    locTankY = Int(Rnd * 426) - 13
    locBulletX2 = -150
    locBulletY2 = -150
    If GameType = "Net" Then
        Sent = False
        Call SendInfo
    End If
    HandlePoints (True)
Else
    Call BitBlt(frmPlay.hDC, locBulletX, locBulletY, frmImages.Picture14.ScaleWidth, frmImages.Picture14.ScaleHeight, frmImages.Picture14.hDC, 0, 0, vbSrcCopy)
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmImages.Picture23.ScaleWidth, frmImages.Picture23.ScaleHeight, frmImages.Picture23.hDC, 0, 0, vbSrcCopy)
    For X = 1 To 5000000
    Next X
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmImages.Picture9.ScaleWidth, frmImages.Picture9.ScaleHeight, frmImages.Picture9.hDC, 0, 0, vbSrcCopy)
    If GameType = "Net" Then
        Sent = False
        Winsock1.SendData "Explode"
        Call SendInfo
    Else
        locTankX2 = Int(Rnd * 605) - 13
        locTankY2 = Int(Rnd * 426) - 13
        locBulletX = -150
        locBulletY = -150
    End If
    HandlePoints (False)
End If
End Sub
Private Sub CheckCollision()
    Dim Player1 As Boolean
        If locBulletX2 > locTankX And locBulletX2 < locTankX + frmImages.Picture15.Width And locBulletY2 > locTankY And locBulletY2 < locTankY + frmImages.Picture15.Height Then
            Shoot2 = False
            Player1 = True
            Call TankExplode(Player1)
        End If
        If locBulletX > locTankX2 And locBulletX < locTankX2 + frmImages.Picture1.Width And locBulletY > locTankY2 And locBulletY < locTankY2 + frmImages.Picture1.Height Then
            Shoot = False
            Player1 = False
            Call TankExplode(Player1)
        End If
End Sub
Private Sub BulletHandle()
        Call sndPlaySound(App.Path & "\Sound\Shoot.wav", &H1)
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
        Call sndPlaySound(App.Path & "\Sound\Shoot.wav", &H1)
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
        

Private Sub cmdListenConnect_Click()
On Error GoTo err:
If optClient.Value = True Then
    Winsock1.RemoteHost = txtAddress.Text
    Winsock1.RemotePort = CLng(txtPort.Text)
    Winsock1.Connect
    txtAddress.Text = "Connecting..."
Else
    Winsock1.LocalPort = CLng(txtPort.Text)
    Winsock1.Listen
    txtAddress.Text = "Waiting..."
End If
    lblAddress.Visible = False
    lblPort.Visible = False
    txtAddress.Visible = True
    txtAddress.Enabled = False
    txtPort.Visible = False
    optClient.Visible = False
    optServer.Visible = False
    cmdListenConnect.Visible = False
Exit Sub
err:
MsgBox err.Description & " - Error number: " & err.Number & vbCrLf
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
        Unload Me
        frmStartUp.Show
    End Select
    End If
End Sub
Private Sub InitializeInfo()
Randomize
Timer3.Enabled = True
ScaleMode = 3
Call SetToFalse
LeftKey = False
DownKey = False
UpKey = False
RightKey = True
LeftKey2 = True
Speed = 1
Speed2 = 1
ScoreP1 = 0
ScoreP2 = 0
AliveP1 = 0
AliveP2 = 0
GreatAliveP1 = 0
GreatAliveP2 = 0
CurrentTime = 0
ClosePlay = False
frmPlay.BackColor = RGB(0, 125, 0)
locBulletX = -150
locBulletY = -150
locBulletX2 = -150
locBulletY2 = -150
If GameType = net And optServer.Value = True Then
    locTankX = Int(Rnd * 605) - 13
    locTankX2 = Int(Rnd * 605) - 13
    locTankY = Int(Rnd * 426) - 13
    locTankY2 = Int(Rnd * 426) - 13
    Call SendInfo
ElseIf GameType = "Key" Or GameType = "Single" Then
    locTankX = Int(Rnd * 605) - 13
    locTankX2 = Int(Rnd * 605) - 13
    locTankY = Int(Rnd * 426) - 13
    locTankY2 = Int(Rnd * 426) - 13
End If
End Sub

Private Sub LoadTheForm()
Call InitializeInfo
frmImages.Picture1.Picture = LoadPicture(App.Path + "\Pictures\PosTankUp2.bmp")
frmImages.Picture2.Picture = LoadPicture(App.Path + "\Pictures\NegTankUp2.bmp")
frmImages.Picture3.Picture = LoadPicture(App.Path + "\Pictures\PosTankDn2.bmp")
frmImages.Picture4.Picture = LoadPicture(App.Path + "\Pictures\NegTankDn2.bmp")
frmImages.Picture5.Picture = LoadPicture(App.Path + "\Pictures\PosTankRt2.bmp")
frmImages.Picture6.Picture = LoadPicture(App.Path + "\Pictures\NegTankRt2.bmp")
frmImages.Picture7.Picture = LoadPicture(App.Path + "\Pictures\PosTankLt2.bmp")
frmImages.Picture8.Picture = LoadPicture(App.Path + "\Pictures\NegTankLt2.bmp")
frmImages.Picture9.Picture = LoadPicture(App.Path + "\Pictures\bg.bmp")
frmImages.Picture10.Picture = LoadPicture(App.Path + "\Pictures\PosBulletHor.bmp")
frmImages.Picture11.Picture = LoadPicture(App.Path + "\Pictures\NegBulletHor.bmp")
frmImages.Picture12.Picture = LoadPicture(App.Path + "\Pictures\PosBulletVer.bmp")
frmImages.Picture13.Picture = LoadPicture(App.Path + "\Pictures\NegBulletVer.bmp")
frmImages.Picture14.Picture = LoadPicture(App.Path + "\Pictures\bg.bmp")
frmImages.Picture15.Picture = LoadPicture(App.Path + "\Pictures\PosTankUp.bmp")
frmImages.Picture16.Picture = LoadPicture(App.Path + "\Pictures\NegTankUp.bmp")
frmImages.Picture17.Picture = LoadPicture(App.Path + "\Pictures\PosTankDn.bmp")
frmImages.Picture18.Picture = LoadPicture(App.Path + "\Pictures\NegTankDn.bmp")
frmImages.Picture19.Picture = LoadPicture(App.Path + "\Pictures\PosTankRt.bmp")
frmImages.Picture20.Picture = LoadPicture(App.Path + "\Pictures\NegTankRt.bmp")
frmImages.Picture21.Picture = LoadPicture(App.Path + "\Pictures\PosTankLt.bmp")
frmImages.Picture22.Picture = LoadPicture(App.Path + "\Pictures\NegTankLt.bmp")
frmImages.Picture23.Picture = LoadPicture(App.Path + "\Pictures\TankExplode.bmp")
If GameType = "Single" Then
    Timer2.Enabled = True
End If
frmPoints.Show
End Sub



Private Sub Form_Load()
Move (Screen.Width - Width) / 2 + 1050, (Screen.Height - Height) / 2 + 400
If GameType = "Net" Then
    Timer1.Enabled = False
    lblAddress.Visible = True
    lblPort.Visible = True
    txtAddress.Visible = True
    txtPort.Visible = True
    optClient.Visible = True
    optServer.Visible = True
    cmdListenConnect.Visible = True
    optClient.Value = True
Else
    Call LoadTheForm
End If
End Sub

Private Sub BulletMove()
    If Shoot = True Then
        Call BitBlt(frmPlay.hDC, locBulletX, locBulletY, frmImages.Picture14.ScaleWidth, frmImages.Picture14.ScaleHeight, frmImages.Picture14.hDC, 0, 0, vbSrcCopy)
        If BltDir > 2 Then
            If BltDir = 4 Then
                locBulletX = locBulletX + 15
            ElseIf BltDir = 3 Then
                locBulletX = locBulletX - 15
            End If
            Call BitBlt(frmPlay.hDC, locBulletX, locBulletY, frmImages.Picture10.Width - 2, frmImages.Picture10.Height - 2, frmImages.Picture10.hDC, 0, 0, vbSrcAnd)
            Call BitBlt(frmPlay.hDC, locBulletX, locBulletY, frmImages.Picture11.Width - 2, frmImages.Picture11.Height - 2, frmImages.Picture11.hDC, 0, 0, vbSrcPaint)
        Else
            If BltDir = 2 Then
                locBulletY = locBulletY + 15
            ElseIf BltDir = 1 Then
                locBulletY = locBulletY - 15
            End If
            Call BitBlt(frmPlay.hDC, locBulletX, locBulletY, frmImages.Picture12.Width - 2, frmImages.Picture12.Height - 2, frmImages.Picture12.hDC, 0, 0, vbSrcAnd)
            Call BitBlt(frmPlay.hDC, locBulletX, locBulletY, frmImages.Picture13.Width - 2, frmImages.Picture13.Height - 2, frmImages.Picture13.hDC, 0, 0, vbSrcPaint)
        End If
        If locBulletX > 635 Or locBulletX < -15 Or locBulletY < -15 Or locBulletY > 465 Then
            Shoot = False
            locBulletX = -150
            locBulletY = -150
        End If
    End If
    If Shoot2 = True Then
        Call BitBlt(frmPlay.hDC, locBulletX2, locBulletY2, frmImages.Picture14.ScaleWidth, frmImages.Picture14.ScaleHeight, frmImages.Picture14.hDC, 0, 0, vbSrcCopy)
        If BltDir2 > 2 Then
            If BltDir2 = 4 Then
                locBulletX2 = locBulletX2 + 15
            ElseIf BltDir2 = 3 Then
                locBulletX2 = locBulletX2 - 15
            End If
            Call BitBlt(frmPlay.hDC, locBulletX2, locBulletY2, frmImages.Picture10.Width - 2, frmImages.Picture10.Height - 2, frmImages.Picture10.hDC, 0, 0, vbSrcAnd)
            Call BitBlt(frmPlay.hDC, locBulletX2, locBulletY2, frmImages.Picture11.Width - 2, frmImages.Picture11.Height - 2, frmImages.Picture11.hDC, 0, 0, vbSrcPaint)
        Else
            If BltDir2 = 2 Then
                locBulletY2 = locBulletY2 + 15
            ElseIf BltDir2 = 1 Then
                locBulletY2 = locBulletY2 - 15
            End If
            Call BitBlt(frmPlay.hDC, locBulletX2, locBulletY2, frmImages.Picture12.Width - 2, frmImages.Picture12.Height - 2, frmImages.Picture12.hDC, 0, 0, vbSrcAnd)
            Call BitBlt(frmPlay.hDC, locBulletX2, locBulletY2, frmImages.Picture13.Width - 2, frmImages.Picture13.Height - 2, frmImages.Picture13.hDC, 0, 0, vbSrcPaint)
        End If
        If locBulletX2 > 635 Or locBulletX2 < -15 Or locBulletY2 < -15 Or locBulletY2 > 465 Then
            Shoot2 = False
            locBulletX2 = -150
            locBulletY2 = -150
        End If
    End If
    Call CheckCollision
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ClosePlay = True
    If Winsock1.State = 7 Then
        Winsock1.Close
    End If
    If GameType <> "Single" Then
        frmStartUp.Show
    Else
        If frmPlay.Timer2.Interval <> 1 Or ScoreP1 <> 25 Then
            frmStartUp.Show
        End If
    End If
    Unload frmPoints
End Sub

Private Sub optClient_Click()
txtAddress.Visible = True
lblAddress.Visible = True
cmdListenConnect.Caption = "Connect"
End Sub

Private Sub optServer_Click()
txtAddress.Visible = False
lblAddress.Visible = False
cmdListenConnect.Caption = "Listen"
End Sub

Private Sub Timer1_Timer()
    On Error GoTo err:
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmPlay.ScaleWidth, frmPlay.ScaleHeight, frmImages.Picture9.hDC, 0, 0, vbSrcCopy)
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmPlay.ScaleWidth, frmPlay.ScaleHeight, frmImages.Picture9.hDC, 0, 0, vbSrcCopy)
If LeftKey = True Then
    locTankX = locTankX - Speed
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmImages.Picture7.Width - 2, frmImages.Picture7.Height - 2, frmImages.Picture7.hDC, 0, 0, vbSrcAnd)
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmImages.Picture8.Width - 2, frmImages.Picture8.Height - 2, frmImages.Picture8.hDC, 0, 0, vbSrcPaint)
ElseIf RightKey = True Then
    locTankX = locTankX + Speed
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmImages.Picture5.Width - 2, frmImages.Picture5.Height - 2, frmImages.Picture5.hDC, 0, 0, vbSrcAnd)
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmImages.Picture6.Width - 2, frmImages.Picture6.Height - 2, frmImages.Picture6.hDC, 0, 0, vbSrcPaint)
ElseIf UpKey = True Then
    locTankY = locTankY - Speed
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmImages.Picture1.Width - 2, frmImages.Picture1.Height - 2, frmImages.Picture1.hDC, 0, 0, vbSrcAnd)
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmImages.Picture2.Width - 2, frmImages.Picture2.Height - 2, frmImages.Picture2.hDC, 0, 0, vbSrcPaint)
ElseIf DownKey = True Then
    locTankY = locTankY + Speed
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmImages.Picture3.Width - 2, frmImages.Picture3.Height - 2, frmImages.Picture3.hDC, 0, 0, vbSrcAnd)
    Call BitBlt(frmPlay.hDC, locTankX, locTankY, frmImages.Picture4.Width - 2, frmImages.Picture4.Height - 2, frmImages.Picture4.hDC, 0, 0, vbSrcPaint)
End If
If LeftKey2 = True Then
    locTankX2 = locTankX2 - Speed2
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmImages.Picture21.Width - 2, frmImages.Picture21.Height - 2, frmImages.Picture21.hDC, 0, 0, vbSrcAnd)
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmImages.Picture22.Width - 2, frmImages.Picture22.Height - 2, frmImages.Picture22.hDC, 0, 0, vbSrcPaint)
ElseIf RightKey2 = True Then
    locTankX2 = locTankX2 + Speed2
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmImages.Picture19.Width - 2, frmImages.Picture19.Height - 2, frmImages.Picture19.hDC, 0, 0, vbSrcAnd)
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmImages.Picture20.Width - 2, frmImages.Picture20.Height - 2, frmImages.Picture20.hDC, 0, 0, vbSrcPaint)
ElseIf UpKey2 = True Then
    locTankY2 = locTankY2 - Speed2
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmImages.Picture15.Width - 2, frmImages.Picture15.Height - 2, frmImages.Picture15.hDC, 0, 0, vbSrcAnd)
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmImages.Picture16.Width - 2, frmImages.Picture16.Height - 2, frmImages.Picture16.hDC, 0, 0, vbSrcPaint)
ElseIf DownKey2 = True Then
    locTankY2 = locTankY2 + Speed2
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmImages.Picture17.Width - 2, frmImages.Picture17.Height - 2, frmImages.Picture17.hDC, 0, 0, vbSrcAnd)
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmImages.Picture18.Width - 2, frmImages.Picture18.Height - 2, frmImages.Picture18.hDC, 0, 0, vbSrcPaint)
End If
If locTankX > 592 Or locTankX < -13 Or locTankY > 413 Or locTankY < -13 Then
    LeftKey = False
    UpKey = False
    DownKey = False
    RightKey = False
    If locTankX > 592 Then
        DownKey = True
        locTankX = 592
    End If
    If locTankX < -13 Then
        UpKey = True
        locTankX = -13
    End If
    If locTankY > 413 Then
        LeftKey = True
        locTankY = 413
    End If
    If locTankY < -13 Then
        RightKey = True
        locTankY = -13
    End If
End If
If locTankX2 > 592 Or locTankX2 < -13 Or locTankY2 > 413 Or locTankY2 < -13 Then
    LeftKey2 = False
    UpKey2 = False
    DownKey2 = False
    RightKey2 = False
    If locTankX2 > 592 Then
        DownKey2 = True
        locTankX2 = 591
    End If
    If locTankX2 < -13 Then
        UpKey2 = True
        locTankX2 = -12
    End If
    If locTankY2 > 413 Then
        LeftKey2 = True
        locTankY2 = 412
    End If
    If locTankY2 < -13 Then
        RightKey2 = True
        locTankY2 = -12
    End If
End If
If Shoot = True Or Shoot2 = True Then Call BulletMove
Exit Sub
err:
MsgBox "An error has occured with one of the windows API functions. Please restart your computer and leave nothing running in the background.", vbExclamation
End
End Sub
Private Sub SetToFalse()
    RightKey2 = False
    LeftKey2 = False
    UpKey2 = False
    DownKey2 = False
End Sub
Private Sub Xaxis()
        Call SetToFalse
        If locTankX > locTankX2 Then
            RightKey2 = True
            Shoot2 = True
            Call BulletHandle2
        Else
            LeftKey2 = True
            Shoot2 = True
            Call BulletHandle2
        End If
End Sub
Private Sub Yaxis()
        Call SetToFalse
        If locTankY > locTankY2 Then
            DownKey2 = True
            Shoot2 = True
            Call BulletHandle2
        Else
            UpKey2 = True
            Shoot2 = True
            Call BulletHandle2
        End If
End Sub
Private Sub CheckShootAtPlayer()
    If (Abs(locTankY - locTankY2) / Speed) + 2 > Abs(locTankX - locTankX2) / 15 And (Abs(locTankY - locTankY2) / Speed) - 2 < Abs(locTankX - locTankX2) / 15 And DownKey = True And locTankY < locTankY2 Then
        Call Xaxis
    ElseIf (Abs(locTankY - locTankY2) / Speed) + 2 > Abs(locTankX - locTankX2) / 15 And (Abs(locTankY - locTankY2) / Speed) - 2 < Abs(locTankX - locTankX2) / 15 And UpKey = True And locTankY > locTankY2 Then
        Call Xaxis
    ElseIf (Abs(locTankY - locTankY2) / 15) + 2 > Abs(locTankX - locTankX2) / Speed And (Abs(locTankY - locTankY2) / 15) - 2 < Abs(locTankX - locTankX2) / Speed And RightKey = True And locTankX < locTankX2 Then
        Call Yaxis
    ElseIf (Abs(locTankY - locTankY2) / 15) + 2 > Abs(locTankX - locTankX2) / Speed And (Abs(locTankY - locTankY2) / 15) - 2 < Abs(locTankX - locTankX2) / Speed And LeftKey = True And locTankX > locTankX2 Then
        Call Yaxis
    ElseIf locTankY + 17 > locTankY2 And locTankY - 17 < locTankY2 Then
        Call Xaxis
    ElseIf locTankX + 17 >= locTankX2 And locTankX - 17 <= locTankX2 Then
        Call Yaxis
    End If
End Sub

Private Sub DefenceHandle()
    If (Abs(locTankY - locTankY2) / Speed2) + 7 > Abs(locTankX - locTankX2) / 15 And (Abs(locTankY - locTankY2) / Speed2) - 7 < Abs(locTankX - locTankX2) / 15 And DownKey2 = True And locTankY2 < locTankY Then
        Speed2 = 4
        Call SetToFalse
        UpKey2 = True
        CantEnter = True
    ElseIf (Abs(locTankY - locTankY2) / Speed2) + 7 > Abs(locTankX - locTankX2) / 15 And (Abs(locTankY - locTankY2) / Speed2) - 7 < Abs(locTankX - locTankX2) / 15 And UpKey2 = True And locTankY2 > locTankY Then
        Speed2 = 4
        Call SetToFalse
        DownKey2 = True
        CantEnter = True
    ElseIf (Abs(locTankY - locTankY2) / 15) + 7 > Abs(locTankX - locTankX2) / Speed2 And (Abs(locTankY - locTankY2) / 15) - 7 < Abs(locTankX - locTankX2) / Speed And RightKey2 = True And locTankX2 < locTankX Then
        Speed2 = 4
        Call SetToFalse
        LeftKey2 = True
        CantEnter = True
    ElseIf (Abs(locTankY - locTankY2) / 15) + 7 > Abs(locTankX - locTankX2) / Speed2 And (Abs(locTankY - locTankY2) / 15) - 7 < Abs(locTankX - locTankX2) / Speed2 And LeftKey2 = True And locTankX2 > locTankX Then
        Speed2 = 4
        Call SetToFalse
        RightKey2 = True
        CantEnter = True
    ElseIf locTankY2 + 17 >= locTankY And locTankY2 - 17 < locTankY And Shoot2 = True Then
        Speed2 = 4
        Call SetToFalse
        If locBulletY > locTankY2 And locTankY2 < 200 Then
            DownKey2 = True
        Else
            UpKey2 = True
        End If
        CantEnter = True
    ElseIf locTankX2 + 17 >= locTankX And locTankX2 - 17 <= locTankX And Shoot2 = True Then
        Speed2 = 4
        Call SetToFalse
        If locBulletX > locTankX2 And locTankX2 > 300 Then
            LeftKey2 = True
        Else
            RightKey2 = True
        End If
        CantEnter = True
    End If
    frmPoints.lblSpeedP2.Caption = "Speed:  " & Speed2
End Sub

Private Sub Timer2_Timer()
    If Shoot = True And Timer2.Interval < 500 And CantEnter = False Then
        Call DefenceHandle
    End If
    If Shoot = False Then CantEnter = False
    If Shoot2 = True Then
        If Abs(locTankX2 - locTankX) < Abs(locTankY2 - locTankY) Then
            If locTankX2 < locTankX Then
                Call SetToFalse
                LeftKey2 = True
            Else
                Call SetToFalse
                RightKey2 = True
            End If
        Else
            If locTankY2 < locTankY Then
                Call SetToFalse
                UpKey2 = True
            Else
                Call SetToFalse
                DownKey2 = True
            End If
        End If
    End If
    If Shoot2 = False And CantEnter = False Then
        Speed2 = Speed
        frmPoints.lblSpeedP2.Caption = "Speed:  " & Speed2
        If Abs(locTankY - locTankY2) <= Abs(locTankX - locTankX2) Then
            If locTankY >= locTankY2 Then
                Call SetToFalse
                DownKey2 = True
            ElseIf locTankY < locTankY2 Then
                Call SetToFalse
                UpKey2 = True
            End If
        Else
            If locTankX >= locTankX2 Then
                Call SetToFalse
                RightKey2 = True
            ElseIf locTankX < locTankX2 Then
                Call SetToFalse
                LeftKey2 = True
            End If
        End If
        Call CheckShootAtPlayer
    End If
End Sub


Private Sub Timer3_Timer()
    If Sent = True Then
        send = False
        Call SendInfo
    End If
End Sub

Private Sub Winsock1_Close()
If Winsock1.State <> 0 Then Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
On Error GoTo err:
txtAddress.Visible = False
txtAddress.Enabled = True
Call LoadTheForm
Sent = True
Timer1.Enabled = True
Timer3.Enabled = True
Exit Sub
err:
MsgBox err.Description & " - Error number: " & err.Number & vbCrLf
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
On Error GoTo err:
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID
txtAddress.Visible = False
txtAddress.Enabled = True
Call LoadTheForm
Sent = True
Timer1.Enabled = True
Timer3.Enabled = True
Exit Sub
err:
MsgBox err.Description & " - Error number: " & err.Number & vbCrLf
End Sub
Private Sub Cover()
    Call BitBlt(frmPlay.hDC, locBulletX2, locBulletY2, frmImages.Picture14.ScaleWidth, frmImages.Picture14.ScaleHeight, frmImages.Picture14.hDC, 0, 0, vbSrcCopy)
    Call BitBlt(frmPlay.hDC, locTankX2, locTankY2, frmPlay.ScaleWidth, frmPlay.ScaleHeight, frmImages.Picture9.hDC, 0, 0, vbSrcCopy)
End Sub
Private Sub DataHandle(DataGet As String)
    Call Cover
    If DataGet = "Explode" Then
        TankExplode (True)
    Else
    Dim start, endin As Integer
    start = 1
    endin = InStr(start, DataGet, ",")
    If Mid(DataGet, start, endin - start) = "L" Then
        Call SetToFalse
        LeftKey2 = True
    ElseIf Mid(DataGet, start, endin - start) = "R" Then
        Call SetToFalse
        RightKey2 = True
    ElseIf Mid(DataGet, start, endin - start) = "U" Then
        Call SetToFalse
        UpKey2 = True
    ElseIf Mid(DataGet, start, endin - start) = "D" Then
        Call SetToFalse
        DownKey2 = True
    End If
    start = endin + 1
    endin = InStr(start, DataGet, ",")
    Speed2 = CInt(Mid(DataGet, start, endin - start))
    start = endin + 1
    endin = InStr(start, DataGet, ",")
    If Mid(DataGet, start, endin - start) = "S" Then
        Shoot2 = True
    End If
    start = endin + 1
    endin = InStr(start, DataGet, ",")
    BltDir2 = CInt(Mid(DataGet, start, endin - start))
    start = endin + 1
    endin = InStr(start, DataGet, ",")
    locBulletX2 = CInt(Mid(DataGet, start, endin - start))
    start = endin + 1
    endin = InStr(start, DataGet, ",")
    locBulletY2 = CInt(Mid(DataGet, start, endin - start))
    start = endin + 1
    endin = InStr(start, DataGet, ",")
    locTankX2 = CInt(Mid(DataGet, start, endin - start))
    start = endin + 1
    endin = Len(DataGet)
    locTankY2 = CInt(Mid(DataGet, start, endin - start))
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim DataGet As String
    On Error GoTo err:
    Winsock1.GetData DataGet
    Call DataHandle(DataGet)
    Exit Sub
err:
'MsgBox err.Description & " - Error number: " & err.Number & vbCrLf
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'MsgBox err.Description & " - Error number: " & err.Number & vbCrLf
End Sub
