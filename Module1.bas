Attribute VB_Name = "Module1"
Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Global GameType As String
Global locTankX, locTankX2, locBulletX, locBulletX2 As Integer
Global locTankY, locTankY2, locBulletY, locBulletY2 As Integer
Global Shoot, LeftKey, RightKey, UpKey, DownKey As Boolean
Global Shoot2, LeftKey2, RightKey2, UpKey2, DownKey2 As Boolean
Global Speed, BltDir As Integer
Global Speed2, BltDir2 As Integer
Global ScoreP1, ScoreP2 As Integer
Global CantEnter As Boolean
Global CurrentTime, AliveP2, AliveP1 As Long
Global GreatAliveP1, GreatAliveP2 As Long
Global ClosePlay As Boolean
