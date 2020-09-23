Attribute VB_Name = "Module1"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Type ptype
    UpKey As Boolean
    DownKey As Boolean
    LeftKey As Boolean
    RightKey As Boolean
    Radious As Double
    UpSpeed As Double
    Uping As Integer
    Collidable As Boolean
    AI As Boolean
    TimesOfHit As Double
    StopedUp As Boolean
    HaveToUp As Boolean
    Score As Integer
End Type

Public Type balltype
    LeftSpeed As Double
    UpSpeed As Double
    Radious As Double
    RotSpeed As Double
    Angle As Double
End Type


Public Type XYtype
    X As Double
    Y As Double
End Type

Public P1Data As ptype
Public P2Data As ptype
Public BallData As balltype
Public BallPP As XYtype
