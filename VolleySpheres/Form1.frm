VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   675
   ClientTop       =   795
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ballshad 
      Height          =   495
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   19
      Top             =   4920
      Width           =   735
   End
   Begin VB.PictureBox back 
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   10335
      TabIndex        =   0
      Top             =   -120
      Width           =   10400
      Begin VB.ListBox moves 
         Height          =   1425
         Left            =   4200
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.PictureBox p2shad 
         Height          =   615
         Left            =   7200
         ScaleHeight     =   555
         ScaleWidth      =   915
         TabIndex        =   18
         Top             =   6240
         Width           =   975
      End
      Begin VB.PictureBox p1shad 
         Height          =   615
         Left            =   2160
         ScaleHeight     =   555
         ScaleWidth      =   915
         TabIndex        =   17
         Top             =   6360
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   2280
         Top             =   2640
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bot Level"
         Height          =   1335
         Left            =   7920
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
         Begin VB.OptionButton Option5 
            Caption         =   "Medium"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Easy"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Difficult"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Bot Level"
         Height          =   1335
         Left            =   1440
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
         Begin VB.OptionButton Option2 
            Caption         =   "Medium"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Difficult"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Easy"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "AI"
         Height          =   255
         Left            =   7320
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "AI"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.PictureBox ballpos 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   5160
         ScaleHeight     =   615
         ScaleWidth      =   315
         TabIndex        =   5
         Top             =   0
         Width           =   315
         Begin VB.Line Line3 
            X1              =   120
            X2              =   240
            Y1              =   240
            Y2              =   120
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   120
            Y1              =   120
            Y2              =   240
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   120
            Y1              =   -120
            Y2              =   240
         End
      End
      Begin VB.PictureBox ball 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         FillColor       =   &H000000FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   2040
         ScaleHeight     =   1095
         ScaleWidth      =   1095
         TabIndex        =   4
         Top             =   4320
         Width           =   1095
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            FillStyle       =   7  'Diagonal Cross
            Height          =   855
            Left            =   0
            Shape           =   3  'Circle
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox p2 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   6600
         ScaleHeight     =   1095
         ScaleWidth      =   1095
         TabIndex        =   3
         Top             =   5880
         Width           =   1095
         Begin VB.Shape Shape3 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            Height          =   1215
            Left            =   0
            Shape           =   3  'Circle
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox p1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   1095
         TabIndex        =   2
         Top             =   5880
         Width           =   1095
         Begin VB.Shape Shape2 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Opaque
            Height          =   1215
            Left            =   0
            Shape           =   3  'Circle
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox net 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   4740
         ScaleHeight     =   3615
         ScaleWidth      =   555
         TabIndex        =   1
         Top             =   3960
         Width           =   560
      End
      Begin VB.Label score2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Score"
         Height          =   855
         Left            =   5520
         TabIndex        =   22
         Top             =   120
         Width           =   4815
      End
      Begin VB.Label Score1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Score"
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label lostmsg 
         Alignment       =   2  'Center
         Caption         =   "OOOOOUUUU!!!!!!!!!!!!!!!! P? IS A LOSER!!!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   1320
         TabIndex        =   14
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#############################################

Public ballgot As Integer
Dim ase As Integer
Dim asetoase As Integer
Public loser As Integer
Dim another As Integer
Dim p2Last As Double
Dim p1Last As Double
Dim ballleftLast As Double
Dim balltopLast As Double
'#############################################



Sub movething(ByVal obj As String, ByVal direc As String)
Select Case obj
Case "p1"
    Select Case direc
        Case "u"
            'If Form2.gameType = "mp" And Form2.Atri = "s" Then moves.AddItem "u"
            If P1Data.Uping = 0 Then P1Data.UpSpeed = -280: P1Data.Uping = 1
        Case "d"
        
        Case "l"
            If Form2.gameType = "mp" And Form2.Atri = "s" Then moves.AddItem "l"
            If p1.Left > back.Left + 75 Then
                p1.Left = p1.Left - 90
            End If
        Case "r"
            If Form2.gameType = "mp" And Form2.Atri = "s" Then moves.AddItem "r"
            If p1.Left + p1.Width < net.Left - 75 Then
                p1.Left = p1.Left + 90
            End If
        End Select
Case "p2"
    Select Case direc
        Case "u"
            'If Form2.gameType = "mp" And Form2.Atri = "c" Then moves.AddItem "u"
            If P2Data.Uping = 0 Then P2Data.UpSpeed = -280: P2Data.Uping = 1
        Case "d"
        
        Case "l"
            If Form2.gameType = "mp" And Form2.Atri = "c" Then moves.AddItem "l"
            If p2.Left > net.Left + net.Width + 75 Then
                p2.Left = p2.Left - 90
            End If
        Case "r"
            If Form2.gameType = "mp" And Form2.Atri = "c" Then moves.AddItem "r"
            If p2.Left + p2.Width < back.Width - 75 Then
                p2.Left = p2.Left + 90
            End If
        End Select
End Select
End Sub



Private Sub Check1_Click()
If Check1.Value = 1 Then
    P1Data.AI = True
Else
    P1Data.AI = False
End If
SendKeys "{tab}"
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    P2Data.AI = True
Else
    P2Data.AI = False
End If
SendKeys "{tab}"
End Sub

Public Sub Loadup()
On Error Resume Next
fixall
Dim a As Integer
'If Form2.Atri = "c" Then lanspeed = 20 Else lanspeed = 40
If Form2.Option7.Value = True Then lanspeed = Val(Form2.Option7.Tag)
If Form2.Option8.Value = True Then lanspeed = Val(Form2.Option8.Tag): serverp1speeD = Val(Form2.Option8.Tag)
If Form2.Option9.Value = True Then lanspeed = Val(Form2.Option9.Tag): serverp1speeD = Val(Form2.Option9.Tag)
If Form2.Option10.Value = True Then lanspeed = Val(Form2.Option10.Tag): serverp1speeD = Val(Form2.Option10.Tag)
If Form2.Option11.Value = True Then lanspeed = 99 - Val(Form2.Option11.Tag)

If Form2.Option13.Value = True Then gamespeed = Form2.Slider1.Max - Form2.Slider1.Value Else gamespeed = -1

'==============================
    frmmain.InitialiseGeometry
    frmmain.Render
    DoEvents
    DoEvents
    frmmain.InitialiseGeometry
    frmmain.Render
    DoEvents
    DoEvents
    frmmain.InitialiseGeometry
    frmmain.Render
    DoEvents
    DoEvents
    frmmain.InitialiseGeometry
    frmmain.Render
    DoEvents
    DoEvents
'============================
DoEvents

wee = Timer
    frmmain.InitialiseGeometry
    frmmain.Render
    DoEvents
    DoEvents
wee = Timer - wee

Do
'checkkeys
a = a + 1
If a > gamespeed Then
    Score1.Caption = P1Data.Score
    score2.Caption = P2Data.Score
    movep1datan2
    moveball
    frmmain.InitialiseGeometry
    frmmain.Render
    a = 0
    b = b + 1
    If b > lanspeed And Form2.gameType = "mp" Then
        If Form2.Atri = "c" Then
        'If this is the client
            If moves.ListCount <> 0 Or p2.Top - p2Last <> 0 Then
                Dim rrt As String
                buff = ""
                For n = 0 To moves.ListCount - 1
                    buff = buff & moves.List(n)
                Next n
                rrt = buff & ";" & Int(p2.Top)
                p2Last = p2.Top
                Form2.senddat 2, rrt
                DoEvents
                DoEvents
                DoEvents
                moves.Clear
            End If
        Else
            If ballleftLast - ball.Left <> 0 Or balltopLast - ball.Top <> 0 Then
                Dim rrt2 As String
                rrt2 = Int(ball.Left) & ";" & Int(ball.Top) & "?" & BallData.Angle
                Form2.senddat 5, rrt2
                ballleftLast = ball.Left
                balllefttop = ball.Top
            End If
            If f > serverp1speeD And (moves.ListCount <> 0 Or p1.Top - p1Last <> 0) Then
                buff = ""
                For n = 0 To moves.ListCount - 1
                    buff = buff & moves.List(n)
                Next n
                rrt2 = buff & ";" & Int(p1.Top)
                p1Last = p1.Top
                Form2.senddat 7, rrt2
                moves.Clear
                f = 0
            Else
            f = f + 1
            End If
            DoEvents
            DoEvents
            DoEvents
            DoEvents
        End If
        b = 0
    Else
        If Form2.gameType = "mp" Then b = b + 1
    End If
    
    If gamespeed = -1 Then
        zp = wee
        ert = 3000 - (5000 * (zp / 0.07))
        gamespeed = Int(ert)
    End If
    
End If
DoEvents
Loop
End Sub


Sub fixall()
BallData.LeftSpeed = 0
BallData.RotSpeed = 0
If Me.loser = 0 Then
    ball.Top = back.Height - 3135
    ball.Left = net.Left / 2
End If
P1Data.TimesOfHit = 0
P2Data.TimesOfHit = 0
BallData.LeftSpeed = -60
BallData.UpSpeed = 0
ase = 1
asetoase = 1
BallData.Radious = ball.Width / 2
P1Data.Radious = p1.Width / 2
P2Data.Radious = p2.Width / 2
P1Data.AI = False
P2Data.AI = False
P1Data.Collidable = True
P2Data.Collidable = True
lostmsg.Visible = False
If Check1.Value = 1 Then P1Data.AI = True
If Check2.Value = 1 Then P2Data.AI = True

End Sub

Sub moveball()
If (Form2.gameType = "mp" And Form2.Atri = "s") Or Form2.gameType = "sp" Then
    ase = 0
    
    Dim a As RECT
    Dim b As RECT
    Dim c As RECT
    Dim X, xx1, Y, yy1, dist As Double
    
    'check for p1  collision
    'balls data
    b.Top = ball.Top
    b.bottom = ball.Top + ball.Height
    b.Left = ball.Left
    b.Right = ball.Left + ball.Width
    
    e = GetDist(p1.Left + P1Data.Radious, p1.Top + P1Data.Radious, ball.Left + BallData.Radious, ball.Top + BallData.Radious)
    If e < BallData.Radious + P1Data.Radious Then
        If P1Data.Collidable = True And lostmsg.Visible = False Then
            If ball.Left + (ball.Width / 2) < p1.Left + p1.Width / 2 Then
                xpd = ((p1.Left + (p1.Width + 1) / 2)) - (ball.Left + (ball.Width / 2))
                BallData.LeftSpeed = ((320 * xpd) / p1.Width / 2)
                BallData.LeftSpeed = -Abs(BallData.LeftSpeed)
            Else
                xpd = ((p1.Left + (p1.Width + 1) / 2)) - (ball.Left + (ball.Width / 2))
                BallData.LeftSpeed = ((320 * xpd) / p1.Width / 2)
                BallData.LeftSpeed = Abs(BallData.LeftSpeed)
            End If
            If ball.Top + (ball.Height / 2) < p1.Top + p1.Height / 2 Then
                xpd = ((p1.Top + (p1.Height + 1) / 2)) - (ball.Top + (ball.Height / 2))
                BallData.UpSpeed = Int((450 * xpd) / p1.Height / 2)
                If BallData.UpSpeed < 220 And BallData.UpSpeed > -1 Then BallData.UpSpeed = BallData.UpSpeed + (-P1Data.UpSpeed \ 4)
                If P1Data.UpSpeed = 0 Then BallData.UpSpeed = BallData.UpSpeed + 25
                BallData.UpSpeed = -Abs(BallData.UpSpeed)
            Else
                xpd = ((p1.Top + (p1.Height + 1) / 2)) - (ball.Top + (ball.Height / 2))
                BallData.UpSpeed = Int((450 * xpd) / p1.Height / 2)
                BallData.UpSpeed = Abs(BallData.UpSpeed)
            End If
            Do
                e = GetDist(p1.Left + P1Data.Radious, p1.Top + P1Data.Radious, ball.Left + BallData.Radious, ball.Top + BallData.Radious)
                ball.Top = ball.Top + BallData.UpSpeed
                ball.Left = ball.Left + BallData.LeftSpeed
            Loop Until e > BallData.Radious + P2Data.Radious
            sndPlaySound App.Path & "\Data\dat4.png", 1
            asetoase = 0
            P1Data.TimesOfHit = P1Data.TimesOfHit + 1
            P2Data.TimesOfHit = 0
        End If
        P1Data.Collidable = False
    Else
        P1Data.Collidable = True
    End If
    
    
    
    'p2 data
    c.Top = p2.Top
    c.bottom = p2.Top + p2.Height
    c.Left = p2.Left
    c.Right = p2.Left + p2.Width
    'check for p2 collision
    e = GetDist(p2.Left + P2Data.Radious, p2.Top + P2Data.Radious, ball.Left + BallData.Radious, ball.Top + BallData.Radious)
    If e < BallData.Radious + P2Data.Radious Then
        If P2Data.Collidable = True And lostmsg.Visible = False Then
            If ball.Left + (ball.Width / 2) < p2.Left + p2.Width / 2 Then
                xpd = ((p2.Left + (p2.Width + 1) / 2)) - (ball.Left + (ball.Width / 2))
                BallData.LeftSpeed = (320 * xpd) / p2.Width / 2
                BallData.LeftSpeed = -Abs(BallData.LeftSpeed)
            Else
                xpd = ((p2.Left + (p2.Width + 1) / 2)) - (ball.Left + (ball.Width / 2))
                BallData.LeftSpeed = (320 * xpd) / p2.Width / 2
                BallData.LeftSpeed = Abs(BallData.LeftSpeed)
            End If
            If ball.Top + (ball.Height / 2) < p2.Top + p2.Height / 2 Then
                xpd = ((p2.Top + (p2.Height + 1) / 2)) - (ball.Top + (ball.Height / 2))
                BallData.UpSpeed = Int((450 * xpd) / p2.Height / 2)
                If BallData.UpSpeed < 220 And BallData.UpSpeed > -1 Then BallData.UpSpeed = BallData.UpSpeed + (-P2Data.UpSpeed \ 4)
                BallData.UpSpeed = -Abs(BallData.UpSpeed)
            Else
                xpd = ((p2.Top + (p2.Height + 1) / 2)) - (ball.Top + (ball.Height / 2))
                BallData.UpSpeed = Int((450 * xpd) / p2.Height / 2)
                BallData.UpSpeed = Abs(BallData.UpSpeed)
            End If
            Do
                e = GetDist(p2.Left + P2Data.Radious, p2.Top + P2Data.Radious, ball.Left + BallData.Radious, ball.Top + BallData.Radious)
                ball.Top = ball.Top + BallData.UpSpeed
                ball.Left = ball.Left + BallData.LeftSpeed
            Loop Until e > BallData.Radious + P2Data.Radious
            
            sndPlaySound App.Path & "\Data\dat4.png", 1
            asetoase = 0
            P2Data.TimesOfHit = P2Data.TimesOfHit + 1
            P1Data.TimesOfHit = 0
        End If
        P2Data.Collidable = False
    Else
        P2Data.Collidable = True
    End If
    
    
    If BallData.LeftSpeed < 0 Then
        Dim xr As Double
        xr = Abs(BallData.LeftSpeed) * 0.2 / 150
        BallData.RotSpeed = xr
    End If
    If BallData.LeftSpeed > 0 Then
        Dim xr2 As Double
        xr2 = Abs(BallData.LeftSpeed) * 0.2 / 150
        BallData.RotSpeed = -xr2
    End If
    
    
    'reset times of hit if the ball passed the net
    If (ball.Left > net.Left + net.Width) Then P1Data.TimesOfHit = 0
    If (ball.Left + ball.Width < net.Left) Then P2Data.TimesOfHit = 0
    
    
    'check if they hit it too many times
    If P1Data.TimesOfHit > 3 Then
        If lostmsg.Visible = False Then
            P1Data.TimesOfHit = -1000
            loserslost 1
        End If
    End If
    If P2Data.TimesOfHit > 3 Then
        If lostmsg.Visible = False Then
            P2Data.TimesOfHit = -1000
            loserslost 2
        End If
    End If
    
    
    
    
    'check if it hit on the ground
    If ball.Top + ball.Height > back.Height + 70 Then
        ball.Top = back.Height - ball.Height 'put the ball where is should be
        BallData.UpSpeed = -BallData.UpSpeed 'make the ball go the other way on the Y axis
        BallData.UpSpeed = BallData.UpSpeed + 60  'Loss of energy after a collision
        If BallData.LeftSpeed < 3 And BallData.LeftSpeed > -3 Then
            BallData.LeftSpeed = 0
        Else
            If BallData.LeftSpeed < 0 Then BallData.LeftSpeed = BallData.LeftSpeed + 4 Else BallData.LeftSpeed = BallData.LeftSpeed - 4 'Energy Losses
        End If
        'checkrules
        If lostmsg.Visible = False Then
            If ball.Left < net.Left Then yp = 1 Else yp = 2
            loserslost (yp)
        End If
    End If
    
    
    If ball.Top + ball.Height > back.Height - 10 And asetoase = 0 Then
        If BallData.UpSpeed < 21 And BallData.UpSpeed > -21 Then ase = 1
        If BallData.LeftSpeed < 1 And BallData.LeftSpeed > -1 And ase = 1 Then
            BallData.LeftSpeed = 0
        Else
            If BallData.LeftSpeed < 0 Then BallData.LeftSpeed = BallData.LeftSpeed + 2 Else BallData.LeftSpeed = BallData.LeftSpeed - 2 'Energy Losses
        End If
    End If
    
    'check if there is any hit on the sides
    If (ball.Left + ball.Width > back.Width - 70) Or (ball.Left < back.Left + 70) Then
        If (ball.Left + ball.Width > back.Width - 70) Then ball.Left = back.Width - ball.Width - (Abs(BallData.LeftSpeed) + 100)
        If (ball.Left < back.Left + 70) Then ball.Left = Abs(BallData.LeftSpeed) + 20
        BallData.LeftSpeed = -BallData.LeftSpeed
    End If
    
    'check if there is any collision with the net
    
    'nets data
    c.Top = net.Top
    c.bottom = net.Top + net.Height
    c.Left = net.Left
    c.Right = net.Left + net.Width
    
    
    
    'check for net collision
    If IntersectRect(a, b, c) <> 0 Then
        If ball.Top + ball.Height - Abs(BallData.UpSpeed) + 10 < net.Top Then
            ball.Top = net.Top - ball.Height
            BallData.UpSpeed = -(Abs(BallData.UpSpeed) / 2)
        Else
            If ball.Left + BallData.Radious <= net.Left + net.Width / 2 Then ball.Left = net.Left - ball.Width Else ball.Left = net.Left + net.Width
            BallData.LeftSpeed = -BallData.LeftSpeed
        End If
    End If
    
    
    
    'Move Jump p1
    If P1Data.Uping = 1 Then
        If P1Data.UpKey = False And P1Data.AI = False And P1Data.StopedUp = False And P1Data.UpSpeed < -100 And P1Data.HaveToUp = False Then P1Data.UpSpeed = -100: P1Data.StopedUp = True
        p1.Top = p1.Top + P1Data.UpSpeed
        P1Data.UpSpeed = P1Data.UpSpeed + Val(Form2.Slider3.Value)
        If p1.Top + p1.Height > back.Height - 10 Then
            p1.Top = back.Height - p1.Height
            P1Data.UpSpeed = 0
            P1Data.Uping = 0
            P1Data.StopedUp = False
            If P1Data.HaveToUp = True Then P1Data.UpKey = True
        End If
    End If
    
    
        'Move Jump p2
    If P2Data.Uping = 1 Then
        If P2Data.UpKey = False And P2Data.AI = False And P2Data.StopedUp = False And P2Data.UpSpeed < -100 And P2Data.HaveToUp = False Then P2Data.UpSpeed = -100: P2Data.StopedUp = True
        p2.Top = p2.Top + P2Data.UpSpeed
        P2Data.UpSpeed = P2Data.UpSpeed + Val(Form2.Slider3.Value)
        If p2.Top + p2.Height > back.Height - 10 Then
            p2.Top = back.Height - p2.Height
            P2Data.UpSpeed = 0
            P2Data.Uping = 0
            P2Data.StopedUp = False
            If P2Data.HaveToUp = True Then P2Data.UpKey = True
        End If
    End If
    
    
    If (ball.Left < Abs(BallData.LeftSpeed) + 100) Then ball.Left = Abs(BallData.LeftSpeed) + 100: BallData.LeftSpeed = -BallData.LeftSpeed
    
    
    
    If asetoase = 0 Then
        ball.Left = ball.Left + BallData.LeftSpeed
        ball.Top = ball.Top + BallData.UpSpeed
    End If
        
    If ball.Top + ball.Height + 100 < 0 Then BallData.UpSpeed = BallData.UpSpeed + 6
    If ase = 0 Then
        BallData.UpSpeed = BallData.UpSpeed + Val(Form2.Slider2.Value) + 0.2
        If lostmsg.Visible = True Then ball.Top = ball.Top + 10
    Else
        BallData.UpSpeed = 0
    End If
    
    
    ballpos.Left = ball.Left + (ball.Width / 2) - (ballpos.Width / 2)
    
    
    
Else

    If P1Data.Uping = 1 Then
        If P1Data.UpKey = False And P1Data.AI = False And P1Data.StopedUp = False And P1Data.UpSpeed < 0 And P1Data.HaveToUp = False Then P1Data.UpSpeed = -100: P1Data.StopedUp = True
        p1.Top = p1.Top + P1Data.UpSpeed
        P1Data.UpSpeed = P1Data.UpSpeed + 10
        If p1.Top + p1.Height > back.Height - 10 Then
            p1.Top = back.Height - p1.Height
            P1Data.UpSpeed = 0
            P1Data.Uping = 0
            P1Data.StopedUp = False
            If P1Data.HaveToUp = True Then P1Data.UpKey = True
        End If
    End If
    
    
    'Move Jump p2
    If P2Data.Uping = 1 Then
        If P2Data.UpKey = False And P2Data.AI = False And P2Data.StopedUp = False And P2Data.UpSpeed < 0 And P2Data.HaveToUp = False Then P2Data.UpSpeed = -100: P2Data.StopedUp = True
        p2.Top = p2.Top + P2Data.UpSpeed
        P2Data.UpSpeed = P2Data.UpSpeed + 10
        If p2.Top + p2.Height > back.Height - 10 Then
            p2.Top = back.Height - p2.Height
            P2Data.UpSpeed = 0
            P2Data.Uping = 0
            P2Data.StopedUp = False
            If P2Data.HaveToUp = True Then P2Data.UpKey = True
        End If
    End If
    
End If


'move shadow for p1
p1shad.Top = ((back.Height) - ((back.Height - p1.Top) * 1500) / 5865) - 100
p1shad.Left = (p1.Left + ((back.Height - p1.Top) * 1500) / 5865) - 100


'move shadow for p2
p2shad.Top = ((back.Height) - ((back.Height - p2.Top) * 1500) / 5865) - 100
p2shad.Left = (p2.Left + ((back.Height - p2.Top) * 1500) / 5865) - 100


'move shadow for ball
ballshad.Top = ((back.Height) - ((back.Height - ball.Top) * 1500) / 8865) - 250
ballshad.Left = (ball.Left + ((back.Height - ball.Top) * 1500) / 8865) - 100


End Sub


Sub movep1datan2()
If Form2.gameType = "sp" Then
    If Not P1Data.AI Then
    'send p1data moves
        If P1Data.UpKey = True Then movething "p1", "u"
        If P1Data.DownKey = True Then movething "p1", "d"
        If P1Data.RightKey = True Then movething "p1", "r"
        If P1Data.LeftKey = True Then movething "p1", "l"
    Else
    'move the p1 ( AI BOT )
      If lostmsg.Visible = False Then
        e = GetDist(p1.Left + P1Data.Radious, p1.Top + P1Data.Radious, ball.Left + BallData.Radious, ball.Top + BallData.Radious)
        If (ball.Left > net.Left) Or (e > (P1Data.Radious + BallData.Radious) * 6 And Option3.Value = True) Then
            If (p1.Left < (net.Left / 2 - 50)) Then movething "p1", "r": GoTo yuy
            If (p1.Left > (net.Left / 2 + 50)) Then movething "p1", "l"
        Else
            If ((ball.Left + BallData.Radious / 6) < (p1.Left + P1Data.Radious) + 160) Then
                movething "p1", "l"
                If Option3.Value = True And (p1.Left + P1Data.Radious - ball.Left - BallData.Radious > BallData.Radious + 100) Then
                    movething "p1", "l"
                End If
            Else
                movething "p1", "r"
                If Option3.Value = True And (p1.Left + P1Data.Radious - ball.Left - BallData.Radious > BallData.Radious - 100) Then
                    movething "p1", "r"
                End If
            End If
yuy:
    'And (p1.Left + p1.Width + 300 > ball.Left And (ball.Top = back.Height - 3135 And net.Left / 2))
            If ((Option2.Value = True Or Option3.Value = True Or asetoase = 1) And (lostmsg.Visible = False)) Then
                e = GetDist(p1.Left + P1Data.Radious, p1.Top + P1Data.Radious, ball.Left + BallData.Radious, ball.Top + BallData.Radious)
                If (e < (BallData.Radious + P1Data.Radious) * 4) And (ball.Left + BallData.Radious / 6 >= p1.Left + P1Data.Radious) Then movething "p1", "u"
            End If
        End If
      Else
        movething "p1", "r"
      End If
    End If
Else
        'send p1data moves
        If P1Data.UpKey = True Then movething "p1", "u"
        If P1Data.DownKey = True Then movething "p1", "d"
        If P1Data.RightKey = True Then movething "p1", "r"
        If P1Data.LeftKey = True Then movething "p1", "l"
End If
    
    



If Form2.gameType = "sp" Then
    'send p2data moves
    If Not P2Data.AI Then
        If P2Data.UpKey = True Then movething "p2", "u"
        If P2Data.DownKey = True Then movething "p2", "d"
        If P2Data.RightKey = True Then movething "p2", "r"
        If P2Data.LeftKey = True Then movething "p2", "l"
    Else
    'move the p2 ( AI BOT )
      If lostmsg.Visible = False Then
        If ball.Left + ball.Width < net.Left Then
            If (p2.Left < (net.Left + net.Left / 2 - 50)) Then movething "p2", "r": Exit Sub
            If (p2.Left > (net.Left + net.Left / 2 + 50)) Then movething "p2", "l": Exit Sub
        Else
            If ((ball.Left + BallData.Radious + (BallData.Radious - BallData.Radious / 6)) < (p2.Left + P2Data.Radious) - 250) Then
                movething "p2", "l"
                If Option6.Value = True And (p2.Left + BallData.Radious + P2Data.Radious - ball.Left < BallData.Radious) Then
                    movething "p2", "l"
                End If
            Else
                movething "p2", "r"
                If Option6.Value = True And (p2.Left + BallData.Radious + P2Data.Radious - ball.Left > BallData.Radious) Then
                    movething "p2", "r"
                End If
            End If
yuy2:
            If (Option5.Value = True Or Option6.Value = True Or asetoase = 1) And (lostmsg.Visible = False) Then
                e = GetDist(p2.Left + P2Data.Radious, p2.Top + P2Data.Radious, ball.Left + BallData.Radious, ball.Top + BallData.Radious)
                If (e < (BallData.Radious + P2Data.Radious) * 4) And (ball.Left + BallData.Radious + (BallData.Radious - BallData.Radious / 6) <= p2.Left + P2Data.Radious) Then movething "p2", "u"
            End If
        End If
      Else
        movething "p2", "l"
      End If
    End If
Else
        'send p1data moves
        If P2Data.UpKey = True Then movething "p2", "u"
        If P2Data.DownKey = True Then movething "p2", "d"
        If P2Data.RightKey = True Then movething "p2", "r"
        If P2Data.LeftKey = True Then movething "p2", "l"
End If

End Sub




Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Form2.gameType = "sp" Then
    DoEvents
    If KeyCode = vbKeyRight Then P2Data.RightKey = True: P2Data.LeftKey = False
    If KeyCode = vbKeyLeft Then P2Data.LeftKey = True: P2Data.RightKey = False
    If KeyCode = vbKeyUp Then
        P2Data.HaveToUp = True
        If P2Data.Uping = 0 Then P2Data.UpKey = True: P2Data.StopedUp = False Else P2Data.UpKey = False
    End If
    If KeyCode = vbKeyDown Then P2Data.DownKey = True
    
    If KeyCode = vbKeyD Then P1Data.RightKey = True: P1Data.LeftKey = False
    If KeyCode = vbKeyA Then P1Data.LeftKey = True: P1Data.RightKey = False
    If KeyCode = vbKeyW Then
        P1Data.HaveToUp = True
        If P1Data.Uping = 0 Then P1Data.UpKey = True: P1Data.StopedUp = False Else P1Data.UpKey = False
    End If
    If KeyCode = vbKeyS Then P1Data.DownKey = True
Else
    If Form2.Atri = "s" Then
        DoEvents
        If KeyCode = vbKeyRight Then P1Data.RightKey = True: P1Data.LeftKey = False
        If KeyCode = vbKeyLeft Then P1Data.LeftKey = True: P1Data.RightKey = False
        If KeyCode = vbKeyUp Then
            moves.AddItem "Q"
            P1Data.HaveToUp = True
            If P1Data.Uping = 0 Then P1Data.UpKey = True: P1Data.StopedUp = False Else P1Data.UpKey = False
        End If
        If KeyCode = vbKeyDown Then P1Data.DownKey = True
    Else
        DoEvents
        If KeyCode = vbKeyRight Then P2Data.RightKey = True: P2Data.LeftKey = False
        If KeyCode = vbKeyLeft Then P2Data.LeftKey = True: P2Data.RightKey = False
        If KeyCode = vbKeyUp Then
            moves.AddItem "Q"
            P2Data.HaveToUp = True
            If P2Data.Uping = 0 Then P2Data.UpKey = True: P2Data.StopedUp = False Else P2Data.UpKey = False
        End If
        If KeyCode = vbKeyDown Then P2Data.DownKey = True
    End If
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub



Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Form2.gameType = "sp" Then
    If KeyCode = vbKeyRight Then P2Data.RightKey = False
    If KeyCode = vbKeyLeft Then P2Data.LeftKey = False
    If KeyCode = vbKeyUp Then P2Data.UpKey = False: P2Data.HaveToUp = False
    If KeyCode = vbKeyDown Then P2Data.DownKey = False
    
    If KeyCode = vbKeyD Then P1Data.RightKey = False
    If KeyCode = vbKeyA Then P1Data.LeftKey = False
    If KeyCode = vbKeyW Then P1Data.UpKey = False: P1Data.HaveToUp = False
    If KeyCode = vbKeyS Then P1Data.DownKey = False
Else
    If Form2.Atri = "s" Then
        If KeyCode = vbKeyRight Then P1Data.RightKey = False
        If KeyCode = vbKeyLeft Then P1Data.LeftKey = False
        If KeyCode = vbKeyUp Then P1Data.UpKey = False: P1Data.HaveToUp = False: moves.AddItem "W"
        If KeyCode = vbKeyDown Then P1Data.DownKey = False
    Else
        If KeyCode = vbKeyRight Then P2Data.RightKey = False
        If KeyCode = vbKeyLeft Then P2Data.LeftKey = False
        If KeyCode = vbKeyUp Then P2Data.UpKey = False: P2Data.HaveToUp = False: moves.AddItem "W"
        If KeyCode = vbKeyDown Then P2Data.DownKey = False
    End If
End If

End Sub

Function GetDist(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
GetDist = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function

Private Sub Timer1_Timer()
If loser = 2 Then
    ball.Top = back.Height - 3135
    ball.Left = net.Left / 2
Else
    ball.Top = back.Height - 3135
    ball.Left = net.Left + (net.Left / 2)
End If
fixall
    BallData.RotSpeed = 0
    BallData.LeftSpeed = 0
Timer1.Enabled = False
End Sub


'Sub loserslost(ByVal wholost As Integer, ByVal reason As String)
Sub loserslost(ByVal wholost As Integer)
        P1Data.Collidable = False
        P2Data.Collidable = False
        loser = wholost
       
         If loser = 1 Then
            P2Data.Score = P2Data.Score + 1
        Else
            P1Data.Score = P1Data.Score + 1
        End If
        
        If (P1Data.Score > 24) Or (P2Data.Score > 24) Then
         If (P1Data.Score > 24) Then lostmsg.Left = 1440 Else lostmsg.Left = 7080
            lostmsg.Visible = True
            lostmsg.Tag = "show"
            Exit Sub
        End If
        
        'lostmsg.Caption = "OOOOOUUUU!!!!!!!!!!!!!!!! P" & wholost & " IS A LOSER!!!"
        lostmsg.Visible = True
        Timer1.Enabled = True
        DoEvents
        DoEvents
        If Form2.gameType = "mp" And Form2.Atri = "s" Then
            Form2.senddat 18, loser
            DoEvents
            DoEvents
            DoEvents
        End If
        If Form2.gameType = "mp" Then
            If Form2.Atri = "c" And loser = 2 Then
                sndPlaySound App.Path & "\Data\dat6.png", 1
            Else
                sndPlaySound App.Path & "\Data\dat8.png", 1
            End If
        Else
            If P1Data.AI = True And P2Data.AI = False And loser = 1 Then sndPlaySound App.Path & "\Data\dat8.png", 1: Exit Sub
            If P1Data.AI = True And P2Data.AI = False And loser = 2 Then sndPlaySound App.Path & "\Data\dat6.png", 1: Exit Sub
            
            If P1Data.AI = False And P2Data.AI = True And loser = 1 Then sndPlaySound App.Path & "\Data\dat8.png", 1: Exit Sub
            If P1Data.AI = False And P2Data.AI = True And loser = 2 Then sndPlaySound App.Path & "\Data\dat6.png", 1: Exit Sub
            
            sndPlaySound App.Path & "\Data\dat8.png", 1
        
        End If
End Sub

