VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VolleySpheres v1.0   -   George Papadopoulos"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   2280
   End
   Begin VB.Frame Frame9 
      Caption         =   "Credits"
      Height          =   4335
      Left            =   2040
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00808080&
         Height          =   3975
         Left            =   120
         ScaleHeight     =   3915
         ScaleWidth      =   5235
         TabIndex        =   65
         Top             =   240
         Width           =   5295
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   480
            Top             =   1440
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   0
            ScaleHeight     =   945
            ScaleWidth      =   5385
            TabIndex        =   67
            Top             =   -120
            Width           =   5415
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   0
            ScaleHeight     =   1065
            ScaleWidth      =   5265
            TabIndex        =   66
            Top             =   2880
            Width           =   5295
            Begin VB.Image Image6 
               Height          =   855
               Left            =   2160
               Picture         =   "Form2.frx":030A
               Stretch         =   -1  'True
               Top             =   120
               Width           =   855
            End
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2055
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   68
            Text            =   "Form2.frx":C34C
            Top             =   2760
            Width           =   5295
         End
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Options"
      Height          =   4335
      Left            =   2040
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command12 
         Caption         =   "Set Default Gravity"
         Height          =   255
         Left            =   3120
         TabIndex        =   78
         Top             =   2880
         Width           =   2175
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   255
         Left            =   4200
         TabIndex        =   77
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Min             =   3
         Max             =   20
         SelStart        =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   4200
         TabIndex        =   75
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   7
         Min             =   2
         Max             =   20
         SelStart        =   5
         Value           =   5
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   195
         Left            =   4320
         MaxLength       =   12
         TabIndex        =   73
         Text            =   "P2"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   195
         Left            =   4320
         MaxLength       =   12
         TabIndex        =   72
         Text            =   "P1"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Run Game In FullScreen"
         Height          =   255
         Left            =   3120
         TabIndex        =   69
         Top             =   960
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         Caption         =   "Select Network Speed :"
         Height          =   2175
         Left            =   240
         TabIndex        =   56
         Top             =   840
         Width           =   2655
         Begin VB.OptionButton Option11 
            Caption         =   "Custom :"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Tag             =   "5"
            Top             =   1800
            Width           =   975
         End
         Begin VB.OptionButton Option10 
            Caption         =   "33 - 56 Kbps"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Tag             =   "20"
            Top             =   1440
            Width           =   2415
         End
         Begin VB.OptionButton Option9 
            Caption         =   "ISDN"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Tag             =   "10"
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Tx / Cable / xDSL"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Tag             =   "5"
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option7 
            Caption         =   "LAN ( Local Area Network )"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Tag             =   "0"
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   195
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   57
            Text            =   "5"
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Game Speed"
         Height          =   1095
         Left            =   240
         TabIndex        =   52
         Top             =   3120
         Width           =   5055
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Left            =   1680
            TabIndex        =   55
            Top             =   720
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            Max             =   5000
            Value           =   5000
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Custom Speed :"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   3615
         End
         Begin VB.OptionButton Option12 
            Caption         =   "Automatically Adjust Speed For this Computer"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Value           =   -1  'True
            Width           =   3975
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "P2 Keys"
         Height          =   255
         Left            =   3840
         TabIndex        =   51
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "P1 Keys"
         Height          =   255
         Left            =   1920
         TabIndex        =   50
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Players Gravity"
         Height          =   255
         Left            =   3120
         TabIndex        =   76
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Ball Gravity :"
         Height          =   255
         Left            =   3120
         TabIndex        =   74
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Player2 Name :"
         Height          =   255
         Left            =   3120
         TabIndex        =   71
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Player1 Name :"
         Height          =   255
         Left            =   3120
         TabIndex        =   70
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Change Player Keys :"
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Credits"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Options"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Multi Player"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Single Player"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Frame multi 
      Caption         =   "Multi Player"
      Height          =   4335
      Left            =   2040
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame host 
         Caption         =   "Host"
         Height          =   3975
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   5295
         Begin VB.Frame Frame6 
            Caption         =   "Settings"
            Height          =   3615
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   5055
            Begin VB.TextBox Text8 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   2400
               MaxLength       =   20
               TabIndex        =   48
               Text            =   "Host"
               Top             =   1560
               Width           =   2295
            End
            Begin VB.CommandButton Command18 
               Caption         =   "NEXT"
               Height          =   375
               Left            =   3600
               TabIndex        =   46
               Top             =   2280
               Width           =   1095
            End
            Begin VB.TextBox Text6 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   2400
               MaxLength       =   50
               TabIndex        =   45
               Text            =   "Volley Ball"
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   255
               Left            =   2400
               TabIndex        =   43
               Text            =   "2"
               Top             =   480
               Width           =   2295
            End
            Begin VB.Label Label7 
               Caption         =   "Player Nick :"
               Height          =   255
               Left            =   360
               TabIndex        =   47
               Top             =   1560
               Width           =   1335
            End
            Begin VB.Label Label6 
               Caption         =   "Session Name :"
               Height          =   255
               Left            =   360
               TabIndex        =   44
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label4 
               Caption         =   "Max Players :"
               Height          =   255
               Left            =   360
               TabIndex        =   42
               Top             =   480
               Width           =   2415
            End
         End
         Begin VB.ListBox List4 
            Appearance      =   0  'Flat
            Height          =   420
            ItemData        =   "Form2.frx":C3AB
            Left            =   2520
            List            =   "Form2.frx":C3AD
            TabIndex        =   40
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Start Game"
            Height          =   615
            Left            =   3840
            TabIndex        =   26
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   100
            TabIndex        =   25
            Top             =   3480
            Width           =   4935
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   3135
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.Frame join 
         Caption         =   "Join"
         Height          =   3975
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   5295
         Begin VB.Frame Frame5 
            Caption         =   "Lobby"
            Height          =   3615
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Visible         =   0   'False
            Width           =   4935
            Begin VB.ListBox List3 
               Appearance      =   0  'Flat
               Height          =   420
               ItemData        =   "Form2.frx":C3AF
               Left            =   2400
               List            =   "Form2.frx":C3B1
               TabIndex        =   39
               Top             =   360
               Width           =   2175
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               Height          =   2775
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   38
               Top             =   360
               Width           =   4695
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   100
               TabIndex        =   37
               Top             =   3120
               Width           =   4695
            End
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            Height          =   2955
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   4935
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Go Back"
            Height          =   495
            Left            =   4080
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Join Server"
            Height          =   495
            Left            =   2280
            TabIndex        =   29
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Refresh Servers List"
            Height          =   495
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Host A Game ( Server )"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   21
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Join A Game ( Client )"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   22
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1005
         Left            =   960
         TabIndex        =   33
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Select Connection Type"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Image Image5 
         Height          =   3975
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame singlep 
      Caption         =   "SinglePlayer"
      Height          =   4335
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Start Game"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "AI ( Bot )"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Bot Level"
         Height          =   1335
         Left            =   1200
         TabIndex        =   13
         Top             =   840
         Width           =   1215
         Begin VB.OptionButton Option2 
            Caption         =   "Medium"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   650
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Easy"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Difficult"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "AI ( Bot )"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bot Level"
         Height          =   1335
         Left            =   3960
         TabIndex        =   8
         Top             =   840
         Width           =   1215
         Begin VB.OptionButton Option6 
            Caption         =   "Difficult"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Easy"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Medium"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   650
            Width           =   975
         End
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         X1              =   2880
         X2              =   5160
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   5160
         X2              =   5160
         Y1              =   2400
         Y2              =   4200
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   2880
         X2              =   2880
         Y1              =   2400
         Y2              =   4200
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   2400
         X2              =   2400
         Y1              =   2400
         Y2              =   4200
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   2400
         X2              =   240
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   240
         Y1              =   2400
         Y2              =   4200
      End
      Begin VB.Image Image3 
         Height          =   1695
         Left            =   3120
         Picture         =   "Form2.frx":C3B3
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Image Image2 
         Height          =   1695
         Left            =   360
         Picture         =   "Form2.frx":191F5
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   2880
         X2              =   5160
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   2400
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Player 2  ( P2 )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Player 1  ( P1 )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   2640
         X2              =   2640
         Y1              =   120
         Y2              =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   5400
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   2040
      TabIndex        =   31
      Top             =   0
      Width           =   5535
      Begin VB.Image Image4 
         Height          =   3975
         Left            =   120
         Picture         =   "Form2.frx":25207
         Stretch         =   -1  'True
         Top             =   240
         Width           =   5235
      End
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   600
      Picture         =   "Form2.frx":2A0E7
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Game Menu :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Atri As String
Public gameType As String
Public OtherNick As String
Public MyNick As String

Implements DirectXEvent

'1 --> Start Game
'441 --> Chat



Private Sub Command1_Click()
On Error Resume Next
singlep.Visible = True
Frame7.Visible = False
multi.Visible = False
Frame9.Visible = False
PlayClick
End Sub

Sub PlayClick()
On Error Resume Next
sndPlaySound App.Path & "\Data\dat2.png", 1
Timer1.Enabled = False
End Sub



Private Sub Command10_Click()
Form4.Show
End Sub

Private Sub Command11_Click()
On Error Resume Next
PlayClick
senddat 1, ""
DoEvents
DoEvents
DoEvents
Hide
gameType = "mp"
Atri = "s"
DoEvents
DoEvents
frmmain.Show
End Sub


Private Sub Command13_Click()
On Error Resume Next
singlep.Visible = False
multi.Visible = True
host.Visible = False
join.Visible = False
Command7.Visible = True
Command8.Visible = True
List1.Visible = True
PlayClick
End Sub

Private Sub Command12_Click()
Slider2.Value = 6
Slider3.Value = 10
End Sub

Private Sub Command14_Click()
On Error Resume Next
sndPlaySound App.Path & "\Data\dat7.png", 1
DoEvents
DoEvents
UpdateSessionList
DoEvents
End Sub

Private Sub Command15_Click()
On Error Resume Next
If List2.ListIndex = -1 Then MsgBox "Please Select A Game From The List Bellow": Exit Sub
JoinL
gameType = "mp"
Atri = "c"
End Sub

Private Sub Command16_Click()
On Error Resume Next
'Pass the bot settings to the form1 ai options
Form1.Check1.Value = Check1.Value
Form1.Check2.Value = Check2.Value

Form1.Option1.Value = Option1.Value
Form1.Option2.Value = Option2.Value
Form1.Option3.Value = Option3.Value
Form1.Option4.Value = Option4.Value
Form1.Option5.Value = Option5.Value
Form1.Option6.Value = Option6.Value

gameType = "sp"

Me.Hide
frmmain.Show
End Sub

Private Sub Command17_Click()
On Error Resume Next
singlep.Visible = False
multi.Visible = True
host.Visible = False
join.Visible = False
Command7.Visible = True
Command8.Visible = True
List1.Visible = True
PlayClick
End Sub

Private Sub Command18_Click()
UseL "s"
NewL
gameType = "mp"
Atri = "s"
Frame6.Visible = False
sndPlaySound App.Path & "\Data\dat7.png", 1
End Sub

Private Sub Command2_Click()
On Error Resume Next
singlep.Visible = False
multi.Visible = True
host.Visible = False
join.Visible = False
Command7.Visible = True
Command8.Visible = True
List1.Visible = True
Label5.Visible = True
Frame7.Visible = False
Frame9.Visible = False
PlayClick

basMain.Init

'ENUM CONNECTIONS
List1.Clear
Dim lNumConnections As Long
Dim sName As String
Dim lCount As Long
    
lNumConnections = DPEnumConnections.GetCount
    
For lCount = 1 To lNumConnections
    sName = DPEnumConnections.GetName(lCount)
    List1.AddItem sName
Next lCount

End Sub

Private Sub Command3_Click()
On Error Resume Next
PlayClick
Frame7.Visible = True
Frame9.Visible = False
End Sub

Private Sub Command4_Click()
On Error Resume Next
PlayClick
End Sub

Private Sub Command5_Click()
On Error Resume Next
DPlay.Close
Set DPlay = Nothing
PlayClick
DoEvents
MsgBox "VolleySpheres v1.0 By George Papadopoulos" & vbCrLf & "       Visit www.spider-bit.com for more"
DoEvents
DoEvents
Unload Form1
Unload frmmain
End
End Sub

Private Sub Command6_Click()
On Error Resume Next
PlayClick
Frame9.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command7_Click()
On Error Resume Next
singlep.Visible = False
multi.Visible = True
host.Visible = True
join.Visible = False
PlayClick
Frame6.Visible = True
End Sub

Private Sub Command8_Click()
On Error Resume Next
singlep.Visible = False
multi.Visible = True
host.Visible = False
join.Visible = True
sndPlaySound App.Path & "\Data\dat7.png", 1
PlayClick
UseL "c"
End Sub



Private Sub Command9_Click()
Form3.Show
End Sub

Public Sub DirectXEvent_dxCallback(ByVal eventid As Long)
On Error Resume Next

Dim DPMsg As DirectPlayMessage, strFrom As String
Dim MsgType As Long, FromPlayerID As Long, ToPlayerID As Long


yy = DPlay.GetMessageCount(MyPlayerID)
'If yy = 1 Then n13 = 0: GoTo uii
For n13 = 0 To yy
uii:
Set DPMsg = DPlay.Receive(FromPlayerID, ToPlayerID, DPRECEIVE_ALL)
If yy = n13 Then GoTo uy


MsgType = DPMsg.ReadLong

If FromPlayerID = DPID_SYSMSG Then
    Select Case MsgType
        Case DPSYS_DESTROYPLAYERORGROUP, DPSYS_CREATEPLAYERORGROUP
            GetCurrentParticipants
        Case DPSYS_ADDGROUPTOGROUP, DPSYS_ADDPLAYERTOGROUP
            sndPlaySound App.Path & "\data\dat3.png", 1
            GetCurrentParticipants
        Case DPSYS_CREATEPLAYERORGROUP, DPSYS_DELETEGROUPFROMGROUP
            GetCurrentParticipants
        Case DPSYS_DELETEPLAYERFROMGROUP, DPSYS_DESTROYPLAYERORGROUP
            GetCurrentParticipants
        Case DPSYS_HOST
            DPlay.Close
            Frame5.Visible = False
            join.Visible = False
            host.Visible = False
        Case DPSYS_SESSIONLOST
    End Select
Else
    Dim param As String
    Dim e As String
    Dim balx, baly As Double
    Select Case MsgType
        Case 411
        'Server Send A String
            Dim Msg As String
            Msg = DPMsg.ReadString
            strFrom = DPlay.GetPlayerFormalName(FromPlayerID)
            Text4 = Text4 & "< " & strFrom & " >  " & Msg & vbCrLf
            sndPlaySound App.Path & "\Data\dat1.png", 1
            
        Case 422
        'Client Send A String
            Dim Msg2 As String
            Msg2 = DPMsg.ReadString
            strFrom = DPlay.GetPlayerFormalName(FromPlayerID)
            Text1 = Text1 & "< " & strFrom & " >  " & Msg2 & vbCrLf
            sndPlaySound App.Path & "\Data\dat1.png", 1
            
        Case 1
            Me.Hide
            DoEvents
            Me.Hide
            DoEvents
            DoEvents
            Me.gameType = "mp"
            Atri = "c"
            DoEvents
            frmmain.Show
         
        Case 2
            param = DPMsg.ReadString
            r = InStr(1, param, ";")
            p2x = Mid(param, 1, r - 1)
            p2y = Val(Mid(param, r + 1))
            Form1.p2.Top = p2y
            For n = 1 To Len(p2x)
                e = Mid(p2x, n, 1)
                Form1.movething "p2", e
            Next n
            DoEvents
            
        Case 5
            Dim param2 As String
            param2 = DPMsg.ReadString
            r = InStr(1, param2, ";")
            r2 = InStr(1, param2, "?")
            ballx = Val(Mid(param2, 1, r - 1))
            bally = Val(Mid(param2, r + 1, r2 - 1))
            ballangle = Val(Mid(param2, r2 + 1))
            Form1.ball.Top = bally
            Form1.ball.Left = ballx
            BallData.Angle = ballangle
            DoEvents
                
        Case 7
            Dim param3 As String
            param3 = DPMsg.ReadString
            r = InStr(1, param3, ";")
            p1x = Mid(param3, 1, r - 1)
            p1y = Val(Mid(param3, r + 1))
            Form1.p1.Top = p1y
            For n = 1 To Len(p1x)
                e = Mid(p1x, n, 1)
                Form1.movething "p1", e
            Next n
            DoEvents
            
            
        Case 18
            Dim par As String
            par = DPMsg.ReadString
            Form1.loserslost par
            
            
    End Select
End If
'If n13 = 0 Then GoTo uy
Next n13
uy:
End Sub

Private Sub Form_Load()
On Error Resume Next
Image5.Picture = Image4.Picture
basMain.Init
Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", "http://www.spider-bit.com", "", App.Path, 1)
Me.Show
End Sub

Private Sub Image1_Click()
Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", "http://www.spider-bit.com", "", App.Path, 1)

End Sub

Private Sub Image6_Click()
Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", "http://www.spider-bit.com", "", App.Path, 1)

End Sub

Private Sub List1_Click()
Command7.Enabled = True
Command8.Enabled = True
End Sub

Private Sub Option11_Click()
If Option11.Value = True Then
    Text7.Enabled = True
Else
    Text7.Enabled = False
End If
End Sub


Public Sub senddat(ByVal code As Long, ByVal dat As String)
    Dim DPMsg As DirectPlayMessage
    
    Set DPMsg = DPlay.CreateMessage
    
    With DPMsg
        .WriteLong code
        .WriteString dat
    End With
    

     DPlay.Send MyPlayerID, DPID_ALLPLAYERS, DPSEND_DEFAULT, DPMsg
    
    Set DPMsg = Nothing
    If err.Description <> "" Then MsgBox err.Description
End Sub




Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
sndPlaySound App.Path & "\Data\dat5.png", 1


If KeyCode = 13 And Trim(Text2) <> "" Then
    senddat 411, Text2
    Text1 = Text1 & "< " & MyNick & " > " & Text2 & vbCrLf
    Text1.SelStart = Len(Text1.Text)
    Text2 = ""
    Text2.SetFocus
    
    sndPlaySound App.Path & "\Data\dat1.png", 1
    
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
sndPlaySound App.Path & "\Data\dat5.png", 1


If KeyCode = 13 And Trim(Text3) <> "" Then
    
    senddat 422, Text3
    
    Text4 = Text4 & "< " & MyNick & " > " & Text3 & vbCrLf
    Text4.SelStart = Len(Text4.Text)
    Text3 = ""
    Text3.SetFocus
    
    sndPlaySound App.Path & "\Data\dat1.png", 1
End If

End Sub

Private Sub Text7_Change()
Option11.Tag = Text7
End Sub






Public Sub UseL(ByVal cli As String)

On Error GoTo err

Me.MousePointer = vbHourglass

Dim DPAddress As DirectPlayAddress
    
Set DPAddress = DPEnumConnections.GetAddress(List1.ListIndex + 1)
DPlay.InitializeConnection DPAddress
    
On Error GoTo 0
  
If cli = "c" Then UpdateSessionList

Set DPAddress = Nothing

Me.MousePointer = vbDefault

Exit Sub
err:
    MsgBox err.Description, vbSystemModal Or vbOKOnly, err.Source & " has caused an error"
    Me.MousePointer = vbDefault
    Call Command2_Click
End Sub


Public Sub UpdateSessionList()

On Error GoTo err

Me.MousePointer = vbHourglass
    
Dim SessionCount As Integer, X As Integer
Dim SessionData As DirectPlaySessionData
Dim Details As String
    
List2.Clear
    
Set SessionData = DPlay.CreateSessionData
   
SessionData.SetGuidApplication ""
  
Set DPEnumSessions = DPlay.GetDPEnumSessions(SessionData, 0, DPENUMSESSIONS_ALL)
 
SessionCount = DPEnumSessions.GetCount
    
For X = 1 To SessionCount
    Set SessionData = DPEnumSessions.GetItem(X)
    
    If SessionData.GetCurrentPlayers < 2 Then
        Details = SessionData.GetSessionName & " (" & SessionData.GetCurrentPlayers & "/" & SessionData.GetMaxPlayers & ")"
        List2.AddItem Details
    End If
Next X

Me.MousePointer = vbDefault

Exit Sub
err:
    MsgBox err.Description, vbSystemModal Or vbOKOnly, err.Source & " has caused an error"
    Me.MousePointer = vbDefault
    Call Command2_Click
End Sub








Private Sub JoinL()
On Error GoTo err

Me.MousePointer = vbHourglass
    
Dim SessionData As DirectPlaySessionData

Set SessionData = DPEnumSessions.GetItem(List2.ListIndex + 1)
    
DPlay.Open SessionData, DPOPEN_JOIN

'Load frmChat

    NotificationID = DX7.CreateEvent(Me)

    SendMCIString "close all", False

    If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
        End
    End If
    SendMCIString "set cd time format tmsf wait", True


Frame5.Visible = True
    

PlayerName = InputBox("Please enter your player Id", "Name entry")
If PlayerName = "" Then PlayerName = "Client"
If Len(PlayerName) > 20 Then PlayerName = Mid(PlayerName, 1, 20)
MyNick = PlayerName
PlayerHandle = PlayerName & "Handle"

MyPlayerID = DPlay.CreatePlayer(PlayerHandle, PlayerName, NotificationID, 0)
'Hide
'frmChat.Caption = "Volley (" & PlayerName & ")"
'frmChat.LblName.Caption = PlayerName
'frmChat.Show
'frmChat.GetCurrentParticipants
Atri = "c"

GetCurrentParticipants

sndPlaySound App.Path & "\data\dat3.png", 1

Me.MousePointer = vbDefault

Exit Sub
err:
    MsgBox err.Description, vbSystemModal Or vbOKOnly, err.Source & " has caused an error"
    Me.MousePointer = vbDefault
    Call Command2_Click

End Sub


Private Sub NewL()
On Error Resume Next
Me.MousePointer = vbHourglass
    
Dim SessionData As DirectPlaySessionData

Set SessionData = DPlay.CreateSessionData
    
If Text6 = "" Then Text6 = "Volley Ball"

Dim i As String
i = Text6

With SessionData
    .SetMaxPlayers 2
    .SetCurrentPlayers 0
    .SetSessionName i
    .SetGuidApplication AppGuid
    .SetFlags DPSESSION_MIGRATEHOST
    .SetSessionPassword ""
End With

DPlay.Open SessionData, DPOPEN_CREATE

Dim PlayerHandle As String
Dim PlayerName As String
    
    
'Load frmChat

    NotificationID = DX7.CreateEvent(Me)

    SendMCIString "close all", False

    If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
        End
    End If
    SendMCIString "set cd time format tmsf wait", True

    
    
    
If Text8 = "" Then Text8 = "Host"
MyNick = Text8
PlayerName = Text8
PlayerHandle = PlayerName & "Handle"

MyPlayerID = DPlay.CreatePlayer(PlayerHandle, PlayerName, NotificationID, 0)
    
GetCurrentParticipants

Me.MousePointer = vbDefault

End Sub




Public Sub GetCurrentParticipants()
    Dim lCount As Long
    
    List3.Clear
    List4.Clear
    
    
    Set DPEnumPlayers = DPlay.GetDPEnumPlayers("", DPENUMPLAYERS_ALL)
    lCount = DPEnumPlayers.GetCount
    Dim i As String
    i = ""
    Do Until lCount = 0
        If Atri = "c" Then
            i = DPEnumPlayers.GetLongName(lCount)
            If i <> MyNick Then OtherNick = i
            List3.AddItem (i)
        Else
            i = DPEnumPlayers.GetLongName(lCount)
            If i <> MyNick Then OtherNick = i
            List4.AddItem (i)
        End If
        lCount = lCount - 1
    Loop
End Sub

Private Sub Timer1_Timer()
Text9.Top = Text9.Top - 10
If Text9.Top + Text9.Height < Picture3.Height Then Text9.Top = Picture2.Top + Picture2.Top / 1.2
End Sub

Private Sub Timer2_Timer()
Me.Show
Unload Form1
DoEvents
Load Form1
Load frmmain
Timer2.Enabled = False
End Sub
