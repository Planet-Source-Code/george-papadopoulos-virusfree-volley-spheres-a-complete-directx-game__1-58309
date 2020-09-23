VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "P1 Controls"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Move Left : D Key"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Move Right : A Key"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Jump : W Key"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

