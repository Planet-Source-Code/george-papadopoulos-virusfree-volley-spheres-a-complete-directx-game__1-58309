VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "P2 Controls"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3495
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
      Caption         =   "Move Left : Left Key"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Move Right : Righ Key"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Jump : UP Key"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub
