VERSION 5.00
Begin VB.Form welfm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Sign Up"
      Height          =   315
      Left            =   5760
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sign In"
      Height          =   315
      Left            =   5760
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "New User?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Already a Customer?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ONLINE SHOPPING"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "welfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
logfm.Show
Me.Hide
End Sub

Private Sub Command2_Click()
signfm.Show
Me.Hide
End Sub

Private Sub Command3_Click()
cartfm.Show
Me.Hide

End Sub

