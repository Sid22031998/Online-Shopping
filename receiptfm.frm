VERSION 5.00
Begin VB.Form receiptfm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22575
   LinkTopic       =   "Form1"
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CLOSE"
      Height          =   1215
      Left            =   18240
      TabIndex        =   3
      Top             =   3360
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK TO CART"
      Height          =   1215
      Left            =   18240
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Frame1"
      Height          =   10455
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   16455
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "WORK IN PROGRESS... STAY PUT"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   2160
         TabIndex        =   4
         Top             =   2400
         Width           =   9375
      End
      Begin VB.Label user 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   4455
      End
   End
End
Attribute VB_Name = "receiptfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cartfm.Show
Me.Hide

End Sub

Private Sub Command2_Click()
End

End Sub
