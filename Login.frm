VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form logfm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Login"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton loginbtn 
      Caption         =   "Login"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8880
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\College\4th Sem\VB Project\Siddharth\login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\College\4th Sem\VB Project\Siddharth\login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from UserDetails"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "Password"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2280
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      DataField       =   "Username"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   3
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter PASSWORD:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter USERNAME:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "logfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub Command3_Click()
End

End Sub


Private Sub loginbtn_Click()
Adodc1.RecordSource = "select * from UserDetails where Username='" + Text1.Text + "' AND Password='" + Text2.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Login Unsuccessful", vbCritical, "Please Enter the correct Username and Password"
End

Else
MsgBox "Login Successful", vbInformation, "Successful Attempt"
cartfm.Show
Me.Hide

End If
End Sub

