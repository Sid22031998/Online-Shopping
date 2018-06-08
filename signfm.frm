VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form signfm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   12165
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      DataField       =   "Gender"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      TabIndex        =   18
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton cnlbtn 
      Caption         =   "CANCEL"
      Height          =   855
      Left            =   6840
      TabIndex        =   17
      Top             =   10320
      Width           =   4095
   End
   Begin VB.CommandButton regbtn 
      Caption         =   "SIGN UP"
      Height          =   855
      Left            =   1080
      TabIndex        =   16
      Top             =   10320
      Width           =   4095
   End
   Begin MSAdodcLib.Adodc registerado 
      Height          =   375
      Left            =   9360
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "UserDetails"
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
   Begin VB.TextBox txtmail 
      DataField       =   "Email-ID"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   8400
      Width           =   7455
   End
   Begin VB.TextBox txtphone 
      DataField       =   "PhoneNumber"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   7560
      Width           =   7455
   End
   Begin VB.TextBox txtadd 
      DataField       =   "Address"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   13
      Top             =   6360
      Width           =   7455
   End
   Begin VB.TextBox txtcnfpass 
      DataField       =   "ConfPass"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   4440
      Width           =   7455
   End
   Begin VB.TextBox txtpass 
      DataField       =   "Password"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   3480
      Width           =   7455
   End
   Begin VB.TextBox txtuser 
      DataField       =   "Username"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   2520
      Width           =   7455
   End
   Begin VB.TextBox txtname 
      DataField       =   "Name"
      DataSource      =   "registerado"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   1560
      Width           =   7455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   8
      Top             =   8520
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   7
      Top             =   7560
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   6
      Top             =   6600
      Width           =   4695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   5400
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   3480
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SIGN UP"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "signfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim index As Integer
Dim Str As String



Private Sub cnlbtn_Click()
End
End Sub

Private Sub Form_Load()
registerado.Recordset.AddNew

txtname.Text = ""
txtuser.Text = ""
txtpass.Text = ""
txtcnfpass.Text = ""
txtadd.Text = ""
txtphone.Text = ""
txtmail.Text = ""
index = 0
List1.AddItem "Male"
List1.AddItem "Female"

End Sub



Private Sub List1_Click()
index = List1.ListIndex
Select Case index
Case 0
Str = "male"
Case 1
Str = "female"
End Select
End Sub

Private Sub regbtn_Click()
P = Trim$(Me.txtpass.Text)

Dim pos As Integer

registerado.Recordset.Fields("Name") = txtname.Text
registerado.Recordset.Fields("Username") = txtuser.Text

registerado.Recordset.Fields("Password") = txtpass.Text
registerado.Recordset.Fields("ConfPass") = txtcnfpass.Text

registerado.Recordset.Fields("Gender") = Str
registerado.Recordset.Fields("Address") = txtadd.Text
registerado.Recordset.Fields("PhoneNumber") = txtphone.Text
registerado.Recordset.Fields("Email-ID") = txtmail.Text






If txtuser.Text = Empty Then
Call MsgBox("Username mandatory", vbCritical, "UserID is required")
Exit Sub
End If
If txtpass.Text = Empty Then
Call MsgBox("Password Required", vbCritical, "Password Required ")
Exit Sub
End If

If (txtcnfpass.Text = Empty) Then
Call MsgBox("Verification Required", vbCritical, "Verification Required ")
Exit Sub
ElseIf txtpass.Text <> txtcnfpass.Text Then
    
    Call MsgBox("Match Failed.Re-confirm password.", vbCritical, "Wrong Password")
    txtcnfpass.Text = Empty
    
    Exit Sub
    
End If


registerado.Recordset.Update
MsgBox "User Registration Successful. Please Login with Username and Password", vbInformation
txtname.Text = ""
txtuser.Text = ""
txtpass.Text = ""
txtcnfpass.Text = ""
txtadd.Text = ""
txtphone.Text = ""
txtmail.Text = ""
Str = ""
index = 0

welfm.Show
Me.Hide

End Sub



