VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cartfm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "CLEAR CART"
      Height          =   615
      Left            =   5520
      TabIndex        =   7
      Top             =   7560
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2775
      Left            =   960
      TabIndex        =   5
      Top             =   3000
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD TO CART"
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   7560
      Width           =   3015
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Amount"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   6240
      Width           =   3615
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   8
      Top             =   6240
      Width           =   3735
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   2160
      Width           =   5415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Selected:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -720
      TabIndex        =   3
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Online Shopping"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Product List"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
End
Attribute VB_Name = "cartfm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim slno As Integer
Dim pname As String
Dim price As Long
Dim totalbill As Long
Dim index As Integer
Dim rnum As Integer

Private Sub Command1_Click()
If index = 6 Then
MsgBox ("Please Select an item")
Else

rnum = rnum + 1
MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
MSFlexGrid1.TextMatrix(rnum, 0) = slno
MSFlexGrid1.TextMatrix(rnum, 1) = pname
MSFlexGrid1.TextMatrix(rnum, 2) = 1
MSFlexGrid1.TextMatrix(rnum, 3) = price
slno = slno + 1
totalbill = totalbill + price
End If


index = 6
price = 0
pname = ""
Label12.Caption = "---"
Label13.Caption = totalbill
End Sub





Private Sub Command4_Click()
MSFlexGrid1.Clear
rnum = 0
MSFlexGrid1.TextMatrix(rnum, 0) = "SLNO"
MSFlexGrid1.TextMatrix(rnum, 1) = "NAME"
MSFlexGrid1.TextMatrix(rnum, 3) = "PRICE"
MSFlexGrid1.TextMatrix(rnum, 2) = "QUANTITY"
price = 0
pname = ""
totalbill = 0
slno = 1
index = 6
Label13.Caption = "---"
Label12.Caption = "---"
End Sub

Private Sub Form_Load()

List1.AddItem "Washing Machine"
List1.AddItem "Jeans"
List1.AddItem "Shoes"
List1.AddItem "Sunglasses"
List1.AddItem "Innerwear"
List1.AddItem "Chips"
List1.AddItem "Biscuit"
List1.AddItem "Cake"
List1.AddItem "Cold Drink"
List1.AddItem "Chocolate"
rnum = 0
MSFlexGrid1.TextMatrix(rnum, 0) = "SLNO"
MSFlexGrid1.TextMatrix(rnum, 1) = "NAME"
MSFlexGrid1.TextMatrix(rnum, 3) = "PRICE"
MSFlexGrid1.TextMatrix(rnum, 2) = "QUANTITY"
price = 0
pname = ""
totalbill = 0
slno = 1
index = 6

End Sub


Private Sub List1_Click()
index = List1.ListIndex
Select Case index

Case 0
Label12.Caption = "Godrej EasyCare"
price = 35000
pname = "Godrej EasyCare"

Case 1
Label12.Caption = "Denim FeelFree"
price = 999
pname = "Denim FeelFree"

Case 2
Label12.Caption = "Adidas SpringBlade"
price = 3999
pname = "Adidas SpringBlade"

Case 3
Label12.Caption = "FastTrack Aviator"
price = 1499
pname = "FastTrack Aviator"

Case 4
Label12.Caption = "VanHeusen SportWear"
price = 499
pname = "VanHeusen SportWear"

Case 5
Label12.Caption = "Too Yumm"
price = 30
pname = "Too Yumm"

Case 6
Label12.Caption = "UNIBIC Cashew and Nut"
price = 50
pname = "UNIBIC Cashew and Nut"

Case 7
Label12.Caption = "Butter and Honey"
price = 300
pname = "Butter and Honey"

Case 8
Label12.Caption = "Coca-Cola"
price = 80
pname = "Coca-Cola"

Case 9
Label12.Caption = "Dark Chocolate"
price = 120
pname = "Dark Chocolate"


End Select

End Sub

