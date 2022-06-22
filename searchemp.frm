VERSION 5.00
Begin VB.Form searchemp 
   Caption         =   "Search Employee"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   Picture         =   "searchemp.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   19755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Search2 
      BackColor       =   &H00CAA715&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13410
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   960
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   525
      Left            =   10410
      TabIndex        =   37
      Text            =   "-SELECT-"
      Top             =   1009
      Width           =   2895
   End
   Begin VB.CommandButton left 
      BackColor       =   &H00CAA715&
      Height          =   495
      Left            =   14970
      Picture         =   "searchemp.frx":2F358
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1009
      Width           =   615
   End
   Begin VB.CommandButton right 
      BackColor       =   &H00CAA715&
      Height          =   495
      Left            =   15570
      Picture         =   "searchemp.frx":5738C
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1009
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   9240
      ScaleHeight     =   2355
      ScaleWidth      =   1875
      TabIndex        =   34
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton reset 
      BackColor       =   &H00CAA715&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00000000&
      DataField       =   "Phone_No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   9990
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H00CAA715&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton back1 
      BackColor       =   &H00CAA715&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00000000&
      DataField       =   "Calculated"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   16530
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7560
      Width           =   2415
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00000000&
      DataField       =   "Net_Pay"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   11850
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   7560
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00000000&
      DataField       =   "Allowance"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7410
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   7560
      Width           =   2415
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00000000&
      DataField       =   "Basic_Pay"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   7560
      Width           =   2415
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00000000&
      DataField       =   "Gender"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   13830
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00000000&
      DataField       =   "PIN"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1845
      Left            =   13830
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00000000&
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   13830
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00000000&
      DataField       =   "Designation"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   6390
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      DataField       =   "DOJ"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   6390
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      DataField       =   "DOB"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   6390
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      DataField       =   "Emp_ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   6390
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   6390
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   4050
      TabIndex        =   1
      Top             =   990
      Width           =   2415
   End
   Begin VB.CommandButton search1 
      BackColor       =   &H00CAA715&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   5  'Dash-Dot-Dot
      DrawMode        =   8  'Xor Pen
      Height          =   1455
      Index           =   1
      Left            =   10290
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6015
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   5  'Dash-Dot-Dot
      DrawMode        =   8  'Xor Pen
      Height          =   1455
      Index           =   0
      Left            =   3450
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Employee By Designation"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   10650
      TabIndex        =   39
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Details"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   4200
      TabIndex        =   31
      Top             =   1640
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      DrawMode        =   8  'Xor Pen
      Height          =   4695
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   13095
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Details"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3960
      TabIndex        =   32
      Top             =   6680
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      DrawMode        =   8  'Xor Pen
      Height          =   1215
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   18495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Employee By Name Or ID"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3570
      TabIndex        =   30
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7350
      TabIndex        =   29
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculated"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   14610
      TabIndex        =   13
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Pay"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   10170
      TabIndex        =   12
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Allowance"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   5490
      TabIndex        =   11
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Pay"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   810
      TabIndex        =   10
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3510
      TabIndex        =   9
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   11550
      TabIndex        =   8
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   11520
      TabIndex        =   7
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   11550
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Joining"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3510
      TabIndex        =   5
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3510
      TabIndex        =   4
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3510
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   555
      Left            =   3510
      TabIndex        =   2
      Top             =   2520
      Width           =   2760
   End
End
Attribute VB_Name = "searchemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub back1_Click()
Home.Show
searchemp.Hide
End Sub

Private Sub Exit_Click()
End
End Sub
Sub resetfields()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Set Picture1.Picture = Nothing
Combo1.ListIndex = -1
End Sub
Sub display()
Text2.Text = rs!Name
Text3.Text = rs!Emp_ID
Text4.Text = rs!DOB
Text5.Text = rs!DOJ
Text6.Text = rs!Gender
Text7.Text = rs!Designation
Text8.Text = rs!Address
Text9.Text = rs!PIN
Text10.Text = rs!Basic_Pay
Text11.Text = rs!Allowance
Text12.Text = rs!Net_pay
Text13.Text = rs!Calculated
Text14.Text = rs!Phone_No
Picture1.Picture = LoadPicture(rs!Photo)
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Desktop\My Project\Databases\EmployeeDB.mdb;Persist Security Info= False"
rs.Open "select * from empdetails", con, adOpenDynamic, adLockPessimistic

Combo1.AddItem ("Accountant")
Combo1.AddItem ("Assistant")
Combo1.AddItem ("Collection Person")
Combo1.AddItem ("Driver")
Combo1.AddItem ("Writer")

End Sub


Private Sub left_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Else
display
End If
End Sub

Private Sub reset_Click()
resetfields
End Sub

Private Sub right_Click()
rs.MoveNext
If rs.EOF Then
rs.MoveFirst
display
Else
display
End If
End Sub

Private Sub search1_Click()
rs.Close
rs.Open "select * from empdetails where Emp_ID = '" + Text1.Text + "'or Name = '" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
Else
MsgBox "Employee Not Found, Please Try again..."
End If
End Sub

Private Sub Search2_Click()
rs.Close
rs.Open "select * from empdetails where Designation = '" + Combo1.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
Else
MsgBox "Employee Not Found, Please Try again..."
End If
End Sub
