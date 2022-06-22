VERSION 5.00
Begin VB.Form viewemp 
   BackColor       =   &H00800080&
   Caption         =   "View Employee Details"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   LinkTopic       =   "Form1"
   Picture         =   "viewemp.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   19755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   8910
      ScaleHeight     =   2355
      ScaleWidth      =   1875
      TabIndex        =   32
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton last 
      BackColor       =   &H00CAA715&
      Caption         =   "Last"
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
      Left            =   11370
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton previous1 
      BackColor       =   &H00CAA715&
      Caption         =   "Previous"
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
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton Next1 
      BackColor       =   &H00CAA715&
      Caption         =   "Next"
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
      Left            =   9930
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton first1 
      BackColor       =   &H00CAA715&
      Caption         =   "First"
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
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00000000&
      DataField       =   "Phone_No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   13920
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00000000&
      DataField       =   "Calculated"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   9570
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7320
      Width           =   2535
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00000000&
      DataField       =   "Net_Pay"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   14730
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   6600
      Width           =   2535
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00000000&
      DataField       =   "Allowance"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   9570
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   6600
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00000000&
      DataField       =   "Basic_Pay"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   4770
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   6600
      Width           =   2535
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   13920
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   4440
      Width           =   2535
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1785
      Left            =   6120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   3720
      Width           =   2535
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   13920
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      DataField       =   "Gender"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      DataField       =   "DOJ"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   13920
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      DataField       =   "DOB"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton exit1 
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
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8400
      Width           =   1335
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
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      DataField       =   "Emp_ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   13920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   585
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Line Line9 
      DrawMode        =   8  'Xor Pen
      X1              =   2760
      X2              =   17400
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line7 
      DrawMode        =   8  'Xor Pen
      Index           =   1
      X1              =   17400
      X2              =   17400
      Y1              =   6240
      Y2              =   8040
   End
   Begin VB.Line Line8 
      DrawMode        =   8  'Xor Pen
      X1              =   10920
      X2              =   17400
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line7 
      DrawMode        =   8  'Xor Pen
      Index           =   0
      X1              =   2760
      X2              =   2760
      Y1              =   6240
      Y2              =   8040
   End
   Begin VB.Line Line6 
      DrawMode        =   8  'Xor Pen
      X1              =   8760
      X2              =   2760
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line5 
      DrawMode        =   8  'Xor Pen
      X1              =   3480
      X2              =   16680
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line4 
      DrawMode        =   8  'Xor Pen
      X1              =   16680
      X2              =   16680
      Y1              =   1200
      Y2              =   5640
   End
   Begin VB.Line Line3 
      DrawMode        =   8  'Xor Pen
      X1              =   11160
      X2              =   16680
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      DrawMode        =   8  'Xor Pen
      X1              =   3480
      X2              =   3480
      Y1              =   1200
      Y2              =   5640
   End
   Begin VB.Line Line1 
      DrawMode        =   8  'Xor Pen
      X1              =   8520
      X2              =   3480
      Y1              =   1200
      Y2              =   1200
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
      Left            =   8670
      TabIndex        =   35
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2490
      TabIndex        =   12
      Top             =   6600
      Width           =   1935
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
      Left            =   8490
      TabIndex        =   34
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "View Employee Details"
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
      Left            =   7950
      TabIndex        =   33
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11520
      TabIndex        =   27
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7170
      TabIndex        =   15
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   12330
      TabIndex        =   14
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7170
      TabIndex        =   13
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11760
      TabIndex        =   11
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11520
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Joining"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11040
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11400
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "viewemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Dim str As String

Private Sub back1_Click()
Home.Show
viewemp.Hide
End Sub
Private Sub exit1_Click()
End
End Sub
Sub display()
Text1.Text = rs!Name
Text2.Text = rs!Emp_ID
Text3.Text = rs!DOB
Text4.Text = rs!DOJ
Text5.Text = rs!Gender
Text6.Text = rs!Designation
Text7.Text = rs!Address
Text8.Text = rs!PIN
Text9.Text = rs!Basic_Pay
Text10.Text = rs!Allowance
Text11.Text = rs!Net_pay
Text12.Text = rs!Calculated
Text13.Text = rs!Phone_No

End Sub

Private Sub first1_Click()
rs.MoveFirst
display
Picture1.Picture = LoadPicture(rs!Photo)
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\Admin\Desktop\My Project\Databases\EmployeeDB.mdb")
Set rs = db.OpenRecordset("select * from empdetails")
display
Picture1.Picture = LoadPicture(rs!Photo)
End Sub


Private Sub last_Click()
rs.MoveLast
display
Picture1.Picture = LoadPicture(rs!Photo)
End Sub

Private Sub Next1_Click()
rs.MoveNext
If rs.EOF Then
rs.MoveFirst
display
Picture1.Picture = LoadPicture(rs!Photo)
Else
display
Picture1.Picture = LoadPicture(rs!Photo)
End If
End Sub

Private Sub previous1_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Picture1.Picture = LoadPicture(rs!Photo)
Else
display
Picture1.Picture = LoadPicture(rs!Photo)
End If
End Sub
