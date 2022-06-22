VERSION 5.00
Begin VB.Form delemp 
   BackColor       =   &H80000007&
   Caption         =   "Delete Employee"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   LinkTopic       =   "Form1"
   Picture         =   "delemp.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   19755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton delete1 
      BackColor       =   &H00CAA715&
      Caption         =   "Delete"
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
      Left            =   9210
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8280
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
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8280
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
      Left            =   10650
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8280
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
      Left            =   7770
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8280
      Width           =   1335
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
      Left            =   12090
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8280
      Width           =   1335
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   6045
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1440
      Width           =   2535
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   13845
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1440
      Width           =   2535
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
      Left            =   15525
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8280
      Width           =   1335
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
      Left            =   16965
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8280
      Width           =   1335
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   6045
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2160
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   13845
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2160
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   6045
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2880
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   13845
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2880
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
      ForeColor       =   &H80000005&
      Height          =   1785
      Left            =   6045
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3600
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   13845
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4320
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   4695
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   6480
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   9495
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   6480
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   14655
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6480
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   9495
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   7200
      Width           =   2535
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
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   13845
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   8835
      ScaleHeight     =   2355
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
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
      Left            =   3645
      TabIndex        =   31
      Top             =   1440
      Width           =   2055
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
      Left            =   11325
      TabIndex        =   30
      Top             =   1440
      Width           =   2175
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
      Left            =   3405
      TabIndex        =   29
      Top             =   2160
      Width           =   2295
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
      Left            =   10965
      TabIndex        =   28
      Top             =   2160
      Width           =   2535
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
      Left            =   3645
      TabIndex        =   27
      Top             =   2880
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
      Left            =   11445
      TabIndex        =   26
      Top             =   2880
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
      Left            =   3645
      TabIndex        =   25
      Top             =   3600
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
      Left            =   11685
      TabIndex        =   24
      Top             =   4320
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
      Left            =   7095
      TabIndex        =   23
      Top             =   6480
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
      Left            =   12255
      TabIndex        =   22
      Top             =   6480
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
      Left            =   7095
      TabIndex        =   21
      Top             =   7200
      Width           =   2055
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
      Left            =   11445
      TabIndex        =   20
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Employee"
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
      Left            =   8550
      TabIndex        =   19
      Top             =   0
      Width           =   2655
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
      Left            =   8415
      TabIndex        =   18
      Top             =   840
      Width           =   2775
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
      Left            =   2415
      TabIndex        =   17
      Top             =   6480
      Width           =   1935
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
      Left            =   8595
      TabIndex        =   16
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Line Line1 
      DrawMode        =   8  'Xor Pen
      X1              =   8445
      X2              =   3405
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      DrawMode        =   8  'Xor Pen
      X1              =   3405
      X2              =   3405
      Y1              =   1080
      Y2              =   5520
   End
   Begin VB.Line Line3 
      DrawMode        =   8  'Xor Pen
      X1              =   11085
      X2              =   16605
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line4 
      DrawMode        =   8  'Xor Pen
      X1              =   16605
      X2              =   16605
      Y1              =   1080
      Y2              =   5520
   End
   Begin VB.Line Line5 
      DrawMode        =   8  'Xor Pen
      X1              =   3405
      X2              =   16605
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line6 
      DrawMode        =   8  'Xor Pen
      X1              =   8685
      X2              =   2685
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line7 
      DrawMode        =   8  'Xor Pen
      Index           =   0
      X1              =   2685
      X2              =   2685
      Y1              =   6120
      Y2              =   7920
   End
   Begin VB.Line Line8 
      DrawMode        =   8  'Xor Pen
      X1              =   10845
      X2              =   17325
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line7 
      DrawMode        =   8  'Xor Pen
      Index           =   1
      X1              =   17325
      X2              =   17325
      Y1              =   6120
      Y2              =   7920
   End
   Begin VB.Line Line9 
      DrawMode        =   8  'Xor Pen
      X1              =   2685
      X2              =   17325
      Y1              =   7920
      Y2              =   7920
   End
End
Attribute VB_Name = "delemp"
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


Private Sub delete1_Click()
confirm = MsgBox("Do You Want To Delete Employee Profile?", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete
MsgBox "Profile Has Been Deleted Successfully!!!...", vbInformation, "Message"
Else
MsgBox "Profile Is Not Deleted..."
End If
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
End If
Picture1.Picture = LoadPicture(rs!Photo)
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

