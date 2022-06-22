VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form addemp 
   Caption         =   "Add Employee"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   LinkTopic       =   "Form1"
   Picture         =   "addemp.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   19755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "NOTE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   6570
      TabIndex        =   35
      Top             =   7560
      Width           =   6615
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "NET PAY WILL BE AUTOMATICALLY CALCULATED"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   960
         TabIndex        =   36
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9240
      MaxLength       =   10
      TabIndex        =   33
      Top             =   2070
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CAA715&
      Caption         =   "Upload Pic"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3840
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   16776960
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   16800
      TabIndex        =   31
      Text            =   "-SELECT-"
      Top             =   6840
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   5640
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
      Connect         =   $"addemp.frx":3AFCB
      OLEDBString     =   $"addemp.frx":3B052
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from empdetails"
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
   Begin VB.CommandButton addemp1 
      BackColor       =   &H00CAA715&
      Caption         =   "Add Employee"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8280
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   9240
      MaxLength       =   6
      TabIndex        =   26
      Top             =   4680
      Width           =   2655
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
      Left            =   16635
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8430
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
      Left            =   18075
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8430
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   15960
      ScaleHeight     =   2355
      ScaleWidth      =   1875
      TabIndex        =   21
      Top             =   1200
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   2550
      TabIndex        =   20
      Top             =   3750
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   0
      CalendarForeColor=   -2147483637
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483633
      CalendarTrailingForeColor=   16761024
      Format          =   108462081
      CurrentDate     =   44319
      MaxDate         =   73415
      MinDate         =   44197
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   12000
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6840
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   7320
      TabIndex        =   17
      Top             =   6810
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   2280
      TabIndex        =   16
      Top             =   6810
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   2520
      TabIndex        =   12
      Text            =   "-SELECT-"
      Top             =   4470
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10560
      TabIndex        =   10
      Top             =   1230
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9240
      TabIndex        =   9
      Top             =   1230
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2550
      TabIndex        =   8
      Top             =   2070
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   2550
      TabIndex        =   1
      Top             =   1230
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2550
      TabIndex        =   0
      Top             =   2970
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      MousePointer    =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   0
      CalendarForeColor=   16777215
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   16761024
      Format          =   108462081
      CurrentDate     =   36892
      MaxDate         =   43831
      MinDate         =   2
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Line Line1 
      DrawMode        =   8  'Xor Pen
      X1              =   5370
      X2              =   14400
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone NO."
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
      Left            =   7230
      TabIndex        =   34
      Top             =   2070
      Width           =   1815
   End
   Begin VB.Label Label15 
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
      Left            =   14880
      TabIndex        =   29
      Top             =   6870
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Photo"
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
      Left            =   14640
      TabIndex        =   28
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8730
      TabIndex        =   27
      Top             =   5903
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN "
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
      Left            =   8190
      TabIndex        =   25
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label11 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8550
      TabIndex        =   22
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Joining"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label9 
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
      Left            =   10200
      TabIndex        =   15
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label8 
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
      Left            =   5160
      TabIndex        =   14
      Top             =   6810
      Width           =   1935
   End
   Begin VB.Label Label7 
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
      Left            =   330
      TabIndex        =   13
      Top             =   6810
      Width           =   1755
   End
   Begin VB.Label Label6 
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
      Left            =   7680
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label5 
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
      Left            =   270
      TabIndex        =   6
      Top             =   4470
      Width           =   2055
   End
   Begin VB.Label Label4 
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
      Left            =   7800
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   2115
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      Left            =   1350
      TabIndex        =   2
      Top             =   1230
      Width           =   975
   End
End
Attribute VB_Name = "addemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Dim str As String
Private Sub addemp1_Click()
rs.AddNew
rs.Fields("Name").Value = Text1.Text
rs.Fields("Emp_ID").Value = Text2.Text
rs.Fields("DOB").Value = DTPicker1.Value
rs.Fields("DOJ").Value = DTPicker2.Value
If Option1.Value = True Then
rs.Fields("Gender") = Option1.Caption
Else: rs.Fields("Gender") = Option2.Caption
End If
rs.Fields("Designation").Value = Combo1.Text
rs.Fields("Address").Value = Text3.Text
rs.Fields("PIN").Value = Text7.Text
rs.Fields("Basic_Pay").Value = Text4.Text
rs.Fields("Allowance").Value = Text5.Text
rs.Fields("Net_Pay") = Val(Text4.Text) + Val(Text5.Text)
rs.Fields("Calculated").Value = Combo2.Text
rs.Fields("Photo").Value = str
rs.Fields("Phone_NO").Value = Text8.Text
rs.update
MsgBox "Employee Added Successfully!!!..."

End Sub


Private Sub back1_Click()
Home.Show
addemp.Hide
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "JPEG|*jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
End Sub

Private Sub exit1_Click()
If MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo) = vbYes Then
End
Else
Exit Sub
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\Admin\Desktop\My Project\Databases\EmployeeDB.mdb")
Set rs = db.OpenRecordset("select * from empdetails")

DTPicker2.Value = Now

Combo1.AddItem "Accountant"
Combo1.AddItem "Assistant"
Combo1.AddItem "Collection Person"
Combo1.AddItem "Drivier"
Combo1.AddItem "Writer"
Combo2.AddItem "Daily"
Combo2.AddItem "Weekly"
Combo2.AddItem "2 Week Once"
Combo2.AddItem "Monthly"
End Sub

