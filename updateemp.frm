VERSION 5.00
Begin VB.Form updateemp 
   Caption         =   "Update Employee Details"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   LinkTopic       =   "Form1"
   Picture         =   "updateemp.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   19755
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   47
      Top             =   8520
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
      TabIndex        =   46
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton salupt 
      BackColor       =   &H00CAA715&
      Caption         =   "Update"
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
      Left            =   13950
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton addupt 
      BackColor       =   &H00CAA715&
      Caption         =   "Update"
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
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton phnupt 
      BackColor       =   &H00CAA715&
      Caption         =   "Update"
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
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox Text19 
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
      Left            =   13950
      TabIndex        =   42
      Top             =   7320
      Width           =   2295
   End
   Begin VB.TextBox Text18 
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
      Left            =   13950
      TabIndex        =   41
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox Text17 
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
      Left            =   13950
      TabIndex        =   40
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox Text16 
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
      Left            =   9270
      TabIndex        =   39
      Top             =   7680
      Width           =   2295
   End
   Begin VB.TextBox Text15 
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
      Height          =   1655
      Left            =   9270
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox Text14 
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
      Left            =   4470
      TabIndex        =   37
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton last1 
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
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton next1 
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
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton prev1 
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4440
      Width           =   1575
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text13 
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
      Left            =   15630
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text12 
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
      Left            =   15630
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text11 
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
      Left            =   15630
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text10 
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
      Left            =   15630
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text9 
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
      Left            =   11070
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2880
      Width           =   2295
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
      Height          =   1655
      Left            =   11070
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text7 
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
      Left            =   11070
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3600
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   7230
      ScaleHeight     =   2355
      ScaleWidth      =   1875
      TabIndex        =   12
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text6 
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
      Left            =   4590
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text5 
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
      Left            =   4590
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text4 
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
      Left            =   4590
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text3 
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
      Left            =   4590
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
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
      Height          =   585
      Left            =   4590
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
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
      Left            =   4590
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      Height          =   3615
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   5655
   End
   Begin VB.Line Line5 
      X1              =   18120
      X2              =   18120
      Y1              =   360
      Y2              =   5280
   End
   Begin VB.Line Line4 
      X1              =   11880
      X2              =   18120
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   1680
      X2              =   18120
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line2 
      X1              =   1680
      X2              =   1680
      Y1              =   360
      Y2              =   5280
   End
   Begin VB.Line Line1 
      X1              =   7800
      X2              =   1680
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Employee Profiles"
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
      Left            =   7710
      TabIndex        =   48
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label19 
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   12030
      TabIndex        =   36
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label18 
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   12030
      TabIndex        =   35
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label17 
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   12030
      TabIndex        =   34
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label16 
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
      Left            =   7800
      TabIndex        =   33
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label15 
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
      Left            =   7680
      TabIndex        =   32
      Top             =   6000
      Width           =   1335
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1920
      TabIndex        =   31
      Top             =   5880
      Width           =   2415
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   13710
      TabIndex        =   19
      Top             =   3120
      Width           =   1695
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   13710
      TabIndex        =   18
      Top             =   2400
      Width           =   1695
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   13710
      TabIndex        =   17
      Top             =   1680
      Width           =   1695
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   13710
      TabIndex        =   16
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label9 
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
      Left            =   9510
      TabIndex        =   15
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label8 
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
      Left            =   9510
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   8430
      TabIndex        =   13
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desgination"
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
      Left            =   1830
      TabIndex        =   5
      Top             =   4560
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1830
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
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
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1830
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
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
      Left            =   1830
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
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
      Left            =   1830
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
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
      Height          =   495
      Left            =   1830
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.Shape Shape3 
      Height          =   3255
      Left            =   11880
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   4575
   End
End
Attribute VB_Name = "updateemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Dim str As String

Private Sub addupt_Click()
If Text15.Text = "" Or Text16.Text = "" Then
MsgBox "Please Enter All Fields!!!"
Else
rs.Edit
rs.Fields("Address").Value = Text15.Text
rs.Fields("PIN").Value = Text16.Text
rs.update
MsgBox "Address And PIN Updated Successfully!!!..."
Text15.Text = ""
Text16.Text = ""
End If
End Sub

Private Sub back1_Click()
Home.Show
updateemp.Hide
End Sub

Private Sub exit1_Click()
End
End Sub

Private Sub first1_Click()
rs.MoveFirst
display
Picture1.Picture = LoadPicture(rs!Photo)
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\Admin\Desktop\My Project\Databases\EmployeeDB.mdb")
Set rs = db.OpenRecordset("select * from empdetails")
End Sub

Private Sub last1_Click()
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

Private Sub phnupt_Click()
If Text14.Text = "" Then
MsgBox "Please Enter Phone Number!!!"
Else
rs.Edit
rs.Fields("Phone_No").Value = Text14.Text
MsgBox "Phone Number Updated Successfully!!!..."
rs.update
Text14.Text = ""
End If
End Sub

Private Sub prev1_Click()
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
Sub display()
Text1.Text = rs!Name
Text2.Text = rs!Emp_ID
Text3.Text = rs!DOB
Text4.Text = rs!DOJ
Text5.Text = rs!Gender
Text6.Text = rs!Designation
Text7.Text = rs!Phone_No
Text8.Text = rs!Address
Text9.Text = rs!PIN
Text10.Text = rs!Basic_Pay
Text11.Text = rs!Allowance
Text13.Text = rs!Net_pay
Text12.Text = rs!Calculated
End Sub

Private Sub salupt_Click()
If Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Then
MsgBox "Please Fill All Details!!!"
Else
rs.Edit
rs.Fields("Basic_Pay").Value = Text17.Text
rs.Fields("Allowance").Value = Text18.Text
rs.Fields("Calculated").Value = Text19.Text
rs.Fields("Net_Pay").Value = Val(Text17.Text) + Val(Text18.Text)
rs.update
MsgBox "Basic-Pay, Allowance, Calculated-Term and Basic-Pay Updated Successfully!!!..."
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
End If
End Sub
