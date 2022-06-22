VERSION 5.00
Begin VB.Form tes 
   BackColor       =   &H00808000&
   Caption         =   "Add Product"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "tes.frx":0000
   ScaleHeight     =   9270
   ScaleWidth      =   19605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00570B0F&
      Caption         =   "NOTE:"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6360
      TabIndex        =   18
      Top             =   8040
      Width           =   6375
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Price will be calculated atuomatically from data entered in Cost/ Unit and Quantity"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   960
         TabIndex        =   19
         Top             =   120
         Width           =   5535
      End
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
      Height          =   735
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Product_ID"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   8
      Top             =   2520
      Width           =   2295
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
      Height          =   735
      Left            =   17520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8160
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
      Height          =   735
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton add1 
      BackColor       =   &H00CAA715&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Supplier"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   4
      Top             =   6120
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Price"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Quantity"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   2
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Cost_Per_Unit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   1
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Product_Name"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00570B0F&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   6360
      TabIndex        =   9
      Top             =   1200
      Width           =   6375
      Begin VB.Label id1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID"
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
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label name1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
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
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label cost1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost/ Unit"
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
         Index           =   1
         Left            =   480
         TabIndex        =   14
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label quantity1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Index           =   1
         Left            =   480
         TabIndex        =   13
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label price1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Index           =   1
         Left            =   480
         TabIndex        =   12
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label supplier1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label addproduct1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Product To Inventory"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   1
         Left            =   720
         TabIndex        =   10
         Top             =   120
         Width           =   4935
      End
   End
End
Attribute VB_Name = "tes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public db As Database
Public rs As Recordset
Private Sub add1_Click()
rs.AddNew
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Val(Text3.Text) * Val(Text4.Text)
rs.Fields(5) = Text6.Text
rs.update

MsgBox ("Product Added Successfully")

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub
Private Sub back1_Click()
Home.Show
tes.Hide
End Sub
Private Sub exit1_Click()
End
End Sub
Private Sub Form_Load()
Set db = OpenDatabase("C:\Users\Admin\Desktop\My Project\my stock database.mdb")
Set rs = db.OpenRecordset("select * from inventable")
End Sub

Private Sub reset_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub
