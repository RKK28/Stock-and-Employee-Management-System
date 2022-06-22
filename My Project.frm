VERSION 5.00
Begin VB.Form Home 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19755
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "My Project.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   19755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton viewemp1 
      BackColor       =   &H00CAA715&
      Caption         =   "View Emphloyee"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   12270
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CommandButton delemp1 
      BackColor       =   &H00CAA715&
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
      Height          =   975
      Index           =   4
      Left            =   12270
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5363
      Width           =   2655
   End
   Begin VB.CommandButton searchemp1 
      BackColor       =   &H00CAA715&
      Caption         =   "Searh Employee"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   15030
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3203
      Width           =   2655
   End
   Begin VB.CommandButton updateemp1 
      BackColor       =   &H00CAA715&
      Caption         =   "Update Employee"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   15030
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4283
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CAA715&
      Caption         =   "Salary Paid Details"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   15030
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5363
      Width           =   2655
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
      Height          =   975
      Index           =   0
      Left            =   12270
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3203
      Width           =   2655
   End
   Begin VB.CommandButton View 
      BackColor       =   &H00CAA715&
      Caption         =   "View Items"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3083
      Width           =   2655
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
      Height          =   975
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8160
      Width           =   2655
   End
   Begin VB.CommandButton modify1 
      BackColor       =   &H00CAA715&
      Caption         =   "Modify Item"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4163
      Width           =   2655
   End
   Begin VB.CommandButton searchitem1 
      BackColor       =   &H00CAA715&
      Caption         =   "Search Item"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5243
      Width           =   2655
   End
   Begin VB.CommandButton delete1 
      BackColor       =   &H00CAA715&
      Caption         =   "Delete Item"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5243
      Width           =   2655
   End
   Begin VB.CommandButton additem1 
      BackColor       =   &H00CAA715&
      Caption         =   "Add Item"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4163
      Width           =   2655
   End
   Begin VB.CommandButton createbill1 
      BackColor       =   &H00CAA715&
      Caption         =   "Create Bill"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3083
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8520
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Details"
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
      Left            =   13920
      TabIndex        =   15
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Maintainence"
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
      Left            =   3240
      TabIndex        =   14
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addemp1_Click(Index As Integer)
addemp.Show
Home.Hide
End Sub
Private Sub additem1_Click()
tes.Show
Home.Hide
End Sub

Private Sub Command1_Click(Index As Integer)
saldetails.Show
Home.Hide
End Sub

Private Sub delemp1_Click(Index As Integer)
delemp.Show
Home.Hide
End Sub

Private Sub delete1_Click()
deleteprod.Show
Home.Hide
End Sub
Private Sub exit1_Click()
End
End Sub

Private Sub modify1_Click()
update.Show
Home.Hide
End Sub

Private Sub searchemp1_Click(Index As Integer)
searchemp.Show
Home.Hide
End Sub

Private Sub updateemp1_Click(Index As Integer)
updateemp.Show
Home.Hide
End Sub

Private Sub View_Click()
View1.Show
Home.Hide
End Sub

Private Sub viewemp1_Click(Index As Integer)
viewemp.Show
Home.Hide
End Sub
