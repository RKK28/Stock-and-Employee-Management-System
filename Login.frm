VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Login 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19755
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   19755
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc loginado 
      Height          =   735
      Left            =   8520
      Top             =   4343
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Desktop\My Project\Databases\Logindb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Desktop\My Project\Databases\Logindb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from logintab"
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
   Begin VB.CommandButton Command1 
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
      Left            =   17400
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   1815
   End
   Begin VB.TextBox password1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "password"
      DataSource      =   "loginado"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox username1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "username"
      DataSource      =   "loginado"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton signin 
      BackColor       =   &H00CAA715&
      Caption         =   "Sign-In"
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
      Left            =   2280
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label username 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "UserName"
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
      Left            =   960
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Delveloped By: Karan Kumar"
      BeginProperty Font 
         Name            =   "Sitka Subheading"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   12840
      TabIndex        =   7
      Top             =   8400
      Width           =   6615
   End
   Begin VB.Label password 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label stockmaintanence 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Maintanence"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   16200
      Left            =   -8400
      Picture         =   "Login.frx":0000
      Top             =   -3000
      Width           =   28800
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()

End Sub


Private Sub Command1_Click()
exit1.Show


End Sub

Private Sub signin_Click()
loginado.RecordSource = "select * from logintab where username='" + username1.Text + "' and password='" + password1.Text + "'"
loginado.Refresh
If loginado.Recordset.EOF Then
MsgBox "Login Failed Try Again...!!!", vbCritical, "Please Enter Correct Username and Password"
Else
MsgBox "Login Successfull...!!!", vbInformation, "Successfull Attempt"
Home.Show
Login.Hide
End If
End Sub

