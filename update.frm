VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form update 
   BackColor       =   &H00808000&
   Caption         =   "Update Product"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   19605
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   -840
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Desktop\My Project\MY STOCK DATABASE.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Desktop\My Project\MY STOCK DATABASE.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from inventable"
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
   Begin VB.TextBox Text4 
      DataField       =   "Quantity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   7
      Top             =   8160
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "Cost/ Unit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   8160
      Width           =   1935
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
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
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
      Left            =   15840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton update1 
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "update.frx":0000
      Height          =   5895
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   17655
      _ExtentX        =   31141
      _ExtentY        =   10398
      _Version        =   393216
      BackColor       =   16744448
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   27
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Modify Items In Inventory"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Note 
      BackColor       =   &H00CAA715&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   960
      TabIndex        =   8
      Top             =   6600
      Width           =   17655
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ii) To Update or Modify ""Product ID, Product Name or Supplier"", then Delete the Product and Add again."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   840
         Width           =   15975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"update.frx":0015
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   10
         Top             =   120
         Width           =   16335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404000&
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
      Height          =   615
      Left            =   5280
      TabIndex        =   5
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
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
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   21600
      Left            =   -2400
      Picture         =   "update.frx":00BC
      Top             =   -6240
      Width           =   38400
   End
End
Attribute VB_Name = "update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back1_Click()
Home.Show
update.Hide
End Sub

Private Sub exit1_Click()
End
End Sub
Private Sub update1_Click()
With Adodc1.Recordset
.update
!Cost_Per_Unit = Text3
!Quantity = Text4
End With
update1.Refresh
End Sub
