VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form View1 
   BackColor       =   &H00808000&
   Caption         =   "View List of All Items"
   ClientHeight    =   9270
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   19605
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   19605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton defaultview 
      BackColor       =   &H00CAA715&
      Caption         =   "Default View"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8520
      Width           =   2415
   End
   Begin VB.CommandButton sortid 
      BackColor       =   &H00CAA715&
      Caption         =   "Sort ID"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8520
      Width           =   1575
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
      Left            =   17520
      MaskColor       =   &H00400040&
      Style           =   1  'Graphical
      TabIndex        =   2
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
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "View1.frx":0000
      Height          =   7755
      Left            =   720
      TabIndex        =   0
      Top             =   585
      Width           =   18165
      _ExtentX        =   32041
      _ExtentY        =   13679
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16711680
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   29
      FormatLocked    =   -1  'True
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
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "View Items In Inventory"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Product_ID"
         Caption         =   "Product_ID"
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
         DataField       =   "Product_Name"
         Caption         =   "Product_Name"
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
      BeginProperty Column02 
         DataField       =   "Cost_Per_Unit"
         Caption         =   "Cost_Per_Unit"
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
      BeginProperty Column03 
         DataField       =   "Quantity"
         Caption         =   "Quantity"
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
      BeginProperty Column04 
         DataField       =   "Price"
         Caption         =   "Price"
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
      BeginProperty Column05 
         DataField       =   "Supplier"
         Caption         =   "Supplier"
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
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3119.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3119.811
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   2520
      Top             =   6960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Image Image1 
      Height          =   15750
      Left            =   -360
      Picture         =   "View1.frx":0015
      Top             =   -1680
      Width           =   25200
   End
End
Attribute VB_Name = "View1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back1_Click()
Home.Show
View1.Hide
End Sub

Private Sub defaultview_Click()
Adodc1.RecordSource = "Select * from inventable"
Adodc1.Refresh
End Sub

Private Sub exit1_Click()
End
End Sub

Private Sub sortid_Click()
Adodc1.Recordset.Sort = DataGrid1.Columns(Product_ID).DataField
End Sub
