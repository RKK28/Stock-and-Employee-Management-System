VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form saldetails 
   BackColor       =   &H00400040&
   BorderStyle     =   0  'None
   Caption         =   "Salary Paid Details"
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19755
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "saldetails.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   19755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3960
      TabIndex        =   31
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton showall1 
      BackColor       =   &H00CAA715&
      Caption         =   "Show All Employee"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6105
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8520
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
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
      Left            =   16830
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   16590
      Locked          =   -1  'True
      MousePointer    =   12  'No Drop
      TabIndex        =   26
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   13950
      Locked          =   -1  'True
      MousePointer    =   12  'No Drop
      TabIndex        =   21
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   11310
      Locked          =   -1  'True
      MousePointer    =   12  'No Drop
      TabIndex        =   20
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1095
      Left            =   8790
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   3120
      Width           =   7575
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
      Height          =   495
      Left            =   12195
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   9675
      TabIndex        =   15
      Top             =   8480
      Width           =   2415
   End
   Begin VB.CommandButton submit1 
      BackColor       =   &H00CAA715&
      Caption         =   "Submit"
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
      Left            =   16830
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   1560
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   $"saldetails.frx":6CB40
      OLEDBString     =   $"saldetails.frx":6CBC7
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from saldetails"
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
   Begin VB.ComboBox Combo2 
      BackColor       =   &H80000007&
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
      Left            =   6030
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3120
      Width           =   2415
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
      Height          =   495
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8520
      Width           =   1215
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
      Height          =   495
      Left            =   16680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   3390
      TabIndex        =   9
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   750
      TabIndex        =   8
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   6030
      Locked          =   -1  'True
      MousePointer    =   12  'No Drop
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000007&
      DataField       =   "Emp_ID"
      DataSource      =   "Adodc2"
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
      Left            =   3390
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   585
      Left            =   750
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1032
      _Version        =   393216
      MousePointer    =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483641
      CalendarForeColor=   -2147483637
      DateIsNull      =   -1  'True
      Format          =   113246209
      CurrentDate     =   44334
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   585
      Left            =   8670
      Locked          =   -1  'True
      MousePointer    =   12  'No Drop
      TabIndex        =   19
      Top             =   1560
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "saldetails.frx":6CC4E
      Height          =   3855
      Left            =   600
      TabIndex        =   13
      Top             =   4440
      Width           =   18495
      _ExtentX        =   32623
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   -2147483643
      HeadLines       =   1
      RowHeight       =   26
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Date"
         Caption         =   "Date"
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
         DataField       =   "Name"
         Caption         =   "Name"
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
         DataField       =   "Emp_ID"
         Caption         =   "Emp_ID"
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
         DataField       =   "Paid"
         Caption         =   "Paid"
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
         DataField       =   "Balance"
         Caption         =   "Balance"
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
         DataField       =   "Mode"
         Caption         =   "Mode"
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
      BeginProperty Column06 
         DataField       =   "Comments"
         Caption         =   "Comments"
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
            WrapText        =   -1  'True
            ColumnWidth     =   2910.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column06 
            WrapText        =   -1  'True
            ColumnWidth     =   4995.213
         EndProperty
      EndProperty
   End
   Begin WMPLibCtl.WindowsMediaPlayer play 
      Height          =   855
      Left            =   360
      TabIndex        =   32
      Top             =   8400
      Visible         =   0   'False
      Width           =   3255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5741
      _cy             =   1508
   End
   Begin VB.Line Line12 
      BorderStyle     =   3  'Dot
      DrawMode        =   16  'Merge Pen
      X1              =   480
      X2              =   19200
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line11 
      BorderStyle     =   3  'Dot
      DrawMode        =   16  'Merge Pen
      X1              =   480
      X2              =   19200
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line10 
      BorderStyle     =   3  'Dot
      DrawMode        =   16  'Merge Pen
      X1              =   6000
      X2              =   13800
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line9 
      DrawMode        =   16  'Merge Pen
      X1              =   14160
      X2              =   19200
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line8 
      DrawMode        =   16  'Merge Pen
      X1              =   19200
      X2              =   19200
      Y1              =   8400
      Y2              =   360
   End
   Begin VB.Line Line7 
      DrawMode        =   16  'Merge Pen
      X1              =   13800
      X2              =   19200
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line6 
      DrawMode        =   16  'Merge Pen
      X1              =   13800
      X2              =   13800
      Y1              =   9120
      Y2              =   8400
   End
   Begin VB.Line Line5 
      DrawMode        =   16  'Merge Pen
      X1              =   6000
      X2              =   13800
      Y1              =   9120
      Y2              =   9120
   End
   Begin VB.Line Line4 
      DrawMode        =   16  'Merge Pen
      X1              =   6000
      X2              =   6000
      Y1              =   8400
      Y2              =   9120
   End
   Begin VB.Line Line3 
      DrawMode        =   16  'Merge Pen
      X1              =   480
      X2              =   6000
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line2 
      DrawMode        =   16  'Merge Pen
      X1              =   480
      X2              =   480
      Y1              =   360
      Y2              =   8400
   End
   Begin VB.Line Line1 
      DrawMode        =   16  'Merge Pen
      X1              =   480
      X2              =   5640
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update Salary Details And View  Details In Database"
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
      Left            =   5565
      TabIndex        =   28
      Top             =   120
      Width           =   8625
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   750
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label12 
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
      Left            =   16560
      TabIndex        =   25
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label11 
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
      Left            =   13920
      TabIndex        =   24
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label10 
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
      Left            =   11310
      TabIndex        =   23
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Baisc Pay"
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
      TabIndex        =   22
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
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
      Left            =   8790
      TabIndex        =   17
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mode"
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
      Left            =   6030
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
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
      Left            =   3390
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid"
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
      Left            =   750
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      Left            =   3390
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "EMP ID"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "saldetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub back1_Click()
Home.Show
saldetails.Hide
End Sub

Private Sub Combo1_Click()
Set rs = New ADODB.Recordset
rs.Open "select * from empdetails where Name = '" & Combo1.Text & "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
Text1.Text = rs!Emp_ID
Text6.Text = rs!Basic_Pay
Text7.Text = rs!Allowance
Text8.Text = rs!Net_pay
Text9.Text = rs!Calculated
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub Command1_Click()
resetfields
End Sub

Private Sub exit1_Click()
If MsgBox("Are You Sure You Want To Exit, This Lead To LogOut And Any Unsaved Changes Will Not Be Saved?", vbExclamation + vbYesNo) = vbYes Then
End
Else
Exit Sub
End If
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Desktop\My Project\Databases\EmployeeDB.mdb;Persist Security Info= False"
rs.Open "select * from empdetails", con, adOpenDynamic, adLockPessimistic
Do While Not rs.EOF
Combo1.AddItem rs!Name
rs.MoveNext
Loop
Set rs = Nothing

DTPicker1.Value = Now

Combo2.AddItem ("Bank via UPI")
Combo2.AddItem ("Cash")

Adodc1.Refresh

End Sub


Private Sub search1_Click()
Adodc1.Recordset.Close
Adodc1.RecordSource = "select * from saldetails where Emp_ID = '" + Text4.Text + "'or Name = '" + Text4.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Not found", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If

End Sub

Private Sub showall1_Click()
Adodc1.Recordset.Close
Adodc1.RecordSource = "select * from saldetails"
Adodc1.Refresh
End Sub

Private Sub submit1_Click()
If Combo1.ListIndex = -1 Then
MsgBox "Select Employee Name!!!", vbExclamation
ElseIf Text2.Text = "" Then
MsgBox "Fill Paid Amount!!!", vbExclamation
ElseIf Text3.Text = "" Then
MsgBox "Fill Balance Amount!!!", vbExclamation
ElseIf Combo2.ListIndex = -1 Then
MsgBox "Select Mode Of Payment!!!", vbExclamation
ElseIf Text5.Text = "" Then
MsgBox "Enter Comments!!!", vbExclamation
Else
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = DTPicker1.Value
Adodc1.Recordset.Fields(1) = Combo1.Text
Adodc1.Recordset.Fields(2) = Text1.Text
Adodc1.Recordset.Fields(3) = Text2.Text
Adodc1.Recordset.Fields(4) = Text3.Text
Adodc1.Recordset.Fields(5) = Combo2.Text
Adodc1.Recordset.Fields(6) = Text5.Text
Adodc1.Recordset.update
MsgBox "Salary Details Updated Successfullly!!!...", vbInformation
resetfields
End If
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
Combo1.ListIndex = -1
Combo2.ListIndex = -1
End Sub
