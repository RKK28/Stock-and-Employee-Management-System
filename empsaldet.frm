VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form empsaldet 
   Caption         =   "See Employee Salary Paid Details"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   19755
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   735
      Left            =   8040
      TabIndex        =   10
      Top             =   2640
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1296
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      _Version        =   393216
      Format          =   108855297
      CurrentDate     =   44331
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Paid On"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Paid"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Employee ID"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "See History"
      Height          =   495
      Left            =   12960
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Update Details"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "empsaldet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
