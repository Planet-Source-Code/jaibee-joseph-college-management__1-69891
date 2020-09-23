VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#1.0#0"; "Candy.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frm_GlobalExamEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Global Exam Type Entry"
   ClientHeight    =   5190
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6990
   Icon            =   "Frm_GlobalMarkEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2160
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=College"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "College"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "MarkGlobalInformation"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frm_GlobalMarkEntry.frx":076A
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enter Exam Type Entry Here"
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16508095
      HeadLines       =   1
      RowHeight       =   18
      TabAction       =   1
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Global Exam Entry"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Serial_Number"
         Caption         =   "Serial Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Exam_Type"
         Caption         =   "Exam Type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Max_Mark"
         Caption         =   "Max Mark"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Min_Mark"
         Caption         =   "Min Mark"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin Candy.CandyButton btnhelp 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Application Help"
      Top             =   4680
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Help"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   15309136
      ColorButtonUp   =   13657888
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   1
   End
   Begin Candy.CandyButton btnrefresh 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Update All Entry"
      Top             =   4200
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Update All Entry"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   15309136
      ColorButtonUp   =   13657888
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   1
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   6480
      Picture         =   "Frm_GlobalMarkEntry.frx":077F
      ToolTipText     =   "Application Help"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Global Examination Type Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_GlobalMarkEntry.frx":0EE9
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "Frm_GlobalExamEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnhelp_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub btnrefresh_Click()
' Close And Reestablish Connection
On Error Resume Next
Call MainConClose
Call MainConEstablish
Merlin "Update All Entry To Form"
End Sub

Private Sub Form_Load()
' When Form Loads
On Error Resume Next
Me.Top = 50
Me.Left = 50
Merlin "Enter Global Entry For Exam Here", "Read"
End Sub

Private Sub Image2_Click()
btnhelp_Click
End Sub
