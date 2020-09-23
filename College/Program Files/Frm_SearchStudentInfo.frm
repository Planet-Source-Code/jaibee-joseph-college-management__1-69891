VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Begin VB.Form Frm_SearchStudentInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Student Information"
   ClientHeight    =   4215
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_SearchStudentInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FBE4BF&
      Height          =   330
      Left            =   1680
      TabIndex        =   4
      ToolTipText     =   "Enter Student Name Here For Search"
      Top             =   720
      Width           =   3735
   End
   Begin Candy.CandyButton btnCancel 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "Unload Form"
      Top             =   3720
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Close"
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
   Begin Candy.CandyButton btnOK 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Press Me To Take Data To Form"
      Top             =   3720
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "OK"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frm_SearchStudentInfo.frx":08CA
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Search Results"
      Top             =   1200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16508095
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
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
      Caption         =   "Search Results"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Admission_Number"
         Caption         =   "Admission Number"
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
         DataField       =   "Student_Name"
         Caption         =   "Student Name"
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
         DataField       =   "Course_Name"
         Caption         =   "Course Name"
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
         DataField       =   "Year_Course"
         Caption         =   "Year Course"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Student Name"
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
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Student Information"
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
      Top             =   240
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   120
      Picture         =   "Frm_SearchStudentInfo.frx":08DF
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Frm_SearchStudentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable Declarations Needed For Code
Dim stusea As ADODB.Recordset
Dim strstu As String
Private Sub btnCancel_Click()
' Unload form when cancel button is pressed
On Error Resume Next
Unload Me
End Sub

Private Sub btnOK_Click()
' Load Data To Form
On Error Resume Next
Frm_StudentEntry.AdmissionNumber.Text = DataGrid1.Columns(0).Text
Frm_StudentEntry.StudentName.Text = DataGrid1.Columns(1).Text
Frm_StudentEntry.cmbcourse.Text = DataGrid1.Columns(2).Text
Frm_StudentEntry.cmbyear.Text = DataGrid1.Columns(3).Text
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
' Events That Should Happen When Form Is Loaded
Me.Top = 700
Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
' Code For Searching Record
On Error Resume Next
If KeyAscii = 13 Then
strstu = "Select Admission_Number, Student_Name, Course_Name, Year_Course from StudentInformation where Student_Name like '" & Trim$(Text1.Text) & "%'"
Set stusea = New ADODB.Recordset
stusea.Open strstu, studentcon, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = stusea
DataGrid1.ReBind
End If
End Sub
