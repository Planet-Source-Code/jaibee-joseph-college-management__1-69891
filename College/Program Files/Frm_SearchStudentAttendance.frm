VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{5E8FF3F9-2372-4C96-A258-479E142BF3EF}#1.0#0"; "XP_ProBar.ocx"
Begin VB.Form Frm_SearchStudentAttendance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Student Attendance Here"
   ClientHeight    =   6270
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9600
   Icon            =   "Frm_SearchStudentAttendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox valuetyp 
      BackColor       =   &H00FBE4BF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Frm_SearchStudentAttendance.frx":076A
      Left            =   6960
      List            =   "Frm_SearchStudentAttendance.frx":0777
      TabIndex        =   2
      ToolTipText     =   "Select Search Type"
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox searchval 
      Appearance      =   0  'Flat
      BackColor       =   &H00FBE4BF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Enter Search Value"
      Top             =   840
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frm_SearchStudentAttendance.frx":07AB
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Results"
      Top             =   1320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7011
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
      ColumnCount     =   8
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
         DataField       =   "Student_Class"
         Caption         =   "Student Class"
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
         DataField       =   "Class_Year"
         Caption         =   "Class Year"
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
      BeginProperty Column04 
         DataField       =   "Subject"
         Caption         =   "Subject"
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
      BeginProperty Column05 
         DataField       =   "Month"
         Caption         =   "Month"
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
      BeginProperty Column06 
         DataField       =   "Working_Days"
         Caption         =   "Working Days"
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
      BeginProperty Column07 
         DataField       =   "Days_Present"
         Caption         =   "Days Present"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin Candy.CandyButton btnprint 
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      ToolTipText     =   "Print Search Results"
      Top             =   5400
      Width           =   4455
      _ExtentX        =   7858
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
      Caption         =   "Print"
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
   Begin Candy.CandyButton btnsearch 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Search Information"
      Top             =   5400
      Width           =   4455
      _ExtentX        =   7858
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
      Caption         =   "Search"
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
   Begin XP_ProBar.UserControl1 bpbar 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   6956042
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   8640
      Picture         =   "Frm_SearchStudentAttendance.frx":07C0
      ToolTipText     =   "Application Help"
      Top             =   240
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   9120
      Picture         =   "Frm_SearchStudentAttendance.frx":0F2A
      ToolTipText     =   "Print Search Results"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Student Attendance Information Here"
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
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Picture         =   "Frm_SearchStudentAttendance.frx":1694
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Value Type"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Search Value"
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
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Frm_SearchStudentAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable Declarations Needed
Dim atnsear As String
Dim stuatnsear As New ADODB.Recordset
Dim sfflag As Boolean

Private Sub btnprint_Click()
' When Print Button Is Clicked
On Error GoTo message

If stuatnsear.State = adStateOpen Then
Set StudentAttnReport.DataSource = stuatnsear
Load StudentAttnReport
StudentAttnReport.Show
Merlin "Print Search Results"
Else
MsgBox "Search Again Then Print", vbCritical, "Print Error"
Exit Sub
End If

Exit Sub
message:
MsgBox "Search Again And Print", vbCritical, "Error Occured"
End Sub

Private Sub btnsearch_Click()
' Code For Search Records
On Error GoTo message

Merlin "Click Me To Search For Results"
again:
bpbar.Value = 0
If valuetyp.Text = "All Records" And searchval.Text = "" Then
atnsear = "Select * from StudentAttendanceInformation"
bpbar.Value = 30
ElseIf valuetyp.Text = "By Student Name" And searchval.Text <> "" Then
atnsear = "Select * from StudentAttendanceInformation where Student_Name like '" & Trim$(searchval.Text) & "%'"
bpbar.Value = 30
ElseIf valuetyp.Text = "By Student Class" And searchval.Text <> "" Then
atnsear = "select * from StudentAttendanceInformation where Student_Class like '" & Trim$(searchval.Text) & "%'"
bpbar.Value = 30
Else
MsgBox "Select Correct Configurations", vbInformation, "Error Occured"
Exit Sub
End If

If (sfflag = False) Then
stuatnsear.Open atnsear, studentcon, adOpenStatic, adLockOptimistic
bpbar.Value = 50
Set DataGrid1.DataSource = stuatnsear
bpbar.Value = 70
DataGrid1.ReBind
sfflag = True
bpbar.Value = 85
Else
sfflag = False
stuatnsear.Close
GoTo again
bpbar.Value = 90
End If

bpbar.Value = 100
bpbar.Value = 0

Exit Sub
message:
MsgBox "Select Correct Configurations", vbCritical, "Error Occured"
End Sub

Private Sub Form_Load()
' Events That Should Happen When Form Is Loaded
On Error Resume Next
Me.Top = 50
Me.Left = 50
Merlin "Search Student Attendance Information Here", "Read"
sfflag = False
End Sub

Private Sub Image2_Click()
' When Print Image Is Clicked
On Error Resume Next
btnprint_Click
End Sub

Private Sub Image3_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub searchval_GotFocus()
Merlin "Enter Search Value According To The Search Type"
End Sub

Private Sub searchval_KeyPress(KeyAscii As Integer)
' When Enter Key Is Pressed
On Error Resume Next
If KeyAscii = 13 Then
btnsearch_Click
End If
End Sub

Private Sub valuetyp_GotFocus()
Merlin "Select One Search Type From Here"
End Sub
