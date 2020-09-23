VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{5E8FF3F9-2372-4C96-A258-479E142BF3EF}#1.0#0"; "XP_ProBar.ocx"
Begin VB.Form Frm_SearchStaffAttendance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Staff Attendance"
   ClientHeight    =   6420
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_SearchStaffAttendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin Candy.CandyButton btnprint 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "Print Search Results"
      Top             =   5520
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
      TabIndex        =   6
      ToolTipText     =   "Search For Data"
      Top             =   5520
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
   Begin VB.ComboBox cmbtype 
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
      ItemData        =   "Frm_SearchStaffAttendance.frx":076A
      Left            =   6720
      List            =   "Frm_SearchStaffAttendance.frx":0777
      TabIndex        =   5
      ToolTipText     =   "Select Search Type"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtval 
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
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Enter Search Value"
      Top             =   840
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frm_SearchStaffAttendance.frx":07A0
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Results"
      Top             =   1320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7223
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
      ColumnCount     =   5
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
         DataField       =   "Staff_Name"
         Caption         =   "Staff Name"
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
         DataField       =   "Department"
         Caption         =   "Department"
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
         DataField       =   "Atn_Date"
         Caption         =   "Attendance Date"
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
         DataField       =   "Attendance_Status"
         Caption         =   "Attendance Status"
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
      EndProperty
   End
   Begin XP_ProBar.UserControl1 bpbar 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6000
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
      Picture         =   "Frm_SearchStaffAttendance.frx":07B5
      ToolTipText     =   "Application Help"
      Top             =   240
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   9120
      Picture         =   "Frm_SearchStaffAttendance.frx":0F1F
      ToolTipText     =   "Print Search Results"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Type"
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Value"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Staff Attendance Records Here"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_SearchStaffAttendance.frx":1689
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "Frm_SearchStaffAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable Declarations
Dim stastr As String
Dim staffatnrec As New ADODB.Recordset
Dim sfflag As Boolean

Private Sub btnprint_Click()
' When Print Button Is Clicked
On Error GoTo message

If staffatnrec.State = adStateOpen Then
Set StaffAttnReport.DataSource = staffatnrec
Load StaffAttnReport
StaffAttnReport.Show
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
' Code For Searching Data
On Error GoTo message

Merlin "Search For Results"
again:
bpbar.Value = 0
If cmbtype.Text = "All Records" And txtval.Text = "" Then
stastr = "Select * from StaffAttendanceInformation order by Serial_Number"
bpbar.Value = 30
ElseIf cmbtype.Text = "By Staff Name" And txtval.Text <> "" Then
stastr = "select * from StaffAttendanceInformation where Staff_Name like '" & Trim$(txtval.Text) & "%'"
bpbar.Value = 30
ElseIf cmbtype.Text = "By Date" And txtval.Text <> "" Then
stastr = "select * from StaffAttendanceInformation where Atn_Date = '" & txtval.Text & "'"
bpbar.Value = 30
Else
MsgBox "Select Correct Options Then Search", vbInformation, "Error Occured"
Exit Sub
End If

If (sfflag = False) Then
staffatnrec.Open stastr, GlobalCon, adOpenStatic, adLockOptimistic
bpbar.Value = 50
Set DataGrid1.DataSource = staffatnrec
bpbar.Value = 70
DataGrid1.ReBind
sfflag = True
bpbar.Value = 85
Else
sfflag = False
staffatnrec.Close
GoTo again
bpbar.Value = 90
End If

bpbar.Value = 100
bpbar.Value = 0

Exit Sub
message:
MsgBox Err.Description, vbCritical, "Error Occured"
End Sub

Private Sub cmbtype_GotFocus()
Merlin "Select One Search Type From Here"
End Sub

Private Sub Form_Load()
On Error Resume Next
' Events That Should Happen When Form Is Loaded
Me.Top = 50
Me.Left = 50
Merlin "You Can Search Staff Attendence Information Here", "Read"
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

Private Sub txtval_GotFocus()
Merlin "Enter One Search Value Here"
End Sub

Private Sub txtval_KeyPress(KeyAscii As Integer)
' When Enter Key Is Pressed
On Error Resume Next
If KeyAscii = 13 Then
btnsearch_Click
End If
End Sub
