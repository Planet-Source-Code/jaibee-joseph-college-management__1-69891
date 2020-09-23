VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{5E8FF3F9-2372-4C96-A258-479E142BF3EF}#1.0#0"; "XP_ProBar.ocx"
Begin VB.Form Frm_SearchStaff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Staff Information"
   ClientHeight    =   6705
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9615
   Icon            =   "Frm_SearchStaff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   Begin Candy.CandyButton btnprint 
      Height          =   375
      Left            =   3300
      TabIndex        =   9
      ToolTipText     =   "Print Search Results"
      Top             =   5880
      Width           =   3015
      _ExtentX        =   5318
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
      TabIndex        =   4
      ToolTipText     =   "Enter Search Value"
      Top             =   720
      Width           =   2295
   End
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
      ItemData        =   "Frm_SearchStaff.frx":076A
      Left            =   7200
      List            =   "Frm_SearchStaff.frx":0777
      TabIndex        =   3
      ToolTipText     =   "Select Search Type"
      Top             =   720
      Width           =   2295
   End
   Begin XP_ProBar.UserControl1 bpbar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6360
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
   Begin Candy.CandyButton BtnHelp 
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      ToolTipText     =   "Application Help"
      Top             =   5880
      Width           =   3015
      _ExtentX        =   5318
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
   Begin Candy.CandyButton BtnSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Search For Data"
      Top             =   5880
      Width           =   3015
      _ExtentX        =   5318
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frm_SearchStaff.frx":079E
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Search Results"
      Top             =   1200
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8070
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
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "Staff_ID"
         Caption         =   "Staff ID"
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
         DataField       =   "Sex"
         Caption         =   "Sex"
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
         DataField       =   "Qualification"
         Caption         =   "Qualification"
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
         DataField       =   "Designation"
         Caption         =   "Designation"
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
      BeginProperty Column06 
         DataField       =   "Subjects_Handled"
         Caption         =   "Subjects Handled"
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
         DataField       =   "Salary_Month"
         Caption         =   "Salary/Month"
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
      BeginProperty Column08 
         DataField       =   "Temporary_Address"
         Caption         =   "Temporary Address"
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
      BeginProperty Column09 
         DataField       =   "Permanent_Address"
         Caption         =   "Permanent Address"
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
      BeginProperty Column10 
         DataField       =   "Phone_Number"
         Caption         =   "Phone Number"
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
      BeginProperty Column11 
         DataField       =   "Mobile_Number"
         Caption         =   "Mobile Number"
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
      BeginProperty Column12 
         DataField       =   "EMail_Address"
         Caption         =   "E-Mail Address"
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
      BeginProperty Column13 
         DataField       =   "Pic_Staff"
         Caption         =   "Staff Picture Path"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   9120
      Picture         =   "Frm_SearchStaff.frx":07B3
      ToolTipText     =   "Print Search Results"
      Top             =   240
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   8640
      Picture         =   "Frm_SearchStaff.frx":0F1D
      ToolTipText     =   "Application Help"
      Top             =   240
      Width           =   360
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
      TabIndex        =   8
      Top             =   720
      Width           =   1455
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
      Left            =   5760
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_SearchStaff.frx":1687
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select One Search Type, Enter Value Of That Type And Press Search Button For Results."
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
      TabIndex        =   6
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "Frm_SearchStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable Declarations
Dim searstr As String
Dim sflag As Boolean
Dim staffrecser As ADODB.Recordset

Private Sub btnhelp_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub btnprint_Click()
On Error GoTo message

If staffrecser.State = adStateOpen Then
Set StaffDetailReport.DataSource = staffrecser
Load StaffDetailReport
StaffDetailReport.Show
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
' Code For Searching Record
On Error GoTo errlabel

Merlin "Search For Results"
again:
bpbar.Value = 0
If (valuetyp.Text = "All Records" And searchval.Text = "") Then
searstr = "Select * from StaffInformation Order by Staff_ID"
bpbar.Value = 30
ElseIf (valuetyp.Text = "By Name" And searchval.Text <> "") Then
searstr = "Select * from StaffInformation where Staff_Name like '" & Trim(searchval.Text) & "%'"
bpbar.Value = 30
ElseIf (valuetyp.Text = "By Staff ID" And searchval.Text <> "") Then
searstr = "Select * from StaffInformation where Staff_ID like '" & Trim(searchval.Text) & "%'"
bpbar.Value = 30
Else
MsgBox "Select Correct Options Then Search", vbInformation, "Error Occured"
Exit Sub
End If

If (sflag = False) Then
staffrecser.Open searstr, staffcon, adOpenStatic, adLockOptimistic
bpbar.Value = 50
Set DataGrid1.DataSource = staffrecser
bpbar.Value = 70
DataGrid1.ReBind
sflag = True
bpbar.Value = 85
Else
sflag = False
staffrecser.Close
GoTo again
bpbar.Value = 90
End If

bpbar.Value = 100
bpbar.Value = 0

Exit Sub
errlabel:
MsgBox Err.Description, vbCritical, "Error Occured"
bpbar.Value = 0
End Sub

Private Sub Form_Load()
' When Form Is Loaded
On Error Resume Next
Me.Top = 50
Me.Left = 50
Set staffrecser = New ADODB.Recordset
sflag = False
Merlin "You Can Search Staff Information Here", "Read"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' When Form Is Unloaded
On Error Resume Next
If staffrecser.State = adStateOpen Then
staffrecser.Close
End If
End Sub

Private Sub Image2_Click()
btnhelp_Click
End Sub

Private Sub Image3_Click()
btnprint_Click
End Sub

Private Sub searchval_GotFocus()
Merlin "Enter One Search Value Here"
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
