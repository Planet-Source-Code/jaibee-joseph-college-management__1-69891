VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#1.0#0"; "Candy.ocx"
Object = "{5E8FF3F9-2372-4C96-A258-479E142BF3EF}#1.0#0"; "XP_ProBar.ocx"
Begin VB.Form Frm_SearchStudent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Student Information"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "Frm_SearchStudent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin Candy.CandyButton btnprint 
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      ToolTipText     =   "Print Search Results"
      Top             =   5760
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Print Record"
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
      Top             =   6240
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
      Left            =   6600
      TabIndex        =   6
      ToolTipText     =   "Application Help"
      Top             =   5760
      Width           =   2895
      _ExtentX        =   5106
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
      TabIndex        =   5
      ToolTipText     =   "Search For Information"
      Top             =   5760
      Width           =   2895
      _ExtentX        =   5106
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
      ItemData        =   "Frm_SearchStudent.frx":076A
      Left            =   7200
      List            =   "Frm_SearchStudent.frx":077A
      TabIndex        =   3
      ToolTipText     =   "Enter Search Type"
      Top             =   600
      Width           =   2295
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
      Top             =   600
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frm_SearchStudent.frx":07B3
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Results"
      Top             =   1080
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
      ColumnCount     =   20
      BeginProperty Column00 
         DataField       =   "Prospectus_Number"
         Caption         =   "Prospectus Number"
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
      BeginProperty Column02 
         DataField       =   "Academic_Year"
         Caption         =   "Academic Year"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "Date_of_Birth"
         Caption         =   "Date of Birth"
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
         DataField       =   "Blood_Group"
         Caption         =   "Blood Group"
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
         DataField       =   "Caste"
         Caption         =   "Caste"
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
         DataField       =   "Religion"
         Caption         =   "Religion"
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
         DataField       =   "Nationality"
         Caption         =   "Nationality"
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
      BeginProperty Column11 
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
      BeginProperty Column12 
         DataField       =   "Roll_Number"
         Caption         =   "Roll Number"
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
      BeginProperty Column14 
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
      BeginProperty Column15 
         DataField       =   "Emergency_Contact"
         Caption         =   "Emergency Contact"
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
      BeginProperty Column16 
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
      BeginProperty Column17 
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
      BeginProperty Column18 
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
      BeginProperty Column19 
         DataField       =   "Pic_Student"
         Caption         =   "Picture Path"
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
            ColumnWidth     =   1800
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   9120
      Picture         =   "Frm_SearchStudent.frx":07C8
      ToolTipText     =   "Print Search Results"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   8640
      Picture         =   "Frm_SearchStudent.frx":0F32
      ToolTipText     =   "Application Help"
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
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   240
      Picture         =   "Frm_SearchStudent.frx":169C
      Top             =   0
      Width           =   360
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
      Left            =   5640
      TabIndex        =   2
      Top             =   600
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
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Frm_SearchStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable Declarations
Dim searstr As String
Dim sflag As Boolean
Dim studentrecser As ADODB.Recordset

Private Sub btnhelp_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub btnprint_Click()
On Error Resume Next
Load Frm_PrintStudent
Frm_PrintStudent.Show
Frm_PrintStudent.Label1.Caption = DataGrid1.Columns(1).Text
Frm_PrintStudent.Label8.Caption = DataGrid1.Columns(3).Text
Frm_PrintStudent.Label12.Caption = DataGrid1.Columns(4).Text
Frm_PrintStudent.Label14.Caption = DataGrid1.Columns(5).Text
Frm_PrintStudent.Label15.Caption = DataGrid1.Columns(6).Text
Frm_PrintStudent.Label16.Caption = DataGrid1.Columns(7).Text
Frm_PrintStudent.Label17.Caption = DataGrid1.Columns(8).Text
Frm_PrintStudent.Label18.Caption = DataGrid1.Columns(9).Text
Frm_PrintStudent.Label19.Caption = DataGrid1.Columns(10).Text
Frm_PrintStudent.Label20.Caption = DataGrid1.Columns(11).Text
Frm_PrintStudent.Label21.Caption = DataGrid1.Columns(12).Text
Frm_PrintStudent.Label22.Caption = DataGrid1.Columns(14).Text
Frm_PrintStudent.Label23.Caption = DataGrid1.Columns(15).Text
Frm_PrintStudent.Label24.Caption = DataGrid1.Columns(16).Text
Frm_PrintStudent.Label25.Caption = DataGrid1.Columns(17).Text
Frm_PrintStudent.Label26.Caption = DataGrid1.Columns(18).Text
Frm_PrintStudent.Image1.Picture = LoadPicture(DataGrid1.Columns(19).Text)
Merlin "Print Search Results"
End Sub

Private Sub btnsearch_Click()
' Code For Searching Record
On Error GoTo errlabel

Merlin "Search For Results"
again:
bpbar.Value = 0
If (valuetyp.Text = "All Records" And searchval.Text = "") Then
searstr = "Select * from StudentInformation Order by Admission_Number"
bpbar.Value = 30
ElseIf (valuetyp.Text = "By Name" And searchval.Text <> "") Then
searstr = "Select * from StudentInformation where Student_Name like '" & Trim(searchval.Text) & "%'"
bpbar.Value = 30
ElseIf (valuetyp.Text = "By Admission Number" And searchval.Text <> "") Then
searstr = "Select * from StudentInformation where Admission_Number like '" & Trim(searchval.Text) & "%'"
bpbar.Value = 30
ElseIf (valuetyp.Text = "By Class" And searchval.Text <> "") Then
searstr = "Select * from StudentInformation where Course_Name like '" & Trim(searchval.Text) & "%'"
bpbar.Value = 30
Else
MsgBox "Select Correct Configuration Options", vbInformation, "Error Occured"
Exit Sub
End If

If (sflag = False) Then
studentrecser.Open searstr, studentcon, adOpenStatic, adLockOptimistic
bpbar.Value = 50
Set DataGrid1.DataSource = studentrecser
bpbar.Value = 70
DataGrid1.ReBind
sflag = True
bpbar.Value = 85
Else
sflag = False
studentrecser.Close
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
' Events That Should Happen When Form Is Loaded
On Error Resume Next
Me.Top = 50
Me.Left = 50
Set studentrecser = New ADODB.Recordset
sflag = False
Merlin "Search Student Information Here", "Read"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Events That Should Happen When Form Is Unloaded
On Error Resume Next
If studentrecser.State = adStateOpen Then
studentrecser.Close
End If
End Sub

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub Image3_Click()
On Error Resume Next
btnprint_Click
End Sub

Private Sub searchval_GotFocus()
Merlin "Enter Search Value Here To Search"
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
