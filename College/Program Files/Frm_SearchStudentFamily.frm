VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{5E8FF3F9-2372-4C96-A258-479E142BF3EF}#1.0#0"; "XP_ProBar.ocx"
Begin VB.Form Frm_SearchStudentFamily 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Student Family Information"
   ClientHeight    =   5895
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9630
   Icon            =   "Frm_SearchStudentFamily.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin Candy.CandyButton btnprint 
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      ToolTipText     =   "Print Search Results"
      Top             =   5040
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
      TabIndex        =   2
      ToolTipText     =   "Enter Search Value Here"
      Top             =   600
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
      ItemData        =   "Frm_SearchStudentFamily.frx":076A
      Left            =   7200
      List            =   "Frm_SearchStudentFamily.frx":077A
      TabIndex        =   1
      ToolTipText     =   "Select One Search Type"
      Top             =   600
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frm_SearchStudentFamily.frx":07BE
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Search Results"
      Top             =   1080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6800
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
      Caption         =   "Search Resullts"
      ColumnCount     =   9
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "Fathers_Name"
         Caption         =   "Father's Name"
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
         DataField       =   "Fathers_Occupation"
         Caption         =   "Father's Occupation"
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
         DataField       =   "FMobile_Number"
         Caption         =   "Father's Mobile Number"
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
         DataField       =   "FOffice_Number"
         Caption         =   "Father's Office Number"
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
         DataField       =   "Mothers_Name"
         Caption         =   "Mother's Name"
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
         DataField       =   "Mothers_Occupation"
         Caption         =   "Mother's Occupation"
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
         DataField       =   "MMobile_Number"
         Caption         =   "Mother's Mobile Number"
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
         DataField       =   "MOffice_Number"
         Caption         =   "Mother's Office Number"
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
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin XP_ProBar.UserControl1 bpbar 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5520
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
      TabIndex        =   7
      ToolTipText     =   "Application Help"
      Top             =   5040
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
      TabIndex        =   8
      ToolTipText     =   "Click Me To Search"
      Top             =   5040
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
   Begin VB.Image Image3 
      Height          =   360
      Left            =   9120
      Picture         =   "Frm_SearchStudentFamily.frx":07D3
      ToolTipText     =   "Print Search Results"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   8640
      Picture         =   "Frm_SearchStudentFamily.frx":0F3D
      ToolTipText     =   "Application Help"
      Top             =   120
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
      TabIndex        =   5
      Top             =   600
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
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_SearchStudentFamily.frx":16A7
      Top             =   0
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
      TabIndex        =   3
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Frm_SearchStudentFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable Declarations Needed For Code
Dim searftr As ADODB.Recordset
Dim seaf As String
Dim sfflag As Boolean

Private Sub btnhelp_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub btnprint_Click()
On Error GoTo message

If searftr.State = adStateOpen Then
Set StudentFamilyReport.DataSource = searftr
Load StudentFamilyReport
StudentFamilyReport.Show
Merlin "Print Search Results"
Else
MsgBox "Search Again And Then Print", vbCritical, "Print Error"
Exit Sub
End If

Exit Sub
message:
MsgBox "Search Again And Print", vbCritical, "Error Occured"
End Sub

Private Sub btnsearch_Click()
' Code For Searching Record
On Error GoTo errmess

Merlin "Search For Results"
again:
bpbar.Value = 0
If (valuetyp.Text = "All Records" And searchval.Text = "") Then
seaf = "Select all Student_Name, Fathers_Name, Fathers_Occupation, FMobile_Number, FOffice_Number, Mothers_Name, Mothers_Occupation, MMobile_Number, MOffice_Number from StudentInformation, FamilyInformation where StudentInformation.Admission_Number=FamilyInformation.Admission_Number"
bpbar.Value = 30
ElseIf (valuetyp.Text = "By Student Name" And searchval.Text <> "") Then
seaf = "Select Student_Name, Fathers_Name, Fathers_Occupation, FMobile_Number, FOffice_Number, Mothers_Name, Mothers_Occupation, MMobile_Number, MOffice_Number from StudentInformation, FamilyInformation where StudentInformation.Admission_Number=FamilyInformation.Admission_Number and Student_Name like '" & Trim$(searchval.Text) & "%'"
bpbar.Value = 30
ElseIf (valuetyp.Text = "By Fathers Name" And searchval.Text <> "") Then
seaf = "Select all Student_Name, Fathers_Name, Fathers_Occupation, FMobile_Number, FOffice_Number, Mothers_Name, Mothers_Occupation, MMobile_Number, MOffice_Number from StudentInformation, FamilyInformation where StudentInformation.Admission_Number=FamilyInformation.Admission_Number and Fathers_Name like '" & Trim$(searchval.Text) & "%'"
bpbar.Value = 30
ElseIf (valuetyp.Text = "By Mothers Name" And searchval.Text <> "") Then
seaf = "Select all Student_Name, Fathers_Name, Fathers_Occupation, FMobile_Number, FOffice_Number, Mothers_Name, Mothers_Occupation, MMobile_Number, MOffice_Number from StudentInformation, FamilyInformation where StudentInformation.Admission_Number=FamilyInformation.Admission_Number and Mothers_Name like '" & Trim$(searchval.Text) & "%'"
bpbar.Value = 30
Else
MsgBox "Select Correct Search Options", vbInformation, "Correct Search Options"
Exit Sub
End If

If (sfflag = False) Then
searftr.Open seaf, studentcon, adOpenStatic, adLockOptimistic
bpbar.Value = 50
Set DataGrid1.DataSource = searftr
bpbar.Value = 70
DataGrid1.ReBind
sfflag = True
bpbar.Value = 85
Else
sfflag = False
searftr.Close
GoTo again
bpbar.Value = 90
End If

bpbar.Value = 100
bpbar.Value = 0

Exit Sub
errmess:
MsgBox Err.Description, vbCritical, "Error Occured"
bpbar.Value = 0
End Sub

Private Sub Form_Load()
' Events That Should Happen When Form Is Loaded
On Error Resume Next
Me.Top = 50
Me.Left = 50
Set searftr = New ADODB.Recordset
sfflag = False
Merlin "Search Student Family Information From Here", "Read"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Events That Should Happen When Form Is Unloaded
On Error Resume Next
If searftr.State = adStateOpen Then
searftr.Close
End If
End Sub

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub Image3_Click()
btnprint_Click
End Sub

Private Sub searchval_GotFocus()
Merlin "Enter One Search Value According To The Search Type"
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
