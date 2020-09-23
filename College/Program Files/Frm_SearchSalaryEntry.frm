VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{5E8FF3F9-2372-4C96-A258-479E142BF3EF}#1.0#0"; "XP_ProBar.ocx"
Begin VB.Form Frm_SearchSalaryEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Staff Salary Entry"
   ClientHeight    =   6525
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9600
   Icon            =   "Frm_SearchSalaryEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   6
      ToolTipText     =   "Enter Search Value"
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
      ItemData        =   "Frm_SearchSalaryEntry.frx":076A
      Left            =   7200
      List            =   "Frm_SearchSalaryEntry.frx":077D
      TabIndex        =   5
      ToolTipText     =   "Select Search Type"
      Top             =   600
      Width           =   2295
   End
   Begin Candy.CandyButton btnprint 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Print Search Data"
      Top             =   5640
      Width           =   2775
      _ExtentX        =   4895
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
   Begin Candy.CandyButton btnhelp 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      ToolTipText     =   "Application Help"
      Top             =   5640
      Width           =   2775
      _ExtentX        =   4895
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
   Begin Candy.CandyButton btnsearch 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Search For Data"
      Top             =   5640
      Width           =   2655
      _ExtentX        =   4683
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
      Bindings        =   "Frm_SearchSalaryEntry.frx":07C6
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Search Results"
      Top             =   1080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7858
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Reciept_Number"
         Caption         =   "Reciept Number"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Staff_Salary"
         Caption         =   "Staff Salary"
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
         DataField       =   "Date_Pay"
         Caption         =   "Date Pay"
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
         DataField       =   "Pay_Amount"
         Caption         =   "Pay Amount"
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
         DataField       =   "Pay_Due"
         Caption         =   "Pay Due"
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
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin XP_ProBar.UserControl1 bpbar 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6120
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
      Picture         =   "Frm_SearchSalaryEntry.frx":07DB
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   9120
      Picture         =   "Frm_SearchSalaryEntry.frx":0F45
      ToolTipText     =   "Print Search Data"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label3 
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
      Left            =   5640
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Staff Salary Entry"
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
      TabIndex        =   4
      Top             =   240
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_SearchSalaryEntry.frx":16AF
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Frm_SearchSalaryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable Declarations
Dim searchstr As String
Dim searstaffsal As New ADODB.Recordset
Dim sfflag As Boolean

Private Sub btnhelp_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub btnprint_Click()
' When Print Button Is Clicked
On Error GoTo message

If searstaffsal.State = adStateOpen Then
Set StaffSalaryReport.DataSource = searstaffsal
Load StaffSalaryReport
StaffSalaryReport.Show
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
' Code For Searching Records
On Error GoTo message

Merlin "Search For Results"
again:
bpbar.Value = 0
If valuetyp.Text = "All Records" And searchval.Text = "" Then
searchstr = "Select * from StaffSalaryInformation Order By Reciept_Number"
bpbar.Value = 30
ElseIf valuetyp.Text = "By Staff Name" And searchval.Text <> "" Then
searchstr = "Select * from StaffSalaryInformation where Staff_Name like '" & Trim$(searchval.Text) & "%'"
bpbar.Value = 30
ElseIf valuetyp.Text = "By Staff ID" And searchval.Text <> "" Then
searchstr = "Select * from StaffSalaryInformation where Staff_ID like '" & CDbl(searchval.Text) & "%'"
bpbar.Value = 30
ElseIf valuetyp.Text = "By Reciept Number" And searchval.Text <> "" Then
searchval = "Select * from StaffSalaryInformation where Reciept_Number like '" & CDbl(searchval.Text) & "%'"
bpbar.Value = 30
ElseIf valuetyp.Text = "By Date" And searchval.Text <> "" Then
searchstr = "Select * from StaffSalaryInformation where Date_Pay = '" & searchval.Text & "'"
bpbar.Value = 30
Else
MsgBox "Select Correct Configuration Options", vbInformation, "Configure Search"
Exit Sub
End If

If (sfflag = False) Then
searstaffsal.Open searchstr, staffcon, adOpenStatic, adLockOptimistic
bpbar.Value = 50
Set DataGrid1.DataSource = searstaffsal
bpbar.Value = 70
DataGrid1.ReBind
sfflag = True
bpbar.Value = 85
Else
sfflag = False
searstaffsal.Close
GoTo again
bpbar.Value = 90
End If

bpbar.Value = 100
bpbar.Value = 0

Exit Sub
message:
MsgBox Err.Description, vbCritical, "Error Occured"
End Sub

Private Sub Form_Load()
' When Form Is Loaded
On Error Resume Next
Me.Top = 50
Me.Left = 50
Merlin "Search Staff Salary Entry Here", "Read"
sfflag = False
End Sub

Private Sub Image2_Click()
' When Print Image Is Clicked
On Error Resume Next
btnprint_Click
End Sub

Private Sub Image3_Click()
btnhelp_Click
End Sub

Private Sub searchval_GotFocus()
Merlin "Enter One Search Value"
End Sub

Private Sub searchval_KeyPress(KeyAscii As Integer)
' When Enter Key Is Pressed
On Error Resume Next
If KeyAscii = 13 Then
btnsearch_Click
End If
End Sub

Private Sub valuetyp_GotFocus()
Merlin "Select One Search Type"
End Sub
