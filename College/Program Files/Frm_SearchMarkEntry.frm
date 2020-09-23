VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{5E8FF3F9-2372-4C96-A258-479E142BF3EF}#1.0#0"; "XP_ProBar.ocx"
Begin VB.Form Frm_SearchMarkEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Mark Entry"
   ClientHeight    =   6270
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9585
   Icon            =   "Frm_SearchMarkEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
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
      Top             =   720
      Width           =   2775
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
      ItemData        =   "Frm_SearchMarkEntry.frx":076A
      Left            =   6720
      List            =   "Frm_SearchMarkEntry.frx":077D
      TabIndex        =   2
      ToolTipText     =   "Select Search Type"
      Top             =   720
      Width           =   2775
   End
   Begin Candy.CandyButton btnprint 
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      ToolTipText     =   "Print Search Data"
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
      TabIndex        =   1
      ToolTipText     =   "Search For Data"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frm_SearchMarkEntry.frx":07D6
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Search Results"
      Top             =   1200
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
      BeginProperty Column03 
         DataField       =   "Exam_Date"
         Caption         =   "Exam Date"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
         DataField       =   "Mark_Obtained"
         Caption         =   "Mark Obtained"
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
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
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
      Picture         =   "Frm_SearchMarkEntry.frx":07EB
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_SearchMarkEntry.frx":0F55
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Student Mark Entry Records Here"
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
      TabIndex        =   7
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Value"
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
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Type"
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
      Left            =   5520
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   9120
      Picture         =   "Frm_SearchMarkEntry.frx":16BF
      ToolTipText     =   "Print Search Data"
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Frm_SearchMarkEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strstum As String
Dim searstumark As New ADODB.Recordset
Dim sfflag As Boolean

Private Sub btnprint_Click()
On Error GoTo message

If searstumark.State = adStateOpen Then
Set StudentMarkReport.DataSource = searstumark
Load StudentMarkReport
StudentMarkReport.Show
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
On Error Resume Next

Merlin "Search For Results"
again:
bpbar.Value = 0
If cmbtype.Text = "All Records" And txtval.Text = "" Then
strstum = "Select * from StudentMarkInformation order by Serial_Number"
bpbar.Value = 30
ElseIf cmbtype.Text = "By Name And Exam Type" And txtval.Text <> "" Then
examtyp = InputBox("Enter Exam Type", "Exam Type")
strstum = "Select * from StudentMarkInformation where Student_Name like '" & Trim$(txtval.Text) & "%' and Exam_Type like '" & examtyp & "%'"
bpbar.Value = 30
ElseIf cmbtype.Text = "By Date And Exam Type" And txtval.Text <> "" Then
examt = InputBox("Enter Exam Date", "Exam Date")
strstum = "Select * from StudentMarkInformation where Exam_Date = '" & Trim$(examt) & "' and Exam_Type like '" & Trim$(txtval.Text) & "%'"
bpbar.Value = 30
ElseIf cmbtype.Text = "By Date Only" And txtval.Text <> "" Then
strstum = "Select * from StudentMarkInformation where Exam_Date = '" & txtval.Text & "'"
bpbar.Value = 30
ElseIf cmbtype.Text = "By Subject" And txtval.Text <> "" Then
strstum = "Select * from StudentMarkInformation where Subject like '" & Trim$(txtval.Text) & "%'"
bpbar.Value = 30
Else
MsgBox "Select Correct Options Then Search", vbInformation, "Error Occured"
Exit Sub
End If

If (sfflag = False) Then
searstumark.Open strstum, GlobalCon, adOpenStatic, adLockOptimistic
bpbar.Value = 50
Set DataGrid1.DataSource = searstumark
bpbar.Value = 70
DataGrid1.ReBind
sfflag = True
bpbar.Value = 85
Else
sfflag = False
searstumark.Close
GoTo again
bpbar.Value = 90
End If

bpbar.Value = 100
bpbar.Value = 0

End Sub

Private Sub cmbtype_GotFocus()
Merlin "Select One Search Type"
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 50
Me.Left = 50
Merlin "Search Student Mark Entry Here", "Read"
sfflag = False
End Sub

Private Sub Image2_Click()
On Error Resume Next
btnprint_Click
End Sub

Private Sub Image3_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub txtval_GotFocus()
Merlin "Enter Search Value Here"
End Sub

Private Sub txtval_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then btnsearch_Click
End Sub
