VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Begin VB.Form Frm_QueryBuilder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Query Builder"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "Frm_QueryBuilder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   9960
   Begin Candy.CandyButton CandyButton4 
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      ToolTipText     =   "Clear Text Box"
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Clear"
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
   Begin Candy.CandyButton CandyButton3 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      ToolTipText     =   "Execute Query For Results"
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Execute"
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
   Begin Candy.CandyButton CandyButton2 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Open Existing Query"
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Open Query"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Candy.CandyButton CandyButton1 
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      ToolTipText     =   "Save Query"
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Save Query"
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
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Execution Results"
      Top             =   3600
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   16508095
      HeadLines       =   1
      RowHeight       =   18
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
      Caption         =   "Query Results"
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
            LCID            =   1033
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
      EndProperty
   End
   Begin VB.TextBox Text1 
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
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      ToolTipText     =   "Enter Query Here"
      Top             =   600
      Width           =   9735
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   9480
      Picture         =   "Frm_QueryBuilder.frx":076A
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "You Can Build Your Own Queries Here"
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
      TabIndex        =   2
      Top             =   240
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_QueryBuilder.frx":0ED4
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Frm_QueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CandyButton1_Click()
Merlin "Save Query To Hard Disk"
CommonDialog1.CancelError = True
On Error GoTo message

CommonDialog1.Filter = "Text Files (*.txt)|*.txt|SQL Query Files (*.sql)|*.sql"  ' Only Show Included Extensions
CommonDialog1.ShowSave
Open CommonDialog1.FileName For Output As #1
     Print #1, Frm_QueryBuilder.Text1.Text
Close #1
CandyButton1.Enabled = False

Exit Sub
message:
If Err.Number = 32755 Then
MsgBox "You Have Cancelled File Save", vbInformation, "Cancel Operation"
Else
MsgBox Err.Description, vbInformation, "Error Occured"
End If
End Sub

Private Sub CandyButton2_Click()
Merlin "Open One Existing Query"
CommonDialog1.CancelError = True
On Error GoTo message

Text1.Text = ""
CommonDialog1.Filter = "Text Files (*.txt)|*.txt|SQL Query File (*.sql)|*.sql"  ' Only Show Included Extensions
CommonDialog1.ShowOpen
Open CommonDialog1.FileName For Input As #1
    While Not EOF(1)
        Input #1, MyString
    Frm_QueryBuilder.Text1.Text = Frm_QueryBuilder.Text1.Text & MyString
    Wend
Close #1

Exit Sub
message:
If Err.Number = 32755 Then
MsgBox "You Have Cancelled File Open", vbInformation, "Cancel Operation"
Else
MsgBox Err.Description, vbInformation, "Error Occured"
End If
End Sub

Private Sub CandyButton3_Click()
On Error GoTo message

Merlin "Execute Query For Results"
If queryobjectims.State = adStateOpen Then queryobjectims.Close

queryobjectims.Open Text1.Text, GlobalCon, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = queryobjectims
DataGrid1.ReBind
DataGrid1.Refresh

Exit Sub
message:
MsgBox Err.Description, vbInformation, "Query Error"
End Sub

Private Sub CandyButton4_Click()
On Error Resume Next
Text1.Text = ""
Text1.SetFocus
Merlin "Clear Text Box"
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 50
Me.Left = 50
Merlin "You Can Create Your Own Queries Here", "Read"
End Sub

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub
