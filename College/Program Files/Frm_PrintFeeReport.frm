VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_PrintFeeReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Fee Report"
   ClientHeight    =   2145
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4320
   Icon            =   "Frm_PrintFeeReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3413
      Caption         =   "Select Details"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkCommand vkCommand1 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Show Report"
         Top             =   1440
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "Preview Report"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         BorderColor     =   11057596
         DrawFocus       =   0   'False
         DrawMouseInRect =   0   'False
         DisabledBackColor=   15070196
         CustomStyle     =   5
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
         Left            =   1440
         TabIndex        =   3
         ToolTipText     =   "Enter Field Value"
         Top             =   960
         Width           =   2415
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
         ItemData        =   "Frm_PrintFeeReport.frx":076A
         Left            =   1440
         List            =   "Frm_PrintFeeReport.frx":077A
         TabIndex        =   2
         ToolTipText     =   "Select Report Type"
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Field Value"
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
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Type"
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frm_PrintFeeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable Declarations
Dim rptrec As ADODB.Recordset
Dim strp As String

Private Sub cmbtype_GotFocus()
Merlin "Select One Report Type"
End Sub

Private Sub Form_Load()
' Events Happen When Form Is Loaded
On Error Resume Next
Me.Top = 50
Me.Left = 50
Merlin "Print Fee Information Report", "Read"
End Sub

Private Sub txtval_GotFocus()
Merlin "Enter Value To Generate Report"
End Sub

Private Sub txtval_KeyPress(KeyAscii As Integer)
' When Enter Key Is Pressed
On Error Resume Next
If KeyAscii = 13 Then
vkCommand1_Click
End If
End Sub

Private Sub vkCommand1_Click()
' Code For Printing Data
On Error Resume Next

Merlin "Show Report For Printing"
If cmbtype.Text = "" Then
MsgBox "Select One Report Type", vbInformation, "Report Type"
ElseIf cmbtype.Text = "All Records" And txtval.Text = "" Then
strp = "select * from FeesInformation order by Reciept_Number"
ElseIf cmbtype.Text = "By Name" And txtval.Text <> "" Then
strp = "select * from FeesInformation where Student_Name like '" & Trim$(txtval.Text) & "%'"
ElseIf cmbtype.Text = "By Admission Number" And txtval.Text <> "" Then
strp = "select * from FeesInformation where Admission_Number like '" & txtval.Text & "%'"
ElseIf cmbtype.Text = "By Reciept Number" And txtval.Text <> "" Then
strp = "select * from FeesInformation where Reciept_Number like '" & txtval.Text & "%'"
Else
MsgBox "Select Correct Report Configurations", vbInformation, "Error Occured"
Exit Sub
End If

Set rptrec = New ADODB.Recordset
rptrec.Open strp, feescon, adOpenStatic, adLockOptimistic
Set FeesInfoReport.DataSource = rptrec
Load FeesInfoReport
FeesInfoReport.Show
End Sub
