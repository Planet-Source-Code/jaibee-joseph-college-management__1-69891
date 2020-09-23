VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_PrintStudentReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Student Report"
   ClientHeight    =   2175
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4335
   Icon            =   "Frm_PrintStudentReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4335
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
         TabIndex        =   3
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
         Left            =   1320
         TabIndex        =   2
         ToolTipText     =   "Enter Report Type Value Here"
         Top             =   960
         Width           =   2535
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
         ItemData        =   "Frm_PrintStudentReport.frx":076A
         Left            =   1320
         List            =   "Frm_PrintStudentReport.frx":077A
         TabIndex        =   1
         ToolTipText     =   "Select Report Type"
         Top             =   480
         Width           =   2535
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
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frm_PrintStudentReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strrep As String
Dim printstudetail As ADODB.Recordset

Private Sub cmbtype_GotFocus()
Merlin "Select One Report Type From Here"
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 50
Me.Left = 50
Merlin "Print Student Reports Here Without Picture", "Read"
End Sub

Private Sub txtval_GotFocus()
Merlin "Enter Appropriate Value Here For Generating Report"
End Sub

Private Sub txtval_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
vkCommand1_Click
End If
End Sub

Private Sub vkCommand1_Click()
On Error Resume Next

Merlin "Show Report For Printing"
If cmbtype.Text = "All Records" And txtval.Text = "" Then
strrep = "Select * from StudentInformation order by Admission_Number"
ElseIf cmbtype.Text = "By Student Name" And txtval.Text <> "" Then
strrep = "Select * from StudentInformation where Student_Name like '" & Trim$(txtval.Text) & "%'"
ElseIf cmbtype.Text = "By Admission Number" And txtval.Text <> "" Then
strrep = "Select * from StudentInformation where Admission_Number like '" & CDbl(txtval.Text) & "%'"
ElseIf cmbtype.Text = "By Class And Year" And txtval.Text <> "" Then
yearclass = InputBox("Enter Class Year", "Enter Year")
strrep = "Select * from StudentInformation where Course_Name like '" & Trim$(txtval.Text) & "%' and Year_Course like '" & yearclass & "%'"
Else
MsgBox "Select Correct Configuration", vbInformation, "Error Occured"
Exit Sub
End If

Set printstudetail = New ADODB.Recordset
printstudetail.Open strrep, GlobalCon, adOpenStatic, adLockOptimistic
Set StudentDetailReport.DataSource = printstudetail
Load StudentDetailReport
StudentDetailReport.Show
End Sub
