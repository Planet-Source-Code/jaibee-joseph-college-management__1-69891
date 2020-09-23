VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_StaffAtnPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Staff Attendance"
   ClientHeight    =   2160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4320
   Icon            =   "Frm_StaffAtnPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
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
         ItemData        =   "Frm_StaffAtnPrint.frx":076A
         Left            =   1200
         List            =   "Frm_StaffAtnPrint.frx":077A
         TabIndex        =   3
         ToolTipText     =   "Select Report Type From Here"
         Top             =   480
         Width           =   2655
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
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Enter Report Type Value"
         Top             =   960
         Width           =   2655
      End
      Begin vkUserContolsXP.vkCommand vkCommand1 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Click Me To Show Report"
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
         TabIndex        =   5
         Top             =   480
         Width           =   975
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
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frm_StaffAtnPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable Declarations Needed For Code
Dim stre As String
Dim datlow As String
Dim dathigh As String
Dim printstaffatn As ADODB.Recordset

Private Sub cmbtype_GotFocus()
Merlin "Select One Report Type From Here"
End Sub

Private Sub Form_Load()
' Events That Should Happen When Form Is Loaded
On Error Resume Next

Me.Top = 50
Me.Left = 50
Merlin "You Can Use Different Parameters And Generate New Reports From Here"
End Sub

Private Sub txtval_GotFocus()
Merlin "Enter Matching Value For The Report Type Here"
End Sub

Private Sub txtval_KeyPress(KeyAscii As Integer)
' When Enter Key Is Pressed
On Error Resume Next

If KeyAscii = 13 Then
vkCommand1_Click
End If
End Sub

Private Sub vkCommand1_Click()
' Code For Searching Record
On Error Resume Next

Merlin "Print Staff Attendance Report"
If cmbtype.Text = "Between Two Dates" And txtval.Text = "" Then
datlow = InputBox("Enter First Date", "Report By Date")
dathigh = InputBox("Enter Second Date", "Report By Date")
stre = "select * from StaffAttendanceInformation where Atn_Date >= '" & datlow & "' and Atn_Date <= '" & dathigh & "'"
ElseIf cmbtype.Text = "All Records" And txtval.Text = "" Then
stre = "select * from StaffAttendanceInformation order by Serial_Number"
ElseIf cmbtype.Text = "By Name" And txtval.Text <> "" Then
stre = "select * from StaffAttendanceInformation where Staff_Name like '" & Trim$(txtval.Text) & "%'"
ElseIf cmbtype.Text = "Between Two Dates and Name" And txtval.Text <> "" Then
datlow = InputBox("Enter First Date", "Report By Date")
dathigh = InputBox("Enter Second Date", "Report By Date")
stre = "select * from StaffAttendanceInformation where Atn_Date >= '" & datlow & "' and Atn_Date <= '" & dathigh & "' and Staff_Name like '" & Trim$(txtval.Text) & "%'"
Else
MsgBox "Select Report Configurations", vbInformation, "Error Occured"
Exit Sub
End If

Set printstaffatn = New ADODB.Recordset
printstaffatn.Open stre, staffcon, adOpenStatic, adLockOptimistic
Set StaffAttnReport.DataSource = printstaffatn
Load StaffAttnReport
StaffAttnReport.Show
End Sub
