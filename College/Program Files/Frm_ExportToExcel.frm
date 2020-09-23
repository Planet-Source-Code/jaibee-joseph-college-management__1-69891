VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_ExportToExcel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Tables To Excel"
   ClientHeight    =   4680
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3840
   Icon            =   "Frm_ExportToExcel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3413
      Caption         =   "Export Staff Table"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Candy.CandyButton expbydepartment 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Export By Department"
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
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
         Caption         =   "Export By Department"
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
      Begin VB.ComboBox cmbdepartment 
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
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Select Department"
         Top             =   960
         Width           =   3135
      End
      Begin Candy.CandyButton expallstaff 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Export Whole Staff Table To Excel"
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
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
         Caption         =   "Export Whole Table"
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
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3413
      Caption         =   "Export Student Table"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbyear 
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
         ItemData        =   "Frm_ExportToExcel.frx":076A
         Left            =   1920
         List            =   "Frm_ExportToExcel.frx":078F
         TabIndex        =   4
         ToolTipText     =   "Select Year"
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cmbcourse 
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
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Select Class"
         Top             =   960
         Width           =   1455
      End
      Begin Candy.CandyButton expbyclass 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Export By Class"
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
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
         Caption         =   "Export Data By Class"
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
      Begin Candy.CandyButton btnexpall 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Export Student Table (Whole) To Excel"
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
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
         Caption         =   "Export Whole Table"
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
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   3360
      Picture         =   "Frm_ExportToExcel.frx":080B
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Export Tables To Excel"
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
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_ExportToExcel.frx":0F75
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Frm_ExportToExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnexpall_Click()
' Export All Data In Student Table
On Error Resume Next
Merlin "Export Whole Student Table To Excel"
Dim expstr As String
expstr = "select * from StudentInformation order by Admission_Number"
studentrec.Open expstr, studentcon, adOpenStatic, adLockOptimistic
Call exporttoexcel(CommonDialog1, studentrec)
End Sub

Private Sub cmbcourse_GotFocus()
Merlin "Select One Course Name From Here"
End Sub

Private Sub cmbdepartment_GotFocus()
Merlin "Select One Department From Here"
End Sub

Private Sub cmbyear_GotFocus()
Merlin "Select Course Year From Here"
End Sub

Private Sub expallstaff_Click()
' Export Whole Staff Table
On Error Resume Next
Merlin "Export Whole Staff Table To Excel"
Dim expstr As String
expstr = "select * from StaffInformation order by Staff_ID"
staffrec.Open expstr, staffcon, adOpenStatic, adLockOptimistic
Call exporttoexcel(CommonDialog1, staffrec)
End Sub

Private Sub expbyclass_Click()
' Export By Class (Student Table)
On Error Resume Next
Merlin "Export Data According To Class Information"
If cmbcourse.Text = "" Then
MsgBox "Select One Class", vbInformation, "Select Class"
Exit Sub
ElseIf cmbyear.Text = "" Then
MsgBox "Select Course Year", vbInformation, "Select Year"
Exit Sub
Else
Dim expstr As String
expstr = "select * from StudentInformation where Course_Name like '" & Trim$(cmbcourse.Text) & "%' and Year_Course like '" & Trim$(cmbyear.Text) & "%'"
studentrec.Open expstr, studentcon, adOpenStatic, adLockOptimistic
Call exporttoexcel(CommonDialog1, studentrec)
End If
End Sub

Private Sub expbydepartment_Click()
' Export By Department
On Error Resume Next
Merlin "Export Data By Department"
If cmbdepartment.Text = "" Then
MsgBox "Select One Department", vbInformation, "Select Department"
Exit Sub
Else
Dim expstr As String
expstr = "select * from StaffInformation where Department like '" & Trim$(cmbdepartment.Text) & "%'"
staffrec.Open expstr, staffcon, adOpenStatic, adLockOptimistic
Call exporttoexcel(CommonDialog1, staffrec)
End If
End Sub

Private Sub Form_Load()
' When Form Is Loaded
On Error Resume Next
Merlin "You Can Export Database Tables From Here", "Read"
Me.Top = 50
Me.Left = 50

exportcourse.Movefirst
Do While Not exportcourse.BOF And Not exportcourse.EOF
   cmbcourse.AddItem exportcourse(1).Value
   exportcourse.Movenext
Loop

exportdepart.Movefirst
Do While Not exportdepart.BOF And Not exportdepart.EOF
   cmbdepartment.AddItem exportdepart(1).Value
   exportdepart.Movenext
Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
' When Form Unloads
On Error Resume Next
Call MainConClose
Call MainConEstablish
End Sub

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub
