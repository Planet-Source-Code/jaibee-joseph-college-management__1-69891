VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Begin VB.Form Frm_PrintStudent 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Candy.CandyButton btnClose 
      Height          =   375
      Left            =   5040
      TabIndex        =   33
      ToolTipText     =   "Close Form"
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Close"
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
   Begin Candy.CandyButton btnprint 
      Height          =   375
      Left            =   5040
      TabIndex        =   32
      ToolTipText     =   "Print Data"
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Information Management System Student Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1508
      TabIndex        =   34
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   31
      Top             =   8160
      Width           =   2775
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   30
      Top             =   7680
      Width           =   2775
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   29
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   28
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   27
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   26
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   25
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   24
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   23
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   22
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   21
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   20
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   19
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   18
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   17
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   16
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Permanent Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Roll Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Group"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      Caption         =   "Caste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "Emergency Contact"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   6720
      Width           =   1455
   End
End
Attribute VB_Name = "Frm_PrintStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
On Error Resume Next
Unload Me
Merlin "Unload Report"
End Sub

Private Sub btnprint_Click()
On Error Resume Next
Merlin "Print Report To View"
btnprint.Visible = False
btnClose.Visible = False
'this code will invoke print dialog box
'for printing bill, any number of copies
'can be printed
Dim intLoopIndex As Integer
    
On Error GoTo Cancel
CommonDialog1.CancelError = True
CommonDialog1.PrinterDefault = True
CommonDialog1.ShowPrinter
For intLoopIndex = 1 To CommonDialog1.Copies
    PrintForm
Next intLoopIndex
Unload Me

Exit Sub
Cancel:
MsgBox "Cancelled Printing", vbInformation, "Print Error"
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 50
Me.Left = 50
Merlin "Print Student Information With Picture From Here", "DoMagic1"
End Sub
