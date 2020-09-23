VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{5E8FF3F9-2372-4C96-A258-479E142BF3EF}#1.0#0"; "XP_ProBar.ocx"
Begin VB.Form Frm_Sidebar 
   BorderStyle     =   0  'None
   Caption         =   "IMS Side Bar"
   ClientHeight    =   8115
   ClientLeft      =   11565
   ClientTop       =   2580
   ClientWidth     =   3495
   ControlBox      =   0   'False
   Icon            =   "Frm_Sidebar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   2712
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4471
      Caption         =   "Software Operations"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkCommand vkCommand4 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Credits Of IMS 1.0.0"
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "Credits IMS"
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
      Begin vkUserContolsXP.vkCommand vkCommand3 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Operating System Information"
         Top             =   1920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "OS Information"
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
      Begin vkUserContolsXP.vkCommand vkCommand2 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "About Information Management System"
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "About IMS 1.0.0"
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
      Begin vkUserContolsXP.vkCommand vkCommand1 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Exit From Software"
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BackColor1      =   16777215
         BackColor2      =   13228765
         BackColorPushed1=   14215660
         BackColorPushed2=   16777215
         Caption         =   "Exit From Software"
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
   End
   Begin vkUserContolsXP.vkTimer vkTimer2 
      Left            =   2880
      Top             =   4920
      _ExtentX        =   926
      _ExtentY        =   926
      Interval        =   5000
      Enabled         =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FBE4BF&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      Picture         =   "Frm_Sidebar.frx":08CA
      ScaleHeight     =   2265
      ScaleWidth      =   3225
      TabIndex        =   8
      Top             =   240
      Width           =   3255
      Begin VB.Label lblTipText 
         BackColor       =   &H00FBE4BF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   2835
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FBE4BF&
         Caption         =   "Did you know..."
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
         Left            =   540
         TabIndex        =   9
         Top             =   120
         Width           =   2655
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Todays Date And Time "
      Top             =   6720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2143
      Caption         =   "TODAY"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_Sidebar.frx":1034
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3360
         Top             =   1320
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "04:12 AM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "11/23/03"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
   End
   Begin vkUserContolsXP.vkTimer vkTimer1 
      Left            =   2880
      Top             =   5640
      _ExtentX        =   926
      _ExtentY        =   926
      Interval        =   50
      Enabled         =   -1  'True
   End
   Begin XP_ProBar.UserControl1 membar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6288
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16576
      Scrolling       =   9
      ShowText        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5424
      Width           =   3255
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5856
      Width           =   3255
   End
End
Attribute VB_Name = "Frm_Sidebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
On Error Resume Next

Me.BackColor = MainMenu.ACPRibbon1.BackColor
Label1.Caption = "Expiry Date" & ":" & " " & GetSetting("VRJ Soft", "Exp Date", "RegExp", dateexp)
Me.Top = 90
Me.Left = 11565
Label4.Caption = Date

Call checkthemeside

' Seed Rnd
Randomize
Call loadtipsall(lblTipText)

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label6.Caption = Time
End Sub

Private Sub vkCommand1_Click()
On Error Resume Next
Unload MainMenu
End Sub

Private Sub vkCommand2_Click()
On Error Resume Next
Call loadabout
End Sub

Private Sub vkCommand3_Click()
On Error Resume Next
Call loadwininfo
End Sub

Private Sub vkCommand4_Click()
On Error Resume Next
Call loadcredit
End Sub

Private Sub vkTimer1_Timer()
' Show Memory Status
On Error Resume Next
Label16.Caption = "Available Memory :" & " " & (CDbl(memsta.AvailablePhysical)) & " " & "KB"
membar.Value = 100 - memsta.MemoryLoad
End Sub

Private Sub vkTimer2_Timer()
On Error Resume Next
Call DoNextTipSide
End Sub
