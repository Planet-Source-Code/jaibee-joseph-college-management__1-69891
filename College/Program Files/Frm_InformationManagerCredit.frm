VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_InformationManagerCredit 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Dialog Caption"
   ClientHeight    =   7485
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   8565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Frm_InformationManagerCredit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkToggleButton vkToggleButton2 
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      ToolTipText     =   "System Information"
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor1      =   16777215
      BackColor2      =   13228765
      BackColorPushed1=   14215660
      BackColorPushed2=   16777215
      Caption         =   "System Info"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
   Begin vkUserContolsXP.vkToggleButton vkToggleButton1 
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      ToolTipText     =   "Unload Form"
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor1      =   16777215
      BackColor2      =   13228765
      BackColorPushed1=   14215660
      BackColorPushed2=   16777215
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      BorderColor     =   11057596
      DrawFocus       =   0   'False
      DrawMouseInRect =   0   'False
      DisabledBackColor=   15070196
      CustomStyle     =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm_InformationManagerCredit.frx":076A
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   7935
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   7335
      Left            =   240
      Picture         =   "Frm_InformationManagerCredit.frx":1224
      Stretch         =   -1  'True
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "Frm_InformationManagerCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
' Code Executed When Form Is Loaded
On Error Resume Next
Set wininfo = New GetWindowsInformation
Me.Top = 700
Me.Left = 1500
Merlin "This Is Information Management System Credits", "DoMagic1"
End Sub

Private Sub vkToggleButton1_Click()
' When Cancel Button Is Pressed
On Error Resume Next
Unload Me
End Sub

Private Sub vkToggleButton2_Click()
' When System Information Button Is Pressed
On Error Resume Next
wininfo.StartSysInfo
End Sub

