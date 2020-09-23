VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#1.0#0"; "Candy.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_TipInformation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tip Of The Day"
   ClientHeight    =   3285
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   4470
   Icon            =   "Frm_TipInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin vkUserContolsXP.vkCheck chkLoadTipsAtStartup 
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Click To Unload Form When Software Starts Next Time"
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Load At Start Up"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Candy.CandyButton cmdNextTip 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Click To Show Next Tool Tip"
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Next Tip"
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
   Begin Candy.CandyButton cmdOK 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Click To Unload Tool Tip Form"
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "OK"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FBE4BF&
      Height          =   2535
      Left            =   120
      Picture         =   "Frm_TipInformation.frx":076A
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   3
      Top             =   120
      Width           =   4215
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
         Height          =   1275
         Left            =   180
         TabIndex        =   5
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label1 
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
         TabIndex        =   4
         Top             =   180
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Frm_TipInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkLoadTipsAtStartup_Click()
    ' save whether or not this form should be displayed at startup
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value
End Sub

Private Sub cmdNextTip_Click()
    Call DoNextTip
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
        
Me.Top = 50
Me.Left = 50

' See if we should be shown at startup
ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
If ShowAtStartup = 0 Then
    Unload Me
    Exit Sub
End If
        
' Set the checkbox, this will force the value to be written back out to the registry
Me.chkLoadTipsAtStartup.Value = vbChecked
    
' Seed Rnd
Randomize
    
Call loadtipsall(lblTipText)
End Sub

