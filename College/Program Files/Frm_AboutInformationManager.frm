VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{7152298A-50FE-11D3-A11A-0080C6F7AC86}#1.1#0"; "SWBLink.ocx"
Begin VB.Form Frm_AboutInformationManager 
   BorderStyle     =   0  'None
   Caption         =   "About Information Management System"
   ClientHeight    =   7305
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   8460
   ClipControls    =   0   'False
   Icon            =   "Frm_AboutInformationManager.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5042.042
   ScaleMode       =   0  'User
   ScaleWidth      =   7944.377
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      MaxLength       =   4
      TabIndex        =   12
      ToolTipText     =   "Enter Registration Key Here"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      MaxLength       =   9
      TabIndex        =   11
      ToolTipText     =   "Enter Registration Key Here"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      MaxLength       =   4
      TabIndex        =   10
      ToolTipText     =   "Enter Registration Key Here"
      Top             =   3000
      Width           =   495
   End
   Begin SWBLink.SWBHyperLink SWBHyperLink1 
      Height          =   255
      Left            =   3570
      TabIndex        =   9
      ToolTipText     =   "Author E-Mail Address"
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      CaptionTextColor=   -2147483630
      CaptionBackColor=   -2147483639
      CaptionHighlightColor=   -2147483646
      Caption         =   "Contact Author"
      Alignment       =   0
      Hyperlink       =   "mailto:jaibee.joseph@gmail.com"
      CaptionFontName =   "Trebuchet MS"
      CaptionFontSize =   8.25
      CaptionFontBold =   -1  'True
      CaptionFontItalic=   0   'False
      CaptionFontUnderline=   0   'False
   End
   Begin vkUserContolsXP.vkCommand cmdSysInfo 
      Height          =   375
      Left            =   6600
      TabIndex        =   5
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
         Name            =   "Arial"
         Size            =   8.25
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
   Begin vkUserContolsXP.vkToggleButton cmdOK 
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      ToolTipText     =   "Exit About"
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor1      =   16777215
      BackColor2      =   13228765
      BackColorPushed1=   14215660
      BackColorPushed2=   16777215
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Registration Key Here"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Version"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      ToolTipText     =   "Windows Version Information"
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Windows User Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Windows Default User"
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Computer Name"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm_AboutInformationManager.frx":000C
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1665
      Left            =   360
      TabIndex        =   3
      Top             =   5160
      Width           =   7935
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version : 1.0.0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1725
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Information Management System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3645
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm_AboutInformationManager.frx":0200
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   4485
   End
   Begin VB.Image Image1 
      Height          =   7455
      Left            =   0
      Picture         =   "Frm_AboutInformationManager.frx":0293
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "Frm_AboutInformationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSysInfo_Click()
' Code For Start To Show Windows Information
On Error Resume Next
wininfo.StartSysInfo
End Sub

Private Sub cmdOK_Click()
' Unload About
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
' Code To Be Executed When Form Is Loaded
On Error Resume Next
  
Me.Top = 700
Me.Left = 1500

' Check For Registration
If keyvalc.checkkeyims = True Then
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
End If

' Show Windows Information
Set info = New COSInfo
Set wininfo = New GetWindowsInformation
Label1.Caption = "Windows Version :" & " " & wininfo.getwinOS
Label3.Caption = "Computer Name :" & " " & info.ComputerName
Label2.Caption = "Windows User :" & " " & wininfo.UserName
Merlin "About Information Management System", "DoMagic1"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
' Save Registration Entry
On Error Resume Next
If KeyAscii = 13 Then
keyvalc.savekeyims Text1.Text, Text2.Text, Text3.Text
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End If
End Sub
