VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#1.0#0"; "Candy.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_LoginMain 
   BorderStyle     =   0  'None
   Caption         =   "IMS 1.0.0  Login Form"
   ClientHeight    =   3885
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   4455
   Icon            =   "Frm_LoginMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin Candy.CandyButton btnCancel 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      ToolTipText     =   "Unload Form"
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Cancel"
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
   Begin Candy.CandyButton btnOK 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Login Using Password And UserName"
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
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
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
      Caption         =   "Enter User Name and Password Here."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtusername 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         ToolTipText     =   "Enter User Name"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtpass 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Enter Password"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox txtname 
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
         ItemData        =   "Frm_LoginMain.frx":076A
         Left            =   1320
         List            =   "Frm_LoginMain.frx":0774
         TabIndex        =   1
         ToolTipText     =   "Select Login Type"
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   3960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Login Type"
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
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter User Name and Password to Enter Software"
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
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "Frm_LoginMain.frx":078D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Frm_LoginMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
' When Cancel Button Is Pressed
On Error Resume Next
Unload Me
Unload MainMenu
End Sub

Private Sub btnOK_Click()
' When Button OK Is Pressed
On Error Resume Next
If txtname.Text = "" Then
   MsgBox "A User Type Is Required.", vbCritical, "User Type Error"
   txtname.SetFocus
   Exit Sub
End If

If txtusername.Locked = False And txtusername.Text = "" Then
   MsgBox "A User Name Is Required", vbInformation, "No User Name"
   txtusername.SetFocus
   Exit Sub
End If

If txtpass.Text = "" Then
   MsgBox "A Password Is Required.", vbCritical, "Password Error"
   txtpass.SetFocus
   Exit Sub
End If

If txtname.Text = "Administrator" Then
Call readadminpass
admcheck = crypt.DeCode(key(0))
Else
Call readuserpass
If userex = True Then
admcheck = crypt.DeCode(keyuser(0))
End If
End If

End Sub

Private Sub Form_Load()
' Code Executed When Form Is Loaded
On Error Resume Next
Set SeCheck = New SecurityClass
Set crypt = New cSimpleCrypt
End Sub

Private Sub txtname_Change()
' Check which controls are to be enabled
On Error Resume Next
Call checkstat
End Sub

Private Sub txtname_Click()
' Check which controls are to be enabled
On Error Resume Next
Call checkstat
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
' When Enter Key Is Pressed
On Error Resume Next
If KeyAscii = 13 Then btnOK_Click
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
' When Enter Key Is Pressed
On Error Resume Next
If KeyAscii = 13 Then btnOK_Click
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
' When Enter Key Is Pressed
On Error Resume Next
If KeyAscii = 13 Then btnOK_Click
End Sub

Public Function checkstat()
' Function For Checking Control Status
If txtname.Text = "Administrator" Then
   txtusername.Locked = True
   txtusername.Text = ""
   txtpass.Text = ""
ElseIf txtname.Text = "User" Then
   txtusername.Locked = False
   txtusername.Text = ""
   txtpass.Text = ""
End If
End Function
