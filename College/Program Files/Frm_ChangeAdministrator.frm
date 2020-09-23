VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#1.0#0"; "Candy.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_ChangeAdministrator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Administrator Password"
   ClientHeight    =   3060
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4680
   Icon            =   "Frm_ChangeAdministrator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin Candy.CandyButton btnCancel 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      ToolTipText     =   "Unload Admin Change Form"
      Top             =   2520
      Width           =   2175
      _ExtentX        =   3836
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
      TabIndex        =   1
      ToolTipText     =   "Click Me To Check Existing Password And To Create New Password"
      Top             =   2520
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Create"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3201
      Caption         =   "Enter New Details Here"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtname 
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
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Administrator User Name"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtpass 
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
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Enter Password Here"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   975
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
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   4200
      Picture         =   "Frm_ChangeAdministrator.frx":076A
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "Frm_ChangeAdministrator.frx":0ED4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Existing Administrator Account."
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
      TabIndex        =   7
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Frm_ChangeAdministrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
' Code For Changing Administrator Password
On Error Resume Next
key = Split(SeCheck.ReadSecurityFile, "//")
If changeas = True Then
GoTo changeadmin

Else

If Frm_ChangeAdministrator.txtname.Text = crypt.DeCode(key(0)) And Frm_ChangeAdministrator.txtpass.Text = crypt.DeCode(key(1)) Then
MsgBox "Verification Compltete", vbInformation, "Verify"
txtpass.Text = ""
txtpass.SetFocus
changeas = True
Merlin "Password Verification Completed"

changeadmin:
If txtpass.Text = "" Then
   MsgBox "A New Password Is Required.", vbInformation, "Password Required"
   txtpass.SetFocus
   Exit Sub
End If
If SeCheck.SaveSecurityFile(crypt.Encode(txtname.Text), crypt.Encode(txtpass.Text)) = True Then
   MsgBox "New Username And Password Is Saved.", vbInformation, "Saved"
   Merlin "New Admin Created"
   Unload Me
Else
   MsgBox "Sorry Unable to Add New Security Information." & _
             vbNewLine & _
             "This Error Is Caused By:" & vbNewLine & _
             "-----------------------------------------------------------------" & _
             vbNewLine & _
             "(x) Unable To Create Security Information Due To Windows File And Folder Authorization.", vbCritical, _
             "Security Information Creation Error."
End If

Else
   MsgBox "Sorry, Incompatible Username and Password.", vbCritical, "Error Occured"
   changeas = False
   txtpass.SetFocus
End If
End If
End Sub

Private Sub Form_Load()
' Code Executed When Form Is Loaded
On Error Resume Next
txtname.Text = "Administrator"
changeas = False
Me.Top = 50
Me.Left = 50
Merlin "Change Administrator From Here", "Read"
End Sub

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
' When Key Is Pressed In Text Box Admin Name
On Error Resume Next
If KeyAscii = 13 Then
btnOK_Click
End If
End Sub

Private Sub txtpass_GotFocus()
Merlin "Enter Password Here"
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
' When Key Is Pressed In Text Box Admin Password
On Error Resume Next
If KeyAscii = 13 Then
btnOK_Click
End If
End Sub
