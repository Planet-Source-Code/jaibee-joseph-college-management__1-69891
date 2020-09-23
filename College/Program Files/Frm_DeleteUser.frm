VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#1.0#0"; "Candy.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_DeleteUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Existing User "
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4695
   Icon            =   "Frm_DeleteUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin Candy.CandyButton btnCancel 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      ToolTipText     =   "Exit From Form"
      Top             =   2640
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
   Begin Candy.CandyButton btnOK 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Delete User"
      Top             =   2640
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
      Caption         =   "Delete"
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
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3201
      Caption         =   "Enter Details Here"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         TabIndex        =   4
         ToolTipText     =   "Enter Password"
         Top             =   1080
         Width           =   2535
      End
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
         TabIndex        =   3
         ToolTipText     =   "Enter User Name Here"
         Top             =   600
         Width           =   2535
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
         TabIndex        =   6
         Top             =   1080
         Width           =   975
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
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   4200
      Picture         =   "Frm_DeleteUser.frx":076A
      ToolTipText     =   "Application Help"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Existing Software User Account."
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
      Top             =   360
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "Frm_DeleteUser.frx":0ED4
      Stretch         =   -1  'True
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "Frm_DeleteUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
' Exit Form
On Error Resume Next
Unload Me
Merlin "Exit From Form"
End Sub

Private Sub btnOK_Click()
' Delete User
On Error Resume Next
Merlin "Delete User Information"
If txtname.Text = "" Then
MsgBox "Enter User Name Then Press Delete", vbInformation, "No User Name"
txtname.SetFocus
Exit Sub
ElseIf txtpass.Text = "" Then
MsgBox "Enter Password Then Press Delete", vbInformation, "No Password"
txtname.SetFocus
Exit Sub
End If

If copyf.FileExists(App.Path & "\User\" & txtname.Text & ".SecurityFile") Then
keyuser = Split(SeCheck.ReadSecurityFileUser(App.Path & "\User\" & txtname.Text), "//")
If txtname.Text = crypt.DeCode(keyuser(0)) And txtpass.Text = crypt.DeCode(keyuser(1)) Then
copyf.DeleteFile App.Path & "\User\" & txtname.Text & ".SecurityFile", True
MsgBox "User Information Deleted", vbInformation, "User Deleted"
txtname.Text = ""
txtpass.Text = ""
Else
MsgBox "Incorrect Password Or UserName", vbInformation, "Incorrect Data"
Exit Sub
End If
Else
MsgBox "No Such User Exist", vbInformation, "No User"
Exit Sub
End If
End Sub

Private Sub Form_Load()
' When Form Is Loaded
On Error Resume Next
Me.Top = 50
Me.Left = 50
Set copyf = New FileSystemObject
Set SeCheck = New SecurityClass
Set crypt = New cSimpleCrypt
Merlin "You Can Delete Software User From Here", "Read"
End Sub

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub txtname_GotFocus()
Merlin "Enter Existing User Name Here"
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
' When Enter Pressed
On Error Resume Next
If KeyAscii = 13 Then btnOK_Click
End Sub

Private Sub txtpass_GotFocus()
Merlin "Enter User Password Here"
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
' When Enter Pressed
On Error Resume Next
If KeyAscii = 13 Then btnOK_Click
End Sub

