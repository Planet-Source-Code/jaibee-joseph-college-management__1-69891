VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_ChangeCharacter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Help Assistant"
   ClientHeight    =   1140
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2745
   Icon            =   "Frm_ChangeCharacter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkCommand vkCommand1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BackColor1      =   16777215
      BackColor2      =   13228765
      BackColorPushed1=   14215660
      BackColorPushed2=   16777215
      Caption         =   "Change"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Frm_ChangeCharacter.frx":076A
      Left            =   240
      List            =   "Frm_ChangeCharacter.frx":0774
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Frm_ChangeCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Me.Top = 50
Me.Left = 50
End Sub

Private Sub vkCommand1_Click()
If Combo1.Text = "Merlin" Then
charactername = "Merlin"
charfile = App.Path & "\Characters\merlin.acs"
SaveSetting App.CompanyName, "Character", "CharName", charactername
SaveSetting App.CompanyName, "Character", "CharFile", charfile
ElseIf Combo1.Text = "Question Mark" Then
charactername = "Qmark"
charfile = App.Path & "\Characters\qmark.acs"
SaveSetting App.CompanyName, "Character", "CharName", charactername
SaveSetting App.CompanyName, "Character", "CharFile", charfile
End If
MsgBox "Character Is Changed, Setting Will" & vbCrLf & "Take Effect On Next Start", vbInformation, "Character Changed"
Unload Me
End Sub
