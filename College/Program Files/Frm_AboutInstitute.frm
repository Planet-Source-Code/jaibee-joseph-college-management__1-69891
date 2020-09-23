VERSION 5.00
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{7152298A-50FE-11D3-A11A-0080C6F7AC86}#1.1#0"; "SWBLink.ocx"
Begin VB.Form Frm_AboutInstitute 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Institution"
   ClientHeight    =   3870
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7920
   Icon            =   "Frm_AboutInstitute.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   8040
      TabIndex        =   17
      Text            =   "Text7"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog diag1 
      Left            =   8040
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ButtonXp.XPButton XPButton3 
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Close"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin ButtonXp.XPButton XPButton2 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OK"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin SWBLink.SWBHyperLink link2 
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   3240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      CaptionTextColor=   0
      CaptionBackColor=   -2147483633
      CaptionHighlightColor=   16576
      Caption         =   ""
      Alignment       =   0
      Hyperlink       =   ""
      CaptionFontName =   "Arial Narrow"
      CaptionFontSize =   8.25
      CaptionFontBold =   -1  'True
      CaptionFontItalic=   0   'False
      CaptionFontUnderline=   0   'False
   End
   Begin SWBLink.SWBHyperLink link1 
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   3240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      CaptionTextColor=   0
      CaptionBackColor=   -2147483633
      CaptionHighlightColor=   16576
      Caption         =   ""
      Alignment       =   0
      Hyperlink       =   ""
      CaptionFontName =   "Arial Narrow"
      CaptionFontSize =   8.25
      CaptionFontBold =   -1  'True
      CaptionFontItalic=   0   'False
      CaptionFontUnderline=   0   'False
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Left            =   2160
      TabIndex        =   6
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Left            =   2160
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin ButtonXp.XPButton XPButton1 
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Browse"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   5520
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Affiliated To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name Of The Institution"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Frm_AboutInstitute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim namein As String
Dim affiliatedin As String
Dim addressin As String
Dim phonein As String
Dim webin As String
Dim emailin As String
Dim logopath As String

Private Sub Form_Load()
On Error Resume Next

Merlin "This Is About The College Where Sofftware Is Used", "Read"
Me.Top = 50
Me.Left = 50

link1.CaptionBackColor = MainMenu.ACPRibbon1.BackColor
link2.CaptionBackColor = MainMenu.ACPRibbon1.BackColor

namein = GetSetting(App.CompanyName, "InsSetting", "InstituteName", namein)
affiliatedin = GetSetting(App.CompanyName, "InsSetting", "InstituteAffi", affiliatedin)
addressin = GetSetting(App.CompanyName, "InsSetting", "InstituteAddress", addressin)
phonein = GetSetting(App.CompanyName, "InsSetting", "PhoneNumber", phonein)
webin = GetSetting(App.CompanyName, "InsSetting", "InstituteWeb", webin)
emailin = GetSetting(App.CompanyName, "InsSetting", "InstituteEMail", emailin)
logopath = GetSetting(App.CompanyName, "InsSetting", "LogoPath", logopath)

If namein <> "" Then
Text1.Text = namein
Text2.Text = affiliatedin
Text3.Text = addressin
Text4.Text = phonein
Text5.Text = webin
Text6.Text = emailin
link1.Caption = webin
link1.Hyperlink = webin
link2.Caption = emailin
link2.Hyperlink = "mailto:" & emailin
Image1.Picture = LoadPicture(logopath)
XPButton2.Caption = "Change"
enadistext (False)
Me.Caption = "About" & "  " & namein
End If

End Sub

Private Sub Text6_LostFocus()
On Error Resume Next

link1.Caption = Text5.Text
link1.Hyperlink = Text5.Text
link2.Caption = Text6.Text
link2.Hyperlink = "mailto:" & Text6.Text
End Sub

Private Sub XPButton1_Click()
On Error Resume Next

Call browseimage(diag1, Image1, Text7)
logopath = diag1.FileName
End Sub

Private Sub XPButton2_Click()
On Error Resume Next

namein = Text1.Text
affiliatedin = Text2.Text
addressin = Text3.Text
phonein = Text4.Text
webin = Text5.Text
emailin = Text6.Text

If namein <> "" And XPButton2.Caption = "OK" Then
SaveSetting App.CompanyName, "InsSetting", "InstituteName", namein
SaveSetting App.CompanyName, "InsSetting", "InstituteAffi", affiliatedin
SaveSetting App.CompanyName, "InsSetting", "InstituteAddress", addressin
SaveSetting App.CompanyName, "InsSetting", "PhoneNumber", phonein
SaveSetting App.CompanyName, "InsSetting", "InstituteWeb", webin
SaveSetting App.CompanyName, "InsSetting", "InstituteEMail", emailin
SaveSetting App.CompanyName, "InsSetting", "LogoPath", logopath
End If

If XPButton2.Caption = "Change" Then
enadistext (True)
XPButton2.Caption = "OK"
ElseIf XPButton2.Caption = "OK" Then
enadistext (False)
XPButton2.Caption = "Change"
End If

Me.Caption = "About" & "  " & namein
End Sub

Private Function enadistext(enaval As Boolean)
Text1.Enabled = enaval
Text2.Enabled = enaval
Text3.Enabled = enaval
Text4.Enabled = enaval
Text5.Enabled = enaval
Text6.Enabled = enaval
XPButton1.Enabled = enaval
End Function

Private Sub XPButton3_Click()
Unload Me
End Sub
