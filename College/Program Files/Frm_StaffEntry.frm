VERSION 5.00
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "prjXTab.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_StaffEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Information Entry"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "Frm_StaffEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   9030
   Begin prjXTab.XTab stfdetailtab 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10186
      TabCount        =   2
      TabCaption(0)   =   "Staff Professional Detail"
      TabContCtrlCnt(0)=   40
      Tab(0)ContCtrlCap(1)=   "cdlgImage"
      Tab(0)ContCtrlCap(2)=   "pictext"
      Tab(0)ContCtrlCap(3)=   "cmbdesignation"
      Tab(0)ContCtrlCap(4)=   "cmbdepartment"
      Tab(0)ContCtrlCap(5)=   "txtsubjectshandled"
      Tab(0)ContCtrlCap(6)=   "cmbsex"
      Tab(0)ContCtrlCap(7)=   "txtqualification"
      Tab(0)ContCtrlCap(8)=   "txtsalarymon"
      Tab(0)ContCtrlCap(9)=   "btnhelp"
      Tab(0)ContCtrlCap(10)=   "btnmovelast"
      Tab(0)ContCtrlCap(11)=   "btnmovenext"
      Tab(0)ContCtrlCap(12)=   "btnmovefirst"
      Tab(0)ContCtrlCap(13)=   "btnmoveprevious"
      Tab(0)ContCtrlCap(14)=   "btnbrowse"
      Tab(0)ContCtrlCap(15)=   "btncancel"
      Tab(0)ContCtrlCap(16)=   "btndelete"
      Tab(0)ContCtrlCap(17)=   "btnedit"
      Tab(0)ContCtrlCap(18)=   "btnsave"
      Tab(0)ContCtrlCap(19)=   "btnaddnew"
      Tab(0)ContCtrlCap(20)=   "txtemailaddress"
      Tab(0)ContCtrlCap(21)=   "txtmobnumber"
      Tab(0)ContCtrlCap(22)=   "txtphonenumber"
      Tab(0)ContCtrlCap(23)=   "txtperaddress"
      Tab(0)ContCtrlCap(24)=   "txttempaddress"
      Tab(0)ContCtrlCap(25)=   "txtstaffname"
      Tab(0)ContCtrlCap(26)=   "txtstaffid"
      Tab(0)ContCtrlCap(27)=   "Label13"
      Tab(0)ContCtrlCap(28)=   "Label12"
      Tab(0)ContCtrlCap(29)=   "Label11"
      Tab(0)ContCtrlCap(30)=   "imgHolder"
      Tab(0)ContCtrlCap(31)=   "Label10"
      Tab(0)ContCtrlCap(32)=   "Label9"
      Tab(0)ContCtrlCap(33)=   "Label8"
      Tab(0)ContCtrlCap(34)=   "Label7"
      Tab(0)ContCtrlCap(35)=   "Label6"
      Tab(0)ContCtrlCap(36)=   "Label5"
      Tab(0)ContCtrlCap(37)=   "Label4"
      Tab(0)ContCtrlCap(38)=   "Label3"
      Tab(0)ContCtrlCap(39)=   "Label2"
      Tab(0)ContCtrlCap(40)=   "Label1"
      TabCaption(1)   =   "Staff Salary Entry"
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "vkFrame1"
      TabTheme        =   3
      ActiveTabBackStartColor=   16316664
      InActiveTabBackStartColor=   15066597
      ActiveTabForeColor=   10972496
      InActiveTabForeColor=   9474192
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   9474192
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      Begin vkUserContolsXP.vkFrame vkFrame1 
         Height          =   4935
         Left            =   -74760
         TabIndex        =   57
         Top             =   600
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8705
         Caption         =   "Enter Staff Salary Details Here"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin ButtonXp.XPButton XPButton1 
            Height          =   330
            Left            =   4800
            TabIndex        =   27
            ToolTipText     =   "Browse Staff Information"
            Top             =   1440
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "..."
            ForeColor       =   -2147483642
            ForeHover       =   0
         End
         Begin VB.TextBox txtrecieptnumber 
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
            Height          =   330
            Left            =   2640
            TabIndex        =   25
            ToolTipText     =   "Enter Pay Receipt Number"
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtstaffnames 
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
            Height          =   330
            Left            =   2640
            TabIndex        =   28
            ToolTipText     =   "Enter Staff Name Here"
            Top             =   1920
            Width           =   2415
         End
         Begin VB.TextBox txtstaffsalarys 
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
            Height          =   330
            Left            =   2640
            TabIndex        =   29
            ToolTipText     =   "Enter Staff Salary"
            Top             =   2400
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker dateselect 
            Height          =   375
            Left            =   2640
            TabIndex        =   30
            ToolTipText     =   "Select Pay Date"
            Top             =   2880
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16508095
            Format          =   49086465
            CurrentDate     =   39434
         End
         Begin VB.TextBox txtpayamount 
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
            Height          =   330
            Left            =   2640
            TabIndex        =   31
            ToolTipText     =   "Enter Pay Amount"
            Top             =   3360
            Width           =   2415
         End
         Begin VB.TextBox txtpaydue 
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
            Height          =   330
            Left            =   2640
            TabIndex        =   32
            Top             =   3840
            Width           =   2415
         End
         Begin Candy.CandyButton btnadd 
            Height          =   375
            Left            =   5160
            TabIndex        =   33
            ToolTipText     =   "Add New Entry"
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Add New"
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
         Begin Candy.CandyButton btnsaves 
            Height          =   375
            Left            =   5160
            TabIndex        =   34
            ToolTipText     =   "Save Record"
            Top             =   1440
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Save"
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
         Begin Candy.CandyButton btnedits 
            Height          =   375
            Left            =   5160
            TabIndex        =   35
            ToolTipText     =   "Edit Record"
            Top             =   1920
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Edit"
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
         Begin Candy.CandyButton btndeletes 
            Height          =   375
            Left            =   5160
            TabIndex        =   36
            ToolTipText     =   "Delete Record"
            Top             =   2400
            Width           =   1935
            _ExtentX        =   3413
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
         Begin Candy.CandyButton btncanceles 
            Height          =   375
            Left            =   5160
            TabIndex        =   37
            ToolTipText     =   "Cancel Entry"
            Top             =   2880
            Width           =   1935
            _ExtentX        =   3413
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
         Begin Candy.CandyButton btnmovefirsts 
            Height          =   330
            Left            =   5160
            TabIndex        =   38
            ToolTipText     =   "Move To First"
            Top             =   3840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "<<"
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
         Begin Candy.CandyButton btnmovepreviouss 
            Height          =   330
            Left            =   5520
            TabIndex        =   39
            ToolTipText     =   "Move To Previous"
            Top             =   3840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "<"
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
         Begin Candy.CandyButton btnmovenexts 
            Height          =   330
            Left            =   6480
            TabIndex        =   40
            ToolTipText     =   "Move To Next"
            Top             =   3840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ">"
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
         Begin Candy.CandyButton btnmovelasts 
            Height          =   330
            Left            =   6840
            TabIndex        =   41
            ToolTipText     =   "Move To Last"
            Top             =   3840
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ">>"
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
         Begin VB.TextBox txtstaffids 
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
            Height          =   330
            Left            =   2640
            TabIndex        =   26
            ToolTipText     =   "Enter Staff ID Here"
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Staff ID"
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
            Left            =   1080
            TabIndex        =   64
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Staff Name"
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
            Left            =   1080
            TabIndex        =   63
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Staff Salary"
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
            Left            =   1080
            TabIndex        =   62
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Amount"
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
            Left            =   1080
            TabIndex        =   61
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Due"
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
            Left            =   1080
            TabIndex        =   60
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   1080
            TabIndex        =   59
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Reciept Number"
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
            Left            =   1080
            TabIndex        =   58
            Top             =   960
            Width           =   1815
         End
      End
      Begin MSComDlg.CommonDialog cdlgImage 
         Left            =   240
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox pictext 
         Height          =   285
         Left            =   1200
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmbdesignation 
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
         Height          =   345
         Left            =   1920
         TabIndex        =   5
         ToolTipText     =   "Select Staff Designation From Here"
         Top             =   2640
         Width           =   2415
      End
      Begin VB.ComboBox cmbdepartment 
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
         Height          =   345
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "Select Staff Department From Here"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox txtsubjectshandled 
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
         Height          =   330
         Left            =   1920
         TabIndex        =   7
         ToolTipText     =   "Subjects Handled By Staff"
         Top             =   3600
         Width           =   2415
      End
      Begin VB.ComboBox cmbsex 
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
         Height          =   345
         ItemData        =   "Frm_StaffEntry.frx":076A
         Left            =   1920
         List            =   "Frm_StaffEntry.frx":0774
         TabIndex        =   3
         ToolTipText     =   "Select Staff Sex From Here"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtqualification 
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
         Height          =   330
         Left            =   1920
         TabIndex        =   4
         ToolTipText     =   "Enter Staff Qualification Here"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtsalarymon 
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
         Height          =   330
         Left            =   1920
         TabIndex        =   8
         ToolTipText     =   "Staff Salary/Month"
         Top             =   4080
         Width           =   2415
      End
      Begin Candy.CandyButton btnhelp 
         Height          =   375
         Left            =   4560
         TabIndex        =   24
         ToolTipText     =   "Application Help"
         Top             =   3480
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Help"
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
      Begin Candy.CandyButton btnmovelast 
         Height          =   255
         Left            =   8040
         TabIndex        =   23
         ToolTipText     =   "Move To Last Record In The Database"
         Top             =   3120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ">>"
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
      Begin Candy.CandyButton btnmovenext 
         Height          =   255
         Left            =   7560
         TabIndex        =   22
         ToolTipText     =   "Move To Next Record In The Database"
         Top             =   3120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ">"
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
      Begin Candy.CandyButton btnmovefirst 
         Height          =   255
         Left            =   6480
         TabIndex        =   20
         Tag             =   "Arrow Buttons Are Used To Move Records To Different Positions."
         ToolTipText     =   "Move To First Record In The Database"
         Top             =   3120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "<<"
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
      Begin Candy.CandyButton btnmoveprevious 
         Height          =   255
         Left            =   6960
         TabIndex        =   21
         ToolTipText     =   "Move To Previous Record In The Database"
         Top             =   3120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "<"
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
      Begin Candy.CandyButton btnbrowse 
         Height          =   375
         Left            =   4560
         TabIndex        =   18
         ToolTipText     =   "Browse For The Picture"
         Top             =   3000
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
         Caption         =   "Browse"
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
      Begin Candy.CandyButton btncancel 
         Height          =   375
         Left            =   6480
         TabIndex        =   19
         ToolTipText     =   "Cancel Edit Or Add New"
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
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
      Begin Candy.CandyButton btndelete 
         Height          =   375
         Left            =   6480
         TabIndex        =   17
         ToolTipText     =   "Delete Current Record"
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
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
      Begin Candy.CandyButton btnedit 
         Height          =   375
         Left            =   6480
         TabIndex        =   16
         ToolTipText     =   "Edit Current Record"
         Top             =   1680
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Edit"
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
      Begin Candy.CandyButton btnsave 
         Height          =   375
         Left            =   6480
         TabIndex        =   15
         ToolTipText     =   "Save Edited Or New Record"
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Save"
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
      Begin Candy.CandyButton btnaddnew 
         Height          =   375
         Left            =   6480
         TabIndex        =   14
         ToolTipText     =   "Add New Record To Database"
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Add New"
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
      Begin VB.TextBox txtemailaddress 
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
         Height          =   330
         Left            =   6000
         TabIndex        =   13
         ToolTipText     =   "Enter E-Maill Address If Any"
         Top             =   5040
         Width           =   2415
      End
      Begin VB.TextBox txtmobnumber 
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
         Height          =   330
         Left            =   6000
         TabIndex        =   12
         ToolTipText     =   "Enter Mobile Number Here "
         Top             =   4560
         Width           =   2415
      End
      Begin VB.TextBox txtphonenumber 
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
         Height          =   330
         Left            =   6000
         TabIndex        =   11
         ToolTipText     =   "Enter Staff House Phone Number Here"
         Top             =   4080
         Width           =   2415
      End
      Begin VB.TextBox txtperaddress 
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
         Height          =   360
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         ToolTipText     =   "Enter Staff Permanent Addres Here"
         Top             =   5040
         Width           =   2415
      End
      Begin VB.TextBox txttempaddress 
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
         Height          =   360
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         ToolTipText     =   "Enter Staff Temporary Address Here"
         Top             =   4560
         Width           =   2415
      End
      Begin VB.TextBox txtstaffname 
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
         Height          =   330
         Left            =   1920
         TabIndex        =   2
         ToolTipText     =   "Enter Staff Name Here"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtstaffid 
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
         Height          =   330
         Left            =   1920
         TabIndex        =   1
         ToolTipText     =   "Enter Staff ID Here"
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Subjects Handled"
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
         TabIndex        =   54
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label12 
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
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         TabIndex        =   52
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Image imgHolder 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   4560
         Stretch         =   -1  'True
         ToolTipText     =   "Staff Picture"
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Salary/Month"
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
         TabIndex        =   51
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label9 
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
         Left            =   4560
         TabIndex        =   50
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   4560
         TabIndex        =   49
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   4560
         TabIndex        =   48
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Temporary Address"
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
         TabIndex        =   46
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
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
         TabIndex        =   45
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Qualification"
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
         TabIndex        =   44
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff Name"
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
         TabIndex        =   43
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID"
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
         TabIndex        =   42
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   8520
      Picture         =   "Frm_StaffEntry.frx":0786
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Staff Details And Salary Information Here."
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
      TabIndex        =   55
      Top             =   240
      Width           =   8295
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_StaffEntry.frx":0EF0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Frm_StaffEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable Declarations Needed
Dim str As String
Dim strs As String
Dim saveflag As Boolean
Dim savef As Boolean
Dim X As String

Private Function locktextbox(stat As Boolean)
' Function To Lock And Unlock Text Boxses In The Form
txtstaffid.Locked = stat
txtstaffname.Locked = stat
cmbsex.Locked = stat
txtqualification.Locked = stat
cmbdesignation.Locked = stat
cmbdepartment.Locked = stat
txtsubjectshandled.Locked = stat
txtsalarymon.Locked = stat
txttempaddress.Locked = stat
txtperaddress.Locked = stat
txtphonenumber.Locked = stat
txtmobnumber.Locked = stat
txtemailaddress.Locked = stat
End Function

Private Sub btnadd_Click()
' Code For Adding New Record To Database
On Error GoTo message

Merlin "Click Me To Add New Entry"
savef = True
Call locktxtstaffsal(False)
Call disablebtnstaffs(False)
Call lockbtnstaffs(True)
Call cleardatastaffs
dateselect.Value = Now
txtrecieptnumber.SetFocus

Exit Sub
message:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnaddnew_Click()
' Code For Adding New Record To Database
On Error GoTo message

saveflag = True
Call locktextbox(False)
Call disablebtnstaff(False)
Call lockbtnstaff(True)
Call cleardatastaff
imgHolder.Picture = LoadPicture(pictext.Text)
txtstaffid.SetFocus
Merlin "Click Me To Add New Staff Entry"

Exit Sub
message:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnbrowse_Click()
' Code For Browsing Image
Merlin "Browse Staff Image"
Call browseimage(cdlgImage, imgHolder, pictext)
End Sub

Private Sub btnCancel_Click()
' Code For Cancelling Editing Or Add New
On Error Resume Next

Call cleardatastaff
Call locktextbox(True)
Call disablebtnstaff(False)
Call lockbtnstaff(False)
Call checkbtnstaff
Merlin "Cancel Add New Or Edit"

If staffrec.BOF And staffrec.EOF Then
MsgBox "No Existing Record, Insert New Record", vbInformation, "No Record"
Else
staffrec.Movefirst
Call showdatastaff
End If
End Sub

Private Sub btncanceles_Click()
' Code For Cancelling Editing Or Add New
On Error Resume Next

Call cleardatastaffs
Call locktxtstaffsal(True)
Call disablebtnstaffs(False)
Call lockbtnstaffs(False)
Call checkbtnstaffs
Merlin "Cancel Add New Or Edit"

If staffsalary.BOF And staffsalary.EOF Then
MsgBox "No Existing Record, Insert New Record", vbInformation, "No Record"
Else
staffsalary.Movefirst
Call showdatastaffs
End If
End Sub

Private Sub btndelete_Click()
' Code To Delete Current Record
On Error GoTo message

Merlin "Delete Current Record"
If MsgBox("Execution Of Command Will Delete Current Datarecord" & vbCrLf & "Are You Sure You Wan't To Delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then

str = "DELETE FROM StaffInformation WHERE "
str = str & "Staff_ID = "
str = str & CDbl(txtstaffid.Text)
staffcon.Execute str
staffrec.Requery

If pictext.Text <> "" Then
copyf.DeleteFile pictext.Text, True
End If

MsgBox "Record Deleted Sucessfully.", vbInformation, "Delete Record"

If staffrec.BOF And staffrec.EOF Then
Call cleardatastaff
MsgBox ("The Previous Record Was Last Record."), vbInformation, "Last Record"
Call checkbtnstaff
imgHolder.Picture = LoadPicture("")
Else
staffrec.Movenext
If staffrec.EOF Then
staffrec.Movelast
End If
Call showdatastaff
End If

End If

Exit Sub
message:
MsgBox "No Existing Record, Insert New Record", vbInformation, "Error Occured"
End Sub

Private Sub btndeletes_Click()
' Code To Delete Current Record
On Error GoTo message

Merlin "Delete Current Record"
If MsgBox("Execution Of Command Will Delete Current Datarecord" & vbCrLf & "Are You Sure You Wan't To Delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then

strs = "DELETE FROM StaffSalaryInformation WHERE "
strs = strs & "Reciept_Number = "
strs = strs & CDbl(txtrecieptnumber.Text)
staffcon.Execute strs
staffsalary.Requery

MsgBox "Record Deleted Sucessfully.", vbInformation, "Delete Record"

If staffsalary.BOF And staffsalary.EOF Then
Call cleardatastaffs
MsgBox ("The Previous Record Was Last Record."), vbInformation, "Last Record"
Call checkbtnstaffs
Else
staffsalary.Movenext
If staffsalary.EOF Then
staffsalary.Movelast
End If
Call showdatastaffs
End If

End If

Exit Sub
message:
MsgBox "No Existing Record, Insert New Record", vbInformation, "Error Occured"
End Sub

Private Sub btnedit_Click()
' Code For Editing Record
On Error GoTo mesa

saveflag = False
Call locktextbox(False)
Call disablebtnstaff(False)
Call lockbtnstaff(True)
Merlin "Edit Current Record"

Exit Sub
mesa:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnedits_Click()
' Code For Editing Record
On Error GoTo mesa

savef = False
Call locktxtstaffsal(False)
Call disablebtnstaffs(False)
Call lockbtnstaffs(True)
Merlin "Edit Current Record"

Exit Sub
mesa:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnhelp_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub btnmovefirst_Click()
' Move To First Record In The Table
On Error GoTo GoFirstError

staffrec.Movefirst
'Show the current data record
Call showdatastaff
 
Exit Sub

GoFirstError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub btnmovefirsts_Click()
' Move To First Record In The Table
On Error GoTo GoFirstError

staffsalary.Movefirst
'Show the current data record
Call showdatastaffs
 
Exit Sub

GoFirstError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub btnmovelast_Click()
' Move To Last Record In The Table
On Error GoTo GoLastError

staffrec.Movelast
' Show the current data record
Call showdatastaff
Exit Sub

GoLastError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub btnmovelasts_Click()
' Move To Last Record In The Table
On Error GoTo GoLastError

staffsalary.Movelast
' Show the current data record
Call showdatastaffs
Exit Sub

GoLastError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub btnmovenext_Click()
' Move To Next Record In The Table
On Error GoTo GoNextError
  
If Not staffrec.EOF Then staffrec.Movenext
If staffrec.EOF And staffrec.RecordCount > 0 Then
' Moved off the end so go back
staffrec.Movelast
End If
' Show the current data record
Call showdatastaff
  
Exit Sub
GoNextError:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnmovenexts_Click()
' Move To Next Record In The Table
On Error GoTo GoNextError
  
If Not staffsalary.EOF Then staffsalary.Movenext
If staffsalary.EOF And staffsalary.RecordCount > 0 Then
' Moved off the end so go back
staffsalary.Movelast
End If
' Show the current data record
Call showdatastaffs
  
Exit Sub
GoNextError:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnmoveprevious_Click()
' Move To Previous Record In The Table
On Error GoTo GoPrevError
  
If Not staffrec.BOF Then staffrec.Moveprevious
If staffrec.BOF And staffrec.RecordCount > 0 Then
    
' Moved off the end so go back
staffrec.Moveprevious
 
End If
' Show the current data record
Call showdatastaff
Exit Sub

GoPrevError:
If Err.Number = 3021 Then
MsgBox ("This Is First Record."), vbInformation, "First Record"
staffrec.Movenext
ElseIf Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, "Error Occured"
End If
End Sub

Private Sub btnmovepreviouss_Click()
' Move To Previous Record In The Table
On Error GoTo GoPrevError
  
If Not staffsalary.BOF Then staffsalary.Moveprevious
If staffsalary.BOF And staffsalary.RecordCount > 0 Then
    
' Moved off the end so go back
staffsalary.Moveprevious
 
End If
' Show the current data record
Call showdatastaffs
Exit Sub

GoPrevError:
If Err.Number = 3021 Then
MsgBox ("This Is First Record."), vbInformation, "First Record"
staffsalary.Movenext
ElseIf Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, "Error Occured"
End If
End Sub

Private Sub btnsave_Click()
' Code For Saving Edited Or New Record
On Error GoTo message

Merlin "Click Me To Save Record"

If checkall = True Then
If pictext.Text = "" And saveflag = True Then
X = MsgBox("Do You Want To Enter Picture Of The Staff", vbInformation + vbYesNo, "Picture Entry")
If X = vbYes Then
btnbrowse_Click
Exit Sub
Else
GoTo savequery
End If
End If

If pictext.Text = strfn Then
pictext.Text = App.Path & "\StaffImages\" & txtstaffid.Text & ".JPG"
copyf.CopyFile strfn, pictext.Text, True
ElseIf strfn = "" And pictext = "" Then
GoTo savequery
End If

savequery:

If saveflag = True Then
str = "INSERT INTO StaffInformation"
str = str & "(Staff_ID, Staff_Name, Sex, Qualification, Designation, Department, Subjects_Handled, Salary_Month, Temporary_Address, Permanent_Address, Phone_Number, Mobile_Number, EMail_Address, Pic_Staff) "
str = str & "VALUES(" & CDbl(txtstaffid.Text) & ", "
str = str & "'" & Trim$(txtstaffname.Text) & "', "
str = str & "'" & Trim$(cmbsex.Text) & "', "
str = str & "'" & Trim$(txtqualification.Text) & "', "
str = str & "'" & Trim$(cmbdesignation.Text) & "', "
str = str & "'" & Trim$(cmbdepartment.Text) & "', "
str = str & "'" & Trim$(txtsubjectshandled.Text) & "', "
str = str & "'" & Trim$(txtsalarymon.Text) & "', "
str = str & "'" & Trim$(txttempaddress.Text) & "', "
str = str & "'" & Trim$(txtperaddress.Text) & "', "
str = str & "'" & Trim$(txtphonenumber.Text) & "', "
str = str & "'" & Trim$(txtmobnumber.Text) & "', "
str = str & "'" & Trim$(txtemailaddress.Text) & "', "
str = str & "'" & Trim$(pictext.Text) & "')"
staffcon.Execute str
Else
str = "UPDATE StaffInformation SET "
str = str & "Staff_ID=" & CDbl(txtstaffid.Text) & ","
str = str & "Staff_Name='" & Trim$(txtstaffname.Text) & "',"
str = str & "Sex='" & Trim$(cmbsex.Text) & "',"
str = str & "Qualification='" & Trim$(txtqualification.Text) & "',"
str = str & "Designation='" & Trim$(cmbdesignation.Text) & "',"
str = str & "Department='" & Trim$(cmbdepartment.Text) & "',"
str = str & "Subjects_Handled='" & Trim$(txtsubjectshandled.Text) & "',"
str = str & "Salary_Month='" & Trim$(txtsalarymon.Text) & "',"
str = str & "Temporary_Address='" & Trim$(txttempaddress.Text) & "',"
str = str & "Permanent_Address='" & Trim$(txtperaddress.Text) & "',"
str = str & "Phone_Number='" & Trim$(txtphonenumber.Text) & "',"
str = str & "Mobile_Number='" & Trim$(txtmobnumber.Text) & "',"
str = str & "EMail_Address='" & Trim$(txtemailaddress.Text) & "',"
str = str & "Pic_Staff='" & Trim$(pictext.Text) & "'"
str = str & " WHERE Staff_ID=" & CDbl(txtstaffid.Text)
staffcon.Execute str
End If

staffrec.Requery
staffrec.Movelast
Call showdatastaff
Call lockbtnstaff(False)
Call checkbtnstaff
Call locktextbox(True)

MsgBox "Record Has Been Successfully Saved", vbInformation, "Saved"
End If

Exit Sub
message:
If Err.Number = -2147217900 Then
MsgBox ("Staff ID Already Exist,Please Enter Another Number"), vbCritical, "Staff ID Exist"
Else
MsgBox Err.Description, vbCritical, "Error Occured"
End If
End Sub

Private Sub btnsaves_Click()
' Code For Saving Edited Or New Record
On Error GoTo message

Merlin "Save Record"
If checkalls = True Then
If savef = True Then
strs = "INSERT INTO StaffSalaryInformation"
strs = strs & "(Reciept_Number, Staff_ID, Staff_Name, Staff_Salary, Date_Pay, Pay_Amount, Pay_Due) "
strs = strs & "VALUES(" & CDbl(txtrecieptnumber.Text) & ", "
strs = strs & CDbl(txtstaffids.Text) & ", "
strs = strs & "'" & Trim$(txtstaffnames.Text) & "', "
strs = strs & CDbl(txtstaffsalarys.Text) & ", "
strs = strs & "'" & dateselect.Value & "', "
strs = strs & CDbl(txtpayamount.Text) & ", "
strs = strs & CDbl(txtpaydue.Text) & ")"
staffcon.Execute strs
Else
strs = "UPDATE StaffSalaryInformation SET "
strs = strs & "Reciept_Number=" & CDbl(txtrecieptnumber.Text) & ","
strs = strs & "Staff_ID=" & CDbl(txtstaffids.Text) & ","
strs = strs & "Staff_Name='" & Trim$(txtstaffnames.Text) & "',"
strs = strs & "Staff_Salary=" & CDbl(txtstaffsalarys.Text) & ","
strs = strs & "Date_Pay='" & dateselect.Value & "',"
strs = strs & "Pay_Amount=" & CDbl(txtpayamount.Text) & ","
strs = strs & "Pay_Due=" & CDbl(txtpaydue.Text) & ""
strs = strs & " WHERE Reciept_Number=" & CDbl(txtrecieptnumber.Text)
staffcon.Execute strs
End If

staffsalary.Requery
staffsalary.Movelast
Call lockbtnstaffs(False)
Call checkbtnstaffs
Call locktxtstaffsal(True)
Call showdatastaffs

MsgBox "Record Has Been Successfully Saved", vbInformation, "Saved"
End If

Exit Sub
message:
If Err.Number = -2147217900 Then
MsgBox ("Staff ID Already Exist,Please Enter Another Number"), vbCritical, "Staff ID Exist"
Else
MsgBox Err.Description, vbCritical, "Error Occured"
End If
End Sub

Private Sub cmbdepartment_GotFocus()
Merlin "Select Staff Department From Here"
End Sub

Private Sub cmbdesignation_GotFocus()
Merlin "Select Staff Designation From Here"
End Sub

Private Sub cmbsex_GotFocus()
Merlin "Select Staff Sex From Here"
End Sub

Private Sub dateselect_GotFocus()
Merlin "Select Payment Date From Here"
End Sub

Private Sub Form_Load()
' Events That Should Happen When Form Is Loaded
On Error GoTo message

GlobalDesignation.Movefirst
Do While Not GlobalDesignation.BOF And Not GlobalDesignation.EOF
   cmbdesignation.AddItem GlobalDesignation(1).Value
   GlobalDesignation.Movenext
Loop

GlobalDepartment.Movefirst
Do While Not GlobalDepartment.BOF And Not GlobalDepartment.EOF
   cmbdepartment.AddItem GlobalDepartment(1).Value
   GlobalDepartment.Movenext
Loop

Call locktxtstaffsal(True)
Call lockbtnstaffs(False)
Call checkbtnstaffs
Call showdatastaffs

Call locktextbox(True)
Call lockbtnstaff(False)
Call checkbtnstaff
Call showdatastaff

Me.Top = 50
Me.Left = 50
Merlin "This Is Where Staff Information Is Entered", "DoMagic1"

Exit Sub
message:
MsgBox Err.Description, vbCritical, "Error Occured"
End Sub

Private Function checkall() As Boolean
' Function To Check Whether All The Entries Are Correct Or Not
Dim stat As Boolean

stat = False

If txtstaffid.Text = "" Or IsNumeric(txtstaffid.Text) = False Then
MsgBox "Enter Staff ID Correctly Then Save", vbInformation, "No Staff ID"
txtstaffid.SetFocus
ElseIf txtstaffname.Text = "" Then
MsgBox "Enter Staff Name", vbInformation, "No Staff Name"
txtstaffname.SetFocus
ElseIf cmbsex.Text = "" Then
MsgBox "Enter Staff Sex", vbInformation, "No Staff Sex"
cmbsex.SetFocus
ElseIf txtqualification.Text = "" Then
MsgBox "Enter Staff Qualification", vbInformation, "Empty Field"
txtqualification.SetFocus
ElseIf cmbdesignation.Text = "" Then
MsgBox "Select Designation", vbInformation, "Empty Designation"
cmbdesignation.SetFocus
ElseIf cmbdepartment.Text = "" Then
MsgBox "Select Department", vbInformation, "Empty Department"
cmbdepartment.SetFocus
ElseIf txtperaddress.Text = "" Then
MsgBox "Enter Staff Permanent Address", vbInformation, "Empty Field"
txtperaddress.SetFocus
ElseIf ValidateEmail(txtemailaddress.Text) = False Then
MsgBox "Wrong E-Mail Entry", vbInformation, "Wrong E-Mail"
txtemailaddress.SetFocus
Else
stat = True
End If

checkall = stat
End Function

Private Function showdatastaff()
' Show Data In The Table
If staffrec.EOF = False And staffrec.BOF = False Then
txtstaffid.Text = staffrec.Fields(0)
txtstaffname.Text = staffrec.Fields(1)
cmbsex.Text = staffrec.Fields(2)
txtqualification.Text = staffrec.Fields(3)
cmbdesignation.Text = staffrec.Fields(4)
cmbdepartment.Text = staffrec.Fields(5)
txtsubjectshandled.Text = staffrec.Fields(6)
txtsalarymon.Text = staffrec.Fields(7)
txttempaddress.Text = staffrec.Fields(8)
txtperaddress.Text = staffrec.Fields(9)
txtphonenumber.Text = staffrec.Fields(10)
txtmobnumber.Text = staffrec.Fields(11)
txtemailaddress.Text = staffrec.Fields(12)
pictext.Text = staffrec.Fields(13)
If copyf.FileExists(pictext.Text) Then
imgHolder.Picture = LoadPicture(pictext.Text)
ElseIf pictext.Text = "" Then
imgHolder.Picture = LoadPicture("")
End If
End If
End Function

Private Function cleardatastaff()
' Clear Data Fields
txtstaffid.Text = ""
txtstaffname.Text = ""
cmbsex.Text = ""
txtqualification.Text = ""
cmbdesignation.Text = ""
cmbdepartment.Text = ""
txtsubjectshandled.Text = ""
txtsalarymon.Text = ""
txttempaddress.Text = ""
txtperaddress.Text = ""
txtphonenumber.Text = ""
txtmobnumber.Text = ""
txtemailaddress.Text = ""
pictext.Text = ""
End Function

Private Function checkbtnstaff()
' Check Whether The Buttons Should Be Enabled Or Not
If staffrec.RecordCount = 0 Then
Call disablebtnstaff(False)
Else
Call disablebtnstaff(True)
End If
End Function

Private Function disablebtnstaff(statb As Boolean)
' Disable Buttons In The Form
btnedit.Enabled = statb
btndelete.Enabled = statb
btnmovefirst.Enabled = statb
btnmoveprevious.Enabled = statb
btnmovenext.Enabled = statb
btnmovelast.Enabled = statb
End Function

Private Function lockbtnstaff(lockst As Boolean)
' Lock And Unlock Buttons
btnsave.Enabled = lockst
btncancel.Enabled = lockst
btnbrowse.Enabled = lockst
End Function

Private Function locktxtstaffsal(statt As Boolean)
' Function To Lock And Unlock Text Box In Staff Salary Entry
txtrecieptnumber.Locked = statt
txtstaffids.Locked = statt
txtstaffnames.Locked = statt
txtstaffsalarys.Locked = statt
txtpayamount.Locked = statt
txtpaydue.Locked = statt
End Function

Private Function lockbtnstaffs(locksts As Boolean)
' Function To Lock And Unlock Buttons In Staff Salary Entry
btnsaves.Enabled = locksts
dateselect.Enabled = locksts
btncanceles.Enabled = locksts
XPButton1.Enabled = locksts
End Function

Private Function disablebtnstaffs(statsb As Boolean)
' Function To Enable And Disable Buttons In Staff Salary Entry
btnedits.Enabled = statsb
btndeletes.Enabled = statsb
btnmovelasts.Enabled = statsb
btnmovenexts.Enabled = statsb
btnmovepreviouss.Enabled = statsb
btnmovefirsts.Enabled = statsb
End Function

Private Function cleardatastaffs()
' Function To Clear Data
txtrecieptnumber.Text = ""
txtstaffids.Text = ""
txtstaffnames.Text = ""
txtstaffsalarys.Text = ""
txtpayamount.Text = ""
txtpaydue.Text = ""
End Function

Private Function checkbtnstaffs()
' Check Whether Buttons Should Be Enabled Or Disabled
If staffsalary.RecordCount = 0 Then
Call disablebtnstaffs(False)
Else
Call disablebtnstaffs(True)
End If
End Function
Private Function showdatastaffs()
' Show Data In The Table
If staffsalary.EOF = False And staffsalary.BOF = False Then
txtrecieptnumber.Text = staffsalary.Fields(0)
txtstaffids.Text = staffsalary.Fields(1)
txtstaffnames.Text = staffsalary.Fields(2)
txtstaffsalarys.Text = staffsalary.Fields(3)
dateselect.Value = staffsalary.Fields(4)
txtpayamount.Text = staffsalary.Fields(5)
txtpaydue.Text = staffsalary.Fields(6)
End If
End Function

Private Function checkalls() As Boolean
' Function To Check Data
Dim stats As Boolean
stats = False

If txtrecieptnumber.Text = "" Or Not IsNumeric(txtrecieptnumber.Text) Then
MsgBox "Enter Reciept Number Correctly", vbInformation, "Incorrect Entry"
ElseIf txtstaffids.Text = "" Or Not IsNumeric(txtstaffids.Text) Then
MsgBox "Enter Staff ID Correctly", vbInformation, "Incorrect Entry"
ElseIf txtstaffnames.Text = "" Then
MsgBox "Enter Staff Name", vbInformation, "Empty Field"
ElseIf txtstaffsalarys.Text = "" Or Not IsNumeric(txtstaffsalarys.Text) Then
MsgBox "Enter Staff Salary Correctly", vbInformation, "Incorrect Entry"
ElseIf txtpayamount.Text = "" Or Not IsNumeric(txtpayamount.Text) Then
MsgBox "Enter Pay Amount Correctly", vbInformation, "Incorrect Entry"
ElseIf txtpaydue.Text = "" Or Not IsNumeric(txtpaydue.Text) Then
MsgBox "Enter Pay Due Correctly", vbInformation, "Incorrect Entry"
Else
stats = True
End If

checkalls = stats
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

If btnsave.Enabled = True Then
exitornot = MsgBox("Exit WithOut Saving The Record", vbQuestion + vbYesNo, "Save Data")
If exitornot = vbNo Then
Cancel = 1
Else
Resume Next
End If
End If

Call MainConClose
Call MainConEstablish
End Sub

Private Sub Image1_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub stfdetailtab_TabSwitch(ByVal iLastActiveTab As Integer)
On Error Resume Next
If stfdetailtab.ActiveTab = 1 Then
Merlin "Staff Salary Entry Is Done Here"
End If
End Sub

Private Sub txtemailaddress_GotFocus()
Merlin "Enter E-Mail Address"
End Sub

Private Sub txtmobnumber_GotFocus()
Merlin "Enter Staff Mobile Number", "DoMagic1"
End Sub

Private Sub txtpayamount_GotFocus()
Merlin "Enter Payment Amount Here"
End Sub

Private Sub txtpaydue_GotFocus()
Merlin "Payment Due To Staff"
End Sub

Private Sub txtperaddress_GotFocus()
Merlin "Enter Permanent Address Here"
End Sub

Private Sub txtphonenumber_GotFocus()
Merlin "Enter Phone Number Here"
End Sub

Private Sub txtqualification_GotFocus()
Merlin "Select Staff Qualification From Here"
End Sub

Private Sub txtrecieptnumber_GotFocus()
Merlin "Enter Payment Receipt Entry"
End Sub

Private Sub txtsalarymon_GotFocus()
Merlin "Enter Staff Salary Per Month Here", "Explain"
End Sub

Private Sub txtstaffids_GotFocus()
Merlin "Enter Staff ID Here"
End Sub

Private Sub txtstaffname_GotFocus()
Merlin "Enter Staff Name Here", "Read"
End Sub

Private Sub txtstaffnames_GotFocus()
Merlin "Enter Staff Name Here"
End Sub

Private Sub txtstaffsalarys_GotFocus()
Merlin "Enter Staff Salary Per Month"
End Sub

Private Sub txtsubjectshandled_GotFocus()
Merlin "Enter Subjects Handled Here"
End Sub

Private Sub txttempaddress_GotFocus()
Merlin "Enter Staff Temporary Address"
End Sub

Private Sub XPButton1_Click()
' Show Search Form
On Error Resume Next
Merlin "Click Me To Search And Enter Staff Detail"
Load Frm_SearchStaffInformation
Frm_SearchStaffInformation.Show
Frm_SearchStaffInformation.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_SearchStaffInformation.Picture = MainMenu.ACPRibbon1.LoadBackground
End Sub
