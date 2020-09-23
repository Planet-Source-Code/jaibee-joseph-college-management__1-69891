VERSION 5.00
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "prjXTab.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{7ECA7ADD-90CB-11D9-B45E-B62B11DAC16E}#1.0#0"; "ButtonXp.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_StudentEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Student Entry"
   ClientHeight    =   7875
   ClientLeft      =   12810
   ClientTop       =   6090
   ClientWidth     =   9600
   Icon            =   "Frm_StudentEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   9600
   Begin MSComDlg.CommonDialog cdlgImage 
      Left            =   -240
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjXTab.XTab studetailtab 
      Height          =   7095
      Left            =   240
      TabIndex        =   31
      Top             =   600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12515
      TabCaption(0)   =   "Academic Information"
      TabContCtrlCnt(0)=   51
      Tab(0)ContCtrlCap(1)=   "btnmovelast"
      Tab(0)ContCtrlCap(2)=   "btnmovenext"
      Tab(0)ContCtrlCap(3)=   "btnmoveprevious"
      Tab(0)ContCtrlCap(4)=   "btnmovefirst"
      Tab(0)ContCtrlCap(5)=   "btnhelp"
      Tab(0)ContCtrlCap(6)=   "btnbrowse"
      Tab(0)ContCtrlCap(7)=   "btncancel"
      Tab(0)ContCtrlCap(8)=   "btndelete"
      Tab(0)ContCtrlCap(9)=   "btnedit"
      Tab(0)ContCtrlCap(10)=   "btnsave"
      Tab(0)ContCtrlCap(11)=   "btnadd"
      Tab(0)ContCtrlCap(12)=   "emercontact"
      Tab(0)ContCtrlCap(13)=   "cmbnationality"
      Tab(0)ContCtrlCap(14)=   "cmbreligion"
      Tab(0)ContCtrlCap(15)=   "cmbcaste"
      Tab(0)ContCtrlCap(16)=   "cmbbloodgroup"
      Tab(0)ContCtrlCap(17)=   "stuname"
      Tab(0)ContCtrlCap(18)=   "cmbsex"
      Tab(0)ContCtrlCap(19)=   "pictext"
      Tab(0)ContCtrlCap(20)=   "rolnumber"
      Tab(0)ContCtrlCap(21)=   "emaaddress"
      Tab(0)ContCtrlCap(22)=   "yearclass"
      Tab(0)ContCtrlCap(23)=   "classsub"
      Tab(0)ContCtrlCap(24)=   "phonumber"
      Tab(0)ContCtrlCap(25)=   "mobnumber"
      Tab(0)ContCtrlCap(26)=   "peraddress"
      Tab(0)ContCtrlCap(27)=   "temaddress"
      Tab(0)ContCtrlCap(28)=   "datbirth"
      Tab(0)ContCtrlCap(29)=   "acayear"
      Tab(0)ContCtrlCap(30)=   "adnumber"
      Tab(0)ContCtrlCap(31)=   "prosnumber"
      Tab(0)ContCtrlCap(32)=   "Label52"
      Tab(0)ContCtrlCap(33)=   "Label51"
      Tab(0)ContCtrlCap(34)=   "Label50"
      Tab(0)ContCtrlCap(35)=   "Label49"
      Tab(0)ContCtrlCap(36)=   "Label48"
      Tab(0)ContCtrlCap(37)=   "Label47"
      Tab(0)ContCtrlCap(38)=   "Label13"
      Tab(0)ContCtrlCap(39)=   "imgHolder"
      Tab(0)ContCtrlCap(40)=   "Label12"
      Tab(0)ContCtrlCap(41)=   "Label1"
      Tab(0)ContCtrlCap(42)=   "Label2"
      Tab(0)ContCtrlCap(43)=   "Label3"
      Tab(0)ContCtrlCap(44)=   "Label4"
      Tab(0)ContCtrlCap(45)=   "Label5"
      Tab(0)ContCtrlCap(46)=   "Label6"
      Tab(0)ContCtrlCap(47)=   "Label7"
      Tab(0)ContCtrlCap(48)=   "Label8"
      Tab(0)ContCtrlCap(49)=   "Label9"
      Tab(0)ContCtrlCap(50)=   "Label10"
      Tab(0)ContCtrlCap(51)=   "Label11"
      TabCaption(1)   =   "Personal Information"
      TabContCtrlCnt(1)=   4
      Tab(1)ContCtrlCap(1)=   "vkFrame1"
      Tab(1)ContCtrlCap(2)=   "picmother"
      Tab(1)ContCtrlCap(3)=   "picfather"
      Tab(1)ContCtrlCap(4)=   "DataGrid1"
      TabCaption(2)   =   "Student Fee Entry"
      TabContCtrlCnt(2)=   18
      Tab(2)ContCtrlCap(1)=   "CmbYear"
      Tab(2)ContCtrlCap(2)=   "vkFrame2"
      Tab(2)ContCtrlCap(3)=   "XPButton6"
      Tab(2)ContCtrlCap(4)=   "RecieptDate"
      Tab(2)ContCtrlCap(5)=   "CmbCourse"
      Tab(2)ContCtrlCap(6)=   "TotalFee"
      Tab(2)ContCtrlCap(7)=   "Remaining"
      Tab(2)ContCtrlCap(8)=   "StudentName"
      Tab(2)ContCtrlCap(9)=   "AdmissionNumber"
      Tab(2)ContCtrlCap(10)=   "RecieptNumber"
      Tab(2)ContCtrlCap(11)=   "Label43"
      Tab(2)ContCtrlCap(12)=   "Label45"
      Tab(2)ContCtrlCap(13)=   "Label44"
      Tab(2)ContCtrlCap(14)=   "Label29"
      Tab(2)ContCtrlCap(15)=   "Label26"
      Tab(2)ContCtrlCap(16)=   "Label25"
      Tab(2)ContCtrlCap(17)=   "Label24"
      Tab(2)ContCtrlCap(18)=   "Label22"
      TabTheme        =   3
      ActiveTabBackStartColor=   16316664
      InActiveTabBackStartColor=   15066597
      InActiveTabBackEndColor=   -2147483626
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
      Begin VB.ComboBox CmbYear 
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
         ItemData        =   "Frm_StudentEntry.frx":076A
         Left            =   -69000
         List            =   "Frm_StudentEntry.frx":078F
         TabIndex        =   45
         Tag             =   "Enter Course Year In The Combo Labelled Year"
         ToolTipText     =   "Select Course Year From Here"
         Top             =   840
         Width           =   2415
      End
      Begin vkUserContolsXP.vkFrame vkFrame2 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   114
         Tag             =   "This Contains Detail Fee Entry Of The Student Fee"
         ToolTipText     =   "Detail Fee Entry"
         Top             =   2760
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7223
         Caption         =   "Student Fee Detail Entry"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin vkUserContolsXP.vkTextBox TotalAmt 
            Height          =   255
            Left            =   3120
            TabIndex        =   131
            ToolTipText     =   "Total Amount"
            Top             =   3600
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            LegendForeColor =   6956042
         End
         Begin Candy.CandyButton Movenext 
            Height          =   255
            Left            =   7440
            TabIndex        =   72
            ToolTipText     =   "Move To Next Record In The Database"
            Top             =   3600
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
         Begin Candy.CandyButton Movelast 
            Height          =   255
            Left            =   8280
            TabIndex        =   73
            ToolTipText     =   "Move To Last Record In The Database"
            Top             =   3600
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
         Begin Candy.CandyButton Moveprevious 
            Height          =   255
            Left            =   1080
            TabIndex        =   70
            ToolTipText     =   "Move To Previous Record In The Database"
            Top             =   3600
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
         Begin Candy.CandyButton Movefirst 
            Height          =   255
            Left            =   240
            TabIndex        =   71
            ToolTipText     =   "Move To First Record In The Database"
            Top             =   3600
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
         Begin Candy.CandyButton Search 
            Height          =   375
            Left            =   7440
            TabIndex        =   69
            ToolTipText     =   "Search Fee Entry"
            Top             =   3120
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
            Caption         =   "Search"
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
         Begin Candy.CandyButton CancelBtn 
            Height          =   375
            Left            =   6000
            TabIndex        =   68
            ToolTipText     =   "Cancel Add New Or Edited Changes"
            Top             =   3120
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
         Begin Candy.CandyButton Delete 
            Height          =   375
            Left            =   4560
            TabIndex        =   67
            ToolTipText     =   "Delete The Current Record"
            Top             =   3120
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
         Begin Candy.CandyButton Edit 
            Height          =   375
            Left            =   3120
            TabIndex        =   66
            ToolTipText     =   "Edit One Existing Record"
            Top             =   3120
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
         Begin Candy.CandyButton SaveBtn 
            Height          =   375
            Left            =   1680
            TabIndex        =   65
            ToolTipText     =   "Save One Edited Or New Record To Database"
            Top             =   3120
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
         Begin Candy.CandyButton AddNew 
            Height          =   375
            Left            =   240
            TabIndex        =   64
            ToolTipText     =   "Add New Record To Database."
            Top             =   3120
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
         Begin VB.TextBox AdmissionFee 
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
            Height          =   315
            Left            =   1680
            TabIndex        =   47
            ToolTipText     =   "Enter Admission Fee"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox CoachingFee 
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
            Height          =   315
            Left            =   1680
            TabIndex        =   48
            ToolTipText     =   "Enter Coaching Fee"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox LibraryFee 
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
            Height          =   315
            Left            =   1680
            TabIndex        =   49
            ToolTipText     =   "Enter Library Fee"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox LabFee 
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
            Height          =   315
            Left            =   1680
            TabIndex        =   50
            ToolTipText     =   "Enter Lab Use Fee"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox SpecialFee 
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
            Height          =   315
            Left            =   1680
            TabIndex        =   51
            ToolTipText     =   "Enter Special Fee"
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox DevelopmentFund 
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
            Height          =   315
            Left            =   4680
            TabIndex        =   52
            ToolTipText     =   "Enter Development Fund"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox Fine 
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
            Height          =   315
            Left            =   4680
            TabIndex        =   53
            ToolTipText     =   "Enter If Any Fine"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox MigrationFee 
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
            Height          =   315
            Left            =   4680
            TabIndex        =   54
            ToolTipText     =   "Migration Fee (Only For Other State Students)"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox EnrolmentFee 
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
            Height          =   315
            Left            =   4680
            TabIndex        =   55
            ToolTipText     =   "Enter Enrolment Fee"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox PhysicalWelfare 
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
            Height          =   315
            Left            =   4680
            TabIndex        =   56
            ToolTipText     =   "Enter Physical Welfare Fee"
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox ComputerFee 
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
            Height          =   315
            Left            =   7440
            TabIndex        =   57
            ToolTipText     =   "Enter Computer Fee"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox CautionDeposit 
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
            Height          =   315
            Left            =   7440
            TabIndex        =   58
            ToolTipText     =   "Enter Caution Deposits"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox Endowment 
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
            Height          =   315
            Left            =   7440
            TabIndex        =   59
            ToolTipText     =   "Enter Endowment"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox ICard 
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
            Height          =   315
            Left            =   7440
            TabIndex        =   60
            ToolTipText     =   "Enter Fee For ICard"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox OtherFee 
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
            Height          =   315
            Left            =   7440
            TabIndex        =   61
            ToolTipText     =   "Enter Other Fee"
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Admission Fee"
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
            TabIndex        =   129
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Coaching Fee"
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
            TabIndex        =   128
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Library Fee"
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
            TabIndex        =   127
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Lab Fee"
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
            TabIndex        =   126
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Special Fee"
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
            TabIndex        =   125
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Development Fund"
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
            Left            =   3120
            TabIndex        =   124
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Fine"
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
            Left            =   3120
            TabIndex        =   123
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Migration Fee"
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
            Left            =   3120
            TabIndex        =   122
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Enrolment Fee"
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
            Left            =   3120
            TabIndex        =   121
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Physical Welfare Fee"
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
            Left            =   3120
            TabIndex        =   120
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Computer Fee"
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
            Left            =   6120
            TabIndex        =   119
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Caution Deposits"
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
            Left            =   6120
            TabIndex        =   118
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Endowment"
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
            Left            =   6120
            TabIndex        =   117
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "I-Card"
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
            Left            =   6120
            TabIndex        =   116
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Other Fee"
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
            Left            =   6120
            TabIndex        =   115
            Top             =   2520
            Width           =   975
         End
      End
      Begin vkUserContolsXP.vkFrame vkFrame1 
         Height          =   3495
         Left            =   -74640
         TabIndex        =   105
         Top             =   720
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6165
         Caption         =   "Enter Family Information Here"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Candy.CandyButton btnsavedetail 
            Height          =   375
            Left            =   360
            TabIndex        =   40
            ToolTipText     =   "Click To Save Family Details"
            Top             =   2640
            Width           =   7575
            _ExtentX        =   13361
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
            Caption         =   "Save Detail"
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
         Begin VB.TextBox fname 
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
            Left            =   1560
            TabIndex        =   32
            Tag             =   "Enter Student Father's Name In The Text Box Labelled Father's Name."
            ToolTipText     =   "Enter Fathers Name Here"
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox fmobilenumber 
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
            Left            =   1560
            TabIndex        =   34
            Tag             =   "Enter Father's Mobile Number In Text Box Labelled Mobile Number."
            ToolTipText     =   "Enter Fathers Mobile Number Here"
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox fofficenumber 
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
            Left            =   1560
            TabIndex        =   35
            Tag             =   "Enter Father's Office Number Here."
            ToolTipText     =   "Enter Fathers Office Number Here"
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox mname 
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
            Left            =   5520
            TabIndex        =   36
            Tag             =   "Enter Mother's Name In The Text Box Labelled Mother's Name."
            ToolTipText     =   "Enter Mothers Name Here"
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox mmobilenumber 
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
            Left            =   5520
            TabIndex        =   38
            Tag             =   "Enter Mother's Mobile Number Here If Any."
            ToolTipText     =   "Enter Mothers Mobile Number Here"
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox mofficenumber 
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
            Left            =   5520
            TabIndex        =   39
            Tag             =   "Enter Mother's Office Number Here If Any."
            ToolTipText     =   "Enter Mothers Office Number Here"
            Top             =   2160
            Width           =   2415
         End
         Begin VB.ComboBox cmbfoccupation 
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
            Height          =   345
            Left            =   1545
            TabIndex        =   33
            Tag             =   "Enter Father's Occupation In The Combo Labelled Occupation."
            ToolTipText     =   "Enter Fathers Occupation Here"
            Top             =   1200
            Width           =   2415
         End
         Begin VB.ComboBox moccupation 
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
            Height          =   345
            Left            =   5505
            TabIndex        =   37
            Tag             =   "Enter Mother's Occupation Here."
            ToolTipText     =   "Enter Mothers Occupation Here"
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
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
            TabIndex        =   113
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
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
            TabIndex        =   112
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label19 
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
            Left            =   360
            TabIndex        =   111
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Mother's Name"
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
            Left            =   4320
            TabIndex        =   110
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Office Number"
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
            TabIndex        =   109
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
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
            Left            =   4320
            TabIndex        =   108
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label23 
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
            Left            =   4320
            TabIndex        =   107
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Office Number"
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
            Left            =   4320
            TabIndex        =   106
            Top             =   2160
            Width           =   1575
         End
      End
      Begin VB.TextBox picmother 
         Height          =   285
         Left            =   -74760
         TabIndex        =   104
         Top             =   7680
         Width           =   735
      End
      Begin VB.TextBox picfather 
         Height          =   285
         Left            =   -74760
         TabIndex        =   103
         Top             =   7200
         Width           =   735
      End
      Begin Candy.CandyButton btnmovelast 
         Height          =   255
         Left            =   8400
         TabIndex        =   28
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
         Left            =   7920
         TabIndex        =   27
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
      Begin Candy.CandyButton btnmoveprevious 
         Height          =   255
         Left            =   7200
         TabIndex        =   26
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
      Begin Candy.CandyButton btnmovefirst 
         Height          =   255
         Left            =   6720
         TabIndex        =   25
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
      Begin Candy.CandyButton btnhelp 
         Height          =   375
         Left            =   4680
         TabIndex        =   29
         Tag             =   "Help Button Is Used To Enable And Disable Me."
         ToolTipText     =   "Application Help"
         Top             =   3480
         Width           =   4095
         _ExtentX        =   7223
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
      Begin Candy.CandyButton btnbrowse 
         Height          =   375
         Left            =   4680
         TabIndex        =   19
         Tag             =   "Browse Button Is Used To Browse Picture Of The Student."
         ToolTipText     =   "Browse For Student Picture"
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
         Left            =   6720
         TabIndex        =   24
         Tag             =   "Using Button Cancel You Can Cancel New Entry Or Editing That Is Not Saved."
         ToolTipText     =   "Cancel Add New Or Edited Changes"
         Top             =   2640
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
      Begin Candy.CandyButton btndelete 
         Height          =   375
         Left            =   6720
         TabIndex        =   23
         Tag             =   "Delete Button Is Used To Delete The Current Record That Appear In The Screen. "
         ToolTipText     =   "Delete The Current Record"
         Top             =   2160
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
         Caption         =   "&Delete"
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
         Left            =   6720
         TabIndex        =   22
         Tag             =   "Edit Button Is Used To Edit The Current Record That Appear In the Screen."
         ToolTipText     =   "Edit One Existing Record"
         Top             =   1680
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
         Caption         =   "&Edit"
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
         Left            =   6720
         TabIndex        =   21
         Tag             =   "Use Save Button To Save Edited Or New Entry. "
         ToolTipText     =   "Save One Edited Or New Record To Database"
         Top             =   1200
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
         Caption         =   "&Save"
         IconHighLiteColor=   0
         CaptionHighLite =   -1  'True
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
      Begin Candy.CandyButton btnadd 
         Height          =   375
         Left            =   6720
         TabIndex        =   20
         Tag             =   "Use Add New Button To Insert New Record To Database."
         ToolTipText     =   "Add New Record To Database."
         Top             =   720
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
         Caption         =   "&Add New"
         IconHighLiteColor=   0
         CaptionHighLite =   -1  'True
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
      Begin VB.TextBox emercontact 
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
         Left            =   6240
         TabIndex        =   15
         Tag             =   "Enter Emergency Contact Number In The Text Box Labelled Emergency Contact."
         ToolTipText     =   "Enter Emergency Contact Here"
         Top             =   5040
         Width           =   2535
      End
      Begin VB.ComboBox cmbnationality 
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
         TabIndex        =   9
         Tag             =   "Select Nationality Of The Student From The Combo Box Labelled Nationality."
         ToolTipText     =   "Select Nationality From Here"
         Top             =   5040
         Width           =   2535
      End
      Begin VB.ComboBox cmbreligion 
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
         ItemData        =   "Frm_StudentEntry.frx":080B
         Left            =   1920
         List            =   "Frm_StudentEntry.frx":080D
         TabIndex        =   8
         Tag             =   "Select Student Religion From The Combo Box Labelled Religion."
         ToolTipText     =   "Select Religion From Here"
         Top             =   4560
         Width           =   2535
      End
      Begin VB.ComboBox cmbcaste 
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
         TabIndex        =   7
         Tag             =   "Select Caste Of The Student From The Combo Box Named Caste."
         ToolTipText     =   "Select Caste From Here"
         Top             =   4080
         Width           =   2535
      End
      Begin VB.ComboBox cmbbloodgroup 
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
         ItemData        =   "Frm_StudentEntry.frx":080F
         Left            =   1920
         List            =   "Frm_StudentEntry.frx":082B
         TabIndex        =   6
         Tag             =   "Select Blood Group Of The Student From The Combo Box Named Blood Group."
         ToolTipText     =   "Select Blood Group From Here"
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox stuname 
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
         TabIndex        =   3
         Tag             =   "Enter Name Of The Student In The Text Box Labelled Student Name."
         ToolTipText     =   "Enter Student Name Here"
         Top             =   2160
         Width           =   2535
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
         ItemData        =   "Frm_StudentEntry.frx":0891
         Left            =   1920
         List            =   "Frm_StudentEntry.frx":089B
         TabIndex        =   4
         Tag             =   "Select Sex Of The Student From The Combo Box Labelled Sex."
         ToolTipText     =   "Select Student Sex From Here"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox pictext 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         TabIndex        =   96
         Top             =   7440
         Visible         =   0   'False
         Width           =   6735
      End
      Begin ButtonXp.XPButton XPButton6 
         Height          =   315
         Left            =   -70680
         TabIndex        =   95
         ToolTipText     =   "Click Me To Load Search Form"
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin MSComCtl2.DTPicker RecieptDate 
         Height          =   375
         Left            =   -69000
         TabIndex        =   46
         Tag             =   "Enter Reciept Date In The Date Picker Labelled Reciept Date"
         ToolTipText     =   "Select Receipt Date"
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16508095
         Format          =   48758785
         CurrentDate     =   39405
      End
      Begin VB.ComboBox CmbCourse 
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
         Left            =   -72840
         TabIndex        =   44
         Tag             =   "Enter Course Name In The Combo Box Labelled Course"
         ToolTipText     =   "Enter Student Course Here"
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox TotalFee 
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
         Left            =   -69000
         TabIndex        =   62
         Tag             =   "Enter Total Fee In The Text Box Labelled Total Fee"
         ToolTipText     =   "Enter Total Fee"
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Remaining 
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
         Left            =   -69000
         TabIndex        =   63
         Tag             =   "Enter Remaining Amount In The Text Box Labelled Remaining"
         ToolTipText     =   "Enter Remaining Fee"
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox StudentName 
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
         Left            =   -72840
         TabIndex        =   43
         Tag             =   "Enter Stuent Name In The Text Box Labelled Student Name"
         ToolTipText     =   "Enter Student Name Here"
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox AdmissionNumber 
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
         Left            =   -72840
         TabIndex        =   42
         Tag             =   "Enter Admission Number In The Text Box Labelled Admission Number"
         ToolTipText     =   "Enter Admission Number Here"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox RecieptNumber 
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
         Left            =   -72840
         TabIndex        =   41
         Tag             =   "Enter Reciept Number In The Text Box Labelled Reciept Number"
         ToolTipText     =   "Enter Receipt Number Here"
         Top             =   840
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Frm_StudentEntry.frx":08AD
         Height          =   2415
         Left            =   -74640
         TabIndex        =   87
         ToolTipText     =   "Existing Records"
         Top             =   4440
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4260
         _Version        =   393216
         BackColor       =   16508095
         HeadLines       =   1
         RowHeight       =   18
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Existing Records"
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "Admission_Number"
            Caption         =   "Admission Number"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Fathers_Name"
            Caption         =   "Father's Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Fathers_Occupation"
            Caption         =   "Father's Occupation"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "FMobile_Number"
            Caption         =   "Father's Mobile Number"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "FOffice_Number"
            Caption         =   "Father's Office Number"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Mothers_Name"
            Caption         =   "Mother's Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Mothers_Occupation"
            Caption         =   "Mother's Occupation"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "MMobile_Number"
            Caption         =   "Mother's Mobile Number"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "MOffice_Number"
            Caption         =   "Mother's Office Number"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.TextBox rolnumber 
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
         TabIndex        =   12
         Tag             =   "Enter Class Roll Number Of The Student In The Text Box Labelled Roll Number."
         ToolTipText     =   "Enter Roll Number Here"
         Top             =   6480
         Width           =   2535
      End
      Begin VB.TextBox emaaddress 
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
         Left            =   6240
         TabIndex        =   18
         Tag             =   "Enter E-Mail Address Of The Student In The Text Box Labelled E-Mail Address."
         ToolTipText     =   "Enter E-Mail Address Here"
         Top             =   6480
         Width           =   2535
      End
      Begin VB.ComboBox yearclass 
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
         ItemData        =   "Frm_StudentEntry.frx":08C2
         Left            =   1920
         List            =   "Frm_StudentEntry.frx":08E7
         TabIndex        =   11
         Tag             =   "Select Course Year From The Combo Labelled Year."
         ToolTipText     =   "Select Course Year From Here"
         Top             =   6000
         Width           =   2535
      End
      Begin VB.ComboBox classsub 
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
         TabIndex        =   10
         Tag             =   "Select Name Of The Course From the Combo Labelled Course Name. "
         ToolTipText     =   "Select Course Name From Here"
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox phonumber 
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
         Left            =   6240
         TabIndex        =   16
         Tag             =   "Enter House Phone Number Of The Student In The Text Box Labelled Phone Number."
         ToolTipText     =   "Enter Phone Number Here"
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox mobnumber 
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
         Left            =   6240
         TabIndex        =   17
         Tag             =   "Enter Mobile Number Of The Student If Any In The Text Box Labelled Mobile Number."
         ToolTipText     =   "Enter Mobile Number Here"
         Top             =   6000
         Width           =   2535
      End
      Begin VB.TextBox peraddress 
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
         Left            =   6240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Tag             =   "Enter Permanent Address Of The Student."
         ToolTipText     =   "Enter Permanent Addess Here"
         Top             =   4560
         Width           =   2535
      End
      Begin VB.TextBox temaddress 
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
         Left            =   6240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Tag             =   "Enter Temporary Address Of The Student In Case If He Is Staying In Any Other Address."
         ToolTipText     =   "Enter Temporary Address Here"
         Top             =   4080
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker datbirth 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Tag             =   "Select Student DOB From The Date Picker Labelled Date Of Birth."
         ToolTipText     =   "Select Student DOB From Here"
         Top             =   3120
         Width           =   2535
         _ExtentX        =   4471
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
         Format          =   48758785
         CurrentDate     =   39377
      End
      Begin VB.TextBox acayear 
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
         TabIndex        =   2
         Tag             =   "Enter Academic Year In The Text Box Named Academic Year."
         ToolTipText     =   "Enter Academic Year Here"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox adnumber 
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
         TabIndex        =   1
         Tag             =   "Enter Admission In The Text Box Labelled Admission Number, This Field Is Compulsory."
         ToolTipText     =   "Enter Admission Number Here"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox prosnumber 
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
         TabIndex        =   0
         Tag             =   "Enter Prospectus Number In The Text Box Labelled Prospectus Number."
         ToolTipText     =   "Enter Prospectus Number Here"
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   -69960
         TabIndex        =   130
         Tag             =   "Enter Course Year In The Combo Box Labelled Year"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "Emergency Contact"
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
         Left            =   4680
         TabIndex        =   102
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
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
         TabIndex        =   101
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
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
         TabIndex        =   100
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Caste"
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
         TabIndex        =   99
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Group"
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
         TabIndex        =   98
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label47 
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
         TabIndex        =   97
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fee"
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
         Left            =   -69960
         TabIndex        =   94
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining"
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
         Left            =   -69960
         TabIndex        =   93
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Reciept Date"
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
         Left            =   -69960
         TabIndex        =   92
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
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
         Left            =   -74280
         TabIndex        =   91
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name"
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
         Left            =   -74280
         TabIndex        =   90
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Number"
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
         Left            =   -74280
         TabIndex        =   89
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Number"
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
         Left            =   -74280
         TabIndex        =   88
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Roll Number"
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
         TabIndex        =   85
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Image imgHolder 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label12 
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
         Left            =   4680
         TabIndex        =   84
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Prospectus Number"
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
         TabIndex        =   83
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Number"
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
         TabIndex        =   82
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name"
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
         TabIndex        =   81
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         Left            =   4680
         TabIndex        =   80
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
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
         TabIndex        =   79
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Course Name"
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
         TabIndex        =   78
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         TabIndex        =   77
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Academic Year"
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
         TabIndex        =   76
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label9 
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
         Left            =   4680
         TabIndex        =   75
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label10 
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
         Left            =   4680
         TabIndex        =   74
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         Left            =   4680
         TabIndex        =   30
         Top             =   6480
         Width           =   1335
      End
   End
   Begin VB.Image imghelp 
      Height          =   360
      Left            =   9000
      Picture         =   "Frm_StudentEntry.frx":0963
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Student Information, Student Family Information and Fee Information Here."
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
      Left            =   720
      TabIndex        =   86
      Top             =   240
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Picture         =   "Frm_StudentEntry.frx":10CD
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Frm_StudentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim saveflag As Boolean
Dim saveflagfee As Boolean
Dim str As String
Dim str1 As String
Dim strf As String
Dim strcf As String
Dim X As String
Private Sub acayear_GotFocus()
Merlin "Enter Current Academic Year Here", "Explain"
End Sub

Private Sub AddNew_Click()
On Error GoTo label

Call LockTxtFee(False)
RecieptNumber.SetFocus
Call lockbtnfee(True)
Call lockbtnfees(False)
Call clearfee
saveflagfee = True
XPButton6.Enabled = True
RecieptDate.Value = Now
Merlin "Add New Receipt Entry"

Exit Sub
label:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub AdmissionFee_LostFocus()
On Error Resume Next
If AdmissionFee.Text = "" Then
AdmissionFee.Text = 0
End If
End Sub

Private Sub showdata()
 If studentrec.EOF = False And studentrec.BOF = False Then
          prosnumber.Text = studentrec.Fields(0)
          adnumber.Text = studentrec.Fields(1)
          acayear.Text = studentrec.Fields(2)
          stuname.Text = studentrec.Fields(3)
          cmbsex.Text = studentrec.Fields(4)
          datbirth.Value = studentrec.Fields(5)
          cmbbloodgroup.Text = studentrec.Fields(6)
          cmbcaste.Text = studentrec.Fields(7)
          cmbreligion.Text = studentrec.Fields(8)
          cmbnationality.Text = studentrec.Fields(9)
          classsub.Text = studentrec.Fields(10)
          yearclass.Text = studentrec.Fields(11)
          rolnumber.Text = studentrec.Fields(12)
          temaddress.Text = studentrec.Fields(13)
          peraddress.Text = studentrec.Fields(14)
          emercontact.Text = studentrec.Fields(15)
          phonumber.Text = studentrec.Fields(16)
          mobnumber.Text = studentrec.Fields(17)
          emaaddress.Text = studentrec.Fields(18)
          pictext.Text = studentrec.Fields(19)
          If copyf.FileExists(pictext.Text) Then
          imgHolder.Picture = LoadPicture(pictext.Text)
          ElseIf pictext.Text = "" Then
          imgHolder.Picture = LoadPicture("")
          End If
 End If
End Sub

Private Sub adnumber_GotFocus()
Merlin "Admission Number Is Entered Here", "Read"
End Sub

Private Sub btnadd_Click()
On Error GoTo lable

Call cleardata
imgHolder.Picture = LoadPicture(pictext.Text)
Call locktextbox(False)
saveflag = True
Call lockbtn(True)
Call disbtnstu(False)
prosnumber.SetFocus
acayear.Text = Format(Now, "YYYY")
Merlin "Click Me To Add New Record To Database", "DoMagic1"

Exit Sub
lable:
MsgBox "Error Occured While Adding New, Sorry For The Interruption", vbInformation, "Error In Saving"
End Sub

Private Sub btnbrowse_Click()
On Error Resume Next
Merlin "Click Me To Browse Record"
Call browseimage(cdlgImage, imgHolder, pictext)
End Sub

' Function To Lock and Unlock Text and Combo Boxses
Private Function locktextbox(locktext As Boolean)

' Lock Text Box
prosnumber.Locked = locktext
adnumber.Locked = locktext
acayear.Locked = locktext
stuname.Locked = locktext
temaddress.Locked = locktext
peraddress.Locked = locktext
rolnumber.Locked = locktext
phonumber.Locked = locktext
mobnumber.Locked = locktext
emaaddress.Locked = locktext
emercontact.Locked = locktext

'Lock Combo Box
classsub.Locked = locktext
yearclass.Locked = locktext
cmbbloodgroup.Locked = locktext
cmbcaste.Locked = locktext
cmbreligion.Locked = locktext
cmbnationality.Locked = locktext
cmbsex.Locked = locktext

End Function

Private Sub btnCancel_Click()
On Error Resume Next

btnadd.Enabled = True
Call disbtnstu(True)
Call cleardata
Call locktextbox(True)
Call lockbtn(False)
Call checkbtn
Merlin "Click Me To Cancel Add New Or Edit"

If studentrec.BOF And studentrec.EOF Then
MsgBox "No Existing Record, Insert New Record", vbInformation, "No Record"
Else
studentrec.Movefirst
Call showdata
End If
End Sub

Private Sub btndelete_Click()
On Error GoTo lable

Merlin "Click Me To Delete Current Record"
Dim admnumber As String
admnumber = adnumber.Text

If MsgBox("Execution Of Command Will Delete Current Datarecord" & vbCrLf & "Are You Sure You Wan't To Delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
str = "DELETE FROM StudentInformation WHERE "
str = str & "Admission_Number = "
str = str & CDbl(adnumber.Text)
studentcon.Execute str
studentrec.Requery

strf = "DELETE FROM FamilyInformation WHERE "
strf = strf & "Admission_Number = "
strf = strf & CDbl(admnumber)
familycon.Execute strf
familyrec.Requery
DataGrid1.ReBind

If pictext.Text <> "" Then
copyf.DeleteFile pictext.Text, True
End If

MsgBox "Record Deleted Sucessfully.", vbInformation, "Delete Record"

If studentrec.BOF And studentrec.EOF Then
Call cleardata
MsgBox ("The Previous Record Was Last Record."), vbInformation, "Last Record"
Call checkbtn
imgHolder.Picture = LoadPicture("")
Else
studentrec.Movenext
If studentrec.EOF Then
studentrec.Movelast
End If
Call showdata
End If

End If
Exit Sub
lable:
MsgBox "No Existing Record, Insert New Record", vbInformation, "Error Occured"
End Sub

Private Sub btnedit_Click()
On Error GoTo label

btnadd.Enabled = False
Call disbtnstu(False)
Call lockbtn(True)
Call locktextbox(False)
saveflag = False
prosnumber.SetFocus
Merlin "Click Me To Edit Current Record", "DoMagic1"

Exit Sub
label:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnhelp_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub btnmovefirst_Click()
On Error GoTo GoFirstError

studentrec.Movefirst
'show thw current data record
Call showdata
 
Exit Sub

GoFirstError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub btnmovelast_Click()
On Error GoTo GoLastError

studentrec.Movelast
'show thw current data record
Call showdata
Exit Sub

GoLastError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub btnmovenext_Click()
On Error GoTo GoNextError
'lblStatus.Caption = "               Move       >"
  
If Not studentrec.EOF Then studentrec.Movenext
If studentrec.EOF And studentrec.RecordCount > 0 Then
'moved off the end so go back
studentrec.Movelast
    
End If
'show thw current data record
Call showdata
  
Exit Sub
GoNextError:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnmoveprevious_Click()
On Error GoTo GoPrevError
  
If Not studentrec.BOF Then studentrec.Moveprevious
If studentrec.BOF And studentrec.RecordCount > 0 Then
    
'moved off the end so go back
studentrec.Moveprevious
 
End If
'show thw current data record
Call showdata
Exit Sub

GoPrevError:
If Err.Number = 3021 Then
MsgBox ("This Is First Record."), vbInformation, "First Record"
studentrec.Movenext
ElseIf Err.Number <> 0 Then
MsgBox Err.Description, vbInformation, "Error Occured"
End If
End Sub

Private Sub btnsave_Click()
On Error GoTo lable

Merlin "Click Me To Save Edited Or New Record"
If checkall = True Then

If pictext.Text = "" And saveflag = True Then
X = MsgBox("Do You Want To Enter Picture Of The Student", vbInformation + vbYesNo, "Picture Entry")
If X = vbYes Then
btnbrowse_Click
Exit Sub
Else
GoTo savequery
End If
End If

If pictext.Text = strfn Then
pictext.Text = App.Path & "\Images\" & adnumber.Text & ".JPG"
copyf.CopyFile strfn, pictext.Text, True
ElseIf strfn = "" And pictext = "" Then
GoTo savequery
End If

savequery:

If saveflag = True Then ' insert new record
str = "INSERT INTO StudentInformation"
str = str & "(Prospectus_Number, Admission_Number, Academic_Year, Student_Name, Sex, Date_of_Birth, Blood_Group, Caste, Religion, Nationality, Course_Name, Year_Course, Roll_Number, Temporary_Address, Permanent_Address, Emergency_Contact, Phone_Number, Mobile_Number, EMail_Address, Pic_Student) "
str = str & "VALUES('" & Trim$(prosnumber.Text) & "', "
str = str & CDbl(adnumber.Text) & ", "
str = str & "'" & Trim$(acayear.Text) & "', "
str = str & "'" & Trim$(stuname.Text) & "', "
str = str & "'" & Trim$(cmbsex.Text) & "', "
str = str & "'" & datbirth.Value & "', "
str = str & "'" & Trim$(cmbbloodgroup.Text) & "', "
str = str & "'" & Trim$(cmbcaste.Text) & "', "
str = str & "'" & Trim$(cmbreligion.Text) & "', "
str = str & "'" & Trim$(cmbnationality.Text) & "', "
str = str & "'" & Trim$(classsub.Text) & "', "
str = str & "'" & Trim$(yearclass.Text) & "', "
str = str & CDbl(rolnumber.Text) & ", "
str = str & "'" & Trim$(temaddress.Text) & "', "
str = str & "'" & Trim$(peraddress.Text) & "', "
str = str & "'" & Trim$(emercontact.Text) & "', "
str = str & "'" & Trim$(phonumber.Text) & "', "
str = str & "'" & Trim$(mobnumber.Text) & "', "
str = str & "'" & Trim$(emaaddress.Text) & "', "
str = str & "'" & Trim$(pictext.Text) & "')"
studentcon.Execute str
Else ' for editing the record
str = "UPDATE StudentInformation SET "
str = str & "Prospectus_Number='" & Trim$(prosnumber.Text) & "',"
str = str & "Admission_Number=" & CDbl(adnumber.Text) & ","
str = str & "Academic_Year='" & Trim$(acayear.Text) & "',"
str = str & "Student_Name='" & Trim$(stuname.Text) & "',"
str = str & "Sex='" & Trim$(cmbsex.Text) & "',"
str = str & "Date_of_Birth='" & datbirth.Value & "',"
str = str & "Blood_Group='" & Trim$(cmbbloodgroup.Text) & "',"
str = str & "Caste='" & Trim$(cmbcaste.Text) & "',"
str = str & "Religion='" & Trim$(cmbreligion.Text) & "',"
str = str & "Nationality='" & Trim$(cmbnationality.Text) & "',"
str = str & "Course_Name='" & Trim$(classsub.Text) & "',"
str = str & "Year_Course='" & Trim$(yearclass.Text) & "',"
str = str & "Roll_Number=" & CDbl(rolnumber.Text) & ","
str = str & "Temporary_Address='" & Trim$(temaddress.Text) & "',"
str = str & "Permanent_Address='" & Trim$(peraddress.Text) & "',"
str = str & "Emergency_Contact='" & Trim$(emercontact.Text) & "',"
str = str & "Phone_Number='" & Trim$(phonumber.Text) & "',"
str = str & "Mobile_Number='" & Trim$(mobnumber.Text) & "',"
str = str & "EMail_Address='" & Trim$(emaaddress.Text) & "',"
str = str & "Pic_Student='" & Trim$(pictext.Text) & "'"
str = str & " WHERE Admission_Number=" & CDbl(adnumber.Text)
studentcon.Execute str
btnadd.Enabled = True
End If

studentrec.Requery
studentrec.Movelast
'show thw current data record
Call showdata
 'message for status of mode

MsgBox ("Record Has Been Sucessfully Saved."), vbInformation, "Saving Record"

Call disbtnstu(True)
Call lockbtn(False)
Call locktextbox(True)
Call checkbtn

End If
Exit Sub
lable:

If Err.Number = -2147217900 Then
MsgBox ("Admission Number Already Exist,Please Enter Another Number"), vbCritical, "Admission Number Exist"
Else
MsgBox Err.Description, vbInformation, "Error Occured"
End If
End Sub

Private Sub btnsavedetail_Click()
On Error GoTo label

If checkfamilydata = True Then

str1 = "INSERT INTO FamilyInformation"
str1 = str1 & "(Admission_Number, Fathers_Name, Fathers_Occupation, FMobile_Number, FOffice_Number, Mothers_Name, Mothers_Occupation, MMobile_Number, MOffice_Number) "
str1 = str1 & "VALUES(" & CDbl(adnumber.Text) & ", "
str1 = str1 & "'" & Trim$(fname.Text) & "', "
str1 = str1 & "'" & Trim$(cmbfoccupation.Text) & "', "
str1 = str1 & "'" & Trim$(fmobilenumber.Text) & "', "
str1 = str1 & "'" & Trim$(fofficenumber.Text) & "', "
str1 = str1 & "'" & Trim$(mname.Text) & "', "
str1 = str1 & "'" & Trim$(moccupation.Text) & "', "
str1 = str1 & "'" & Trim$(mmobilenumber.Text) & "', "
str1 = str1 & "'" & Trim$(mofficenumber.Text) & "') "

studentcon.Execute str1
familyrec.Requery
familyrec.Movelast
DataGrid1.ReBind

Call clearfamilydata
Call locktxtfamily(True)

btnsavedetail.Enabled = False
End If

Exit Sub
label:
MsgBox Err.Description, vbInformation, "Error Occured"

End Sub

Private Function clearfamilydata()
fname.Text = ""
cmbfoccupation.Text = ""
fmobilenumber.Text = ""
fofficenumber.Text = ""
mname.Text = ""
moccupation.Text = ""
mmobilenumber.Text = ""
mofficenumber.Text = ""
End Function

Private Sub CancelBtn_Click()
On Error GoTo label

Merlin "Cancel Save Or Edit"
AddNew.Enabled = True
Call lockbtnfees(True)
Call checkbtnfee
Call lockbtnfee(False)
Call LockTxtFee(True)
XPButton6.Enabled = False

If feesrec.BOF And feesrec.EOF Then
MsgBox "No Existing Record, Insert New Record", vbInformation, "No Record"
Else
feesrec.Movefirst
Call showdatafee
End If

Exit Sub
label:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub CautionDeposit_LostFocus()
On Error Resume Next
If CautionDeposit.Text = "" Then
CautionDeposit.Text = 0
End If
End Sub

Private Sub classsub_GotFocus()
Merlin "Select Student Class From Here"
End Sub

Private Sub cmbbloodgroup_GotFocus()
Merlin "Select Student Blood Group From Here"
End Sub

Private Sub cmbcaste_GotFocus()
Merlin "Select Caste Of The Student", "Read"
End Sub

Private Sub cmbfoccupation_GotFocus()
Merlin "Select Father's Occupation From Here"
End Sub

Private Sub cmbnationality_GotFocus()
Merlin "Select Student Nationality From Here"
End Sub

Private Sub cmbreligion_gotfocus()
Merlin "Select Student Religion From Here"
End Sub

Private Sub cmbsex_GotFocus()
Merlin "Select Sex Of The Student From Here"
End Sub

Private Sub CoachingFee_LostFocus()
On Error Resume Next
If CoachingFee.Text = "" Then
CoachingFee.Text = 0
End If
End Sub

Private Sub ComputerFee_LostFocus()
On Error Resume Next
If ComputerFee.Text = "" Then
ComputerFee.Text = 0
End If
End Sub

Private Sub datbirth_GotFocus()
Merlin "Select DOB Of The Student From Here"
End Sub

Private Sub delete_Click()
On Error GoTo label

Merlin "This Will Delete Current Record"
If MsgBox("Execution Of Command Will Delete Current Datarecord" & vbCrLf & "Are You Sure You Wan't To Delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
strcf = "DELETE FROM FeesInformation WHERE "
strcf = strcf & "Reciept_Number = "
strcf = strcf & CDbl(RecieptNumber.Text)
feescon.Execute strcf
feesrec.Requery

MsgBox "Record Deleted Sucessfully.", vbInformation, "Delete Record"

If feesrec.BOF And feesrec.EOF Then
Call clearfee
MsgBox ("The Previous Record Was Last Record."), vbInformation, "Last Record"
Call checkbtnfee
Else
feesrec.Movenext
If feesrec.EOF Then
feesrec.Movelast
End If
Call showdatafee
End If
End If

Exit Sub
label:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub DevelopmentFund_LostFocus()
On Error Resume Next
If DevelopmentFund.Text = "" Then
DevelopmentFund.Text = 0
End If
End Sub

Private Sub Edit_Click()
On Error GoTo label

AddNew.Enabled = False
Call lockbtnfees(False)
Call LockTxtFee(False)
RecieptNumber.SetFocus
Call lockbtnfee(True)
saveflagfee = False
XPButton6.Enabled = True
Merlin "Edit Current Record"

Exit Sub
label:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub emaaddress_GotFocus()
Merlin "Enter EMail Address Here"
End Sub

Private Sub emercontact_GotFocus()
Merlin "Enter Emergency Contact Number"
End Sub

Private Sub Endowment_LostFocus()
On Error Resume Next
If Endowment.Text = "" Then
Endowment.Text = 0
End If
End Sub

Private Sub EnrolmentFee_LostFocus()
On Error Resume Next
If EnrolmentFee.Text = "" Then
EnrolmentFee.Text = 0
End If
End Sub

Private Sub Fine_LostFocus()
On Error Resume Next
If Fine.Text = "" Then
Fine.Text = 0
End If
End Sub

Private Sub fmobilenumber_GotFocus()
Merlin "Enter Father's Mobile Number Here"
End Sub

Private Sub fname_GotFocus()
Merlin "Enter Father's Name Here", "Read"
End Sub

Private Sub fofficenumber_GotFocus()
Merlin "Enter Father's Office Number Here"
End Sub

Private Sub Form_Load()
On Error GoTo message

Me.Top = 50
Me.Left = 50

'call lock text function
Call locktextbox(True)
Call locktxtfamily(True)
Call LockTxtFee(True)
XPButton6.Enabled = False

Call lockbtn(False)
Call lockbtnfee(False)

Call showdata
Call showdatafee

Call checkbtn
Call checkbtnfee

Set DataGrid1.DataSource = familyrec
DataGrid1.ReBind

GlobalCaste.Movefirst
Do While Not GlobalCaste.BOF And Not GlobalCaste.EOF
   cmbcaste.AddItem GlobalCaste(1).Value
   GlobalCaste.Movenext
Loop

GlobalReligion.Movefirst
Do While Not GlobalReligion.BOF And Not GlobalReligion.EOF
   cmbreligion.AddItem GlobalReligion(1).Value
   GlobalReligion.Movenext
Loop

GlobalNationality.Movefirst
Do While Not GlobalNationality.BOF And Not GlobalNationality.EOF
   cmbnationality.AddItem GlobalNationality(1).Value
   GlobalNationality.Movenext
Loop

GlobalOccupationF.Movefirst
Do While Not GlobalOccupationF.BOF And Not GlobalOccupationF.EOF
   cmbfoccupation.AddItem GlobalOccupationF(1).Value
   GlobalOccupationF.Movenext
Loop

GlobalOccupationM.Movefirst
Do While Not GlobalOccupationM.BOF And Not GlobalOccupationM.EOF
   moccupation.AddItem GlobalOccupationM(1).Value
   GlobalOccupationM.Movenext
Loop

coursestudent.Movefirst
Do While Not coursestudent.BOF And Not coursestudent.EOF
   classsub.AddItem coursestudent(1).Value
   cmbcourse.AddItem coursestudent(1).Value
   coursestudent.Movenext
Loop

Merlin "This Is The Form Where Student Information Is Entered", "Explain"
Exit Sub
message:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Function checkall() As Boolean
Dim stat As Boolean

stat = False

If adnumber.Text = "" Then
MsgBox "Admission Number Is Compulsory Field", vbInformation, "Empty Field"
ElseIf IsNumeric(adnumber.Text) = False Then
MsgBox "Only Numbers Are Allowed In Admission Number", vbInformation, "Invalid Entry"
ElseIf acayear.Text = "" Then
MsgBox "Academic Year Is Compulsory Field", vbInformation, "Empty Field"
ElseIf stuname.Text = "" Then
MsgBox "Student Name Is Compulsory Field", vbInformation, "Empty Field"
ElseIf cmbsex.Text = "" Then
MsgBox "Student Sex Is Compulsory Field", vbInformation, "Empty Field"
ElseIf classsub.Text = "" Then
MsgBox "Course Name Is Compulsory Field", vbInformation, "Empty Field"
ElseIf yearclass.Text = "" Then
MsgBox "Course Year Is Compulsory Field", vbInformation, "Empty Field"
ElseIf rolnumber.Text = "" Then
MsgBox "Roll Number Is Compulsory Field", vbInformation, "Empty Field"
ElseIf peraddress.Text = "" Then
MsgBox "Permanent Address Is Compulsory Field", vbInformation, "Empty Field"
ElseIf ValidateEmail(emaaddress.Text) = False Then
MsgBox "Wrong E-Mail Entry, Check The Mail Address", vbInformation, "Wrong Entry"
ElseIf IsNumeric(rolnumber.Text) = False Then
MsgBox "Only Numbers Are Allowed In Roll Number", vbInformation, "Invalid Entry"
ElseIf checkfamilydatavalid = False Then
MsgBox "Enter Family Information", vbInformation, "Enter Family Data"
Else
stat = True
End If

checkall = stat

End Function

Private Function lockbtn(sta As Boolean)
btnbrowse.Enabled = sta
btncancel.Enabled = sta
btnsave.Enabled = sta
datbirth.Enabled = sta
End Function

Private Function cleardata()
adnumber.Text = ""
prosnumber.Text = ""
acayear.Text = ""
stuname.Text = ""
temaddress.Text = ""
peraddress.Text = ""
classsub.Text = ""
yearclass.Text = ""
rolnumber.Text = ""
phonumber.Text = ""
mobnumber.Text = ""
emaaddress.Text = ""
emercontact.Text = ""
pictext.Text = ""
cmbbloodgroup.Text = ""
cmbcaste.Text = ""
cmbreligion.Text = ""
cmbnationality.Text = ""
cmbsex.Text = ""
End Function

Private Function checkbtn()
If studentrec.RecordCount = 0 Then
btnedit.Enabled = False
btndelete.Enabled = False
btnmovefirst.Enabled = False
btnmoveprevious.Enabled = False
btnmovenext.Enabled = False
btnmovelast.Enabled = False
Else
btnedit.Enabled = True
btndelete.Enabled = True
btnmovefirst.Enabled = True
btnmoveprevious.Enabled = True
btnmovenext.Enabled = True
btnmovelast.Enabled = True
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

If btnsave.Enabled = True Then
g = MsgBox("Exit Without Saving The Record", vbQuestion + vbYesNo, "Save Record")
If g = vbYes Then
Unload Me
Else: Cancel = 1
End If
End If

Call MainConClose
Call MainConEstablish
End Sub

Private Sub HelpBtn_Click()

End Sub

Private Sub ICard_LostFocus()
On Error Resume Next
If ICard.Text = "" Then
ICard.Text = 0
End If
End Sub

Private Function locktxtfamily(s As Boolean)
fname.Locked = s
cmbfoccupation.Locked = s
fmobilenumber.Locked = s
fofficenumber.Locked = s
mname.Locked = s
moccupation.Locked = s
mmobilenumber.Locked = s
mofficenumber.Locked = s
End Function

Private Sub imghelp_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub LabFee_LostFocus()
On Error Resume Next
If LabFee.Text = "" Then
LabFee.Text = 0
End If
End Sub

Private Sub LibraryFee_LostFocus()
On Error Resume Next
If LibraryFee.Text = "" Then
LibraryFee.Text = 0
End If
End Sub

Private Sub MigrationFee_LostFocus()
On Error Resume Next
If MigrationFee.Text = "" Then
MigrationFee.Text = 0
End If
End Sub

Private Sub mmobilenumber_GotFocus()
Merlin "Enter Mother's Mobile Number"
End Sub

Private Sub mname_GotFocus()
Merlin "Enter Mother's Name Here"
End Sub

Private Sub mobnumber_GotFocus()
Merlin "Enter Mobile Number Here", "DoMagic3"
End Sub

Private Sub moccupation_GotFocus()
Merlin "Select Mother's Occupation"
End Sub

Private Sub mofficenumber_GotFocus()
Merlin "Enter Mother's Office Number"
End Sub

Private Sub Movefirst_Click()
On Error GoTo GoFirstError

feesrec.Movefirst
'show thw current data record
Call showdatafee
 
Exit Sub

GoFirstError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub Movelast_Click()
On Error GoTo GoLastError

feesrec.Movelast
'show thw current data record
Call showdatafee
Exit Sub

GoLastError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub Movenext_Click()
On Error GoTo GoNextError
'lblStatus.Caption = "               Move       >"
  
If Not feesrec.EOF Then feesrec.Movenext
If feesrec.EOF And feesrec.RecordCount > 0 Then
'moved off the end so go back
feesrec.Movelast
    
End If
'show thw current data record
Call showdatafee
  
Exit Sub
GoNextError:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub Moveprevious_Click()
On Error GoTo GoPrevError
  
If Not feesrec.BOF Then feesrec.Moveprevious
If feesrec.BOF And feesrec.RecordCount > 0 Then
    
'moved off the end so go back
feesrec.Moveprevious
 
End If
'show thw current data record
Call showdatafee
Exit Sub

GoPrevError:
If Err.Number = 3021 Then
MsgBox ("This Is First Record."), vbInformation, "First Record"
feesrec.Movenext
ElseIf Err.Number <> 0 Then
MsgBox Err.Description, vbInformation, "Error Occured"
End If
End Sub

Private Sub OtherFee_LostFocus()
On Error GoTo label
If OtherFee.Text = "" Then
OtherFee.Text = 0
TotalAmt.Text = Val(AdmissionNumber.Text) + Val(CoachingFee.Text) + Val(LibraryFee.Text) + Val(LabFee.Text) + Val(SpecialFee.Text) + Val(DevelopmentFund.Text) + Val(Fine.Text) + Val(MigrationFee.Text) + Val(EnrolmentFee.Text) + Val(PhysicalWelfare.Text) + Val(ComputerFee.Text) + Val(CautionDeposit.Text) + Val(Endowment.Text) + Val(ICard.Text) + Val(OtherFee.Text)
Else
TotalAmt.Text = Val(AdmissionNumber.Text) + Val(CoachingFee.Text) + Val(LibraryFee.Text) + Val(LabFee.Text) + Val(SpecialFee.Text) + Val(DevelopmentFund.Text) + Val(Fine.Text) + Val(MigrationFee.Text) + Val(EnrolmentFee.Text) + Val(PhysicalWelfare.Text) + Val(ComputerFee.Text) + Val(CautionDeposit.Text) + Val(Endowment.Text) + Val(ICard.Text) + Val(OtherFee.Text)
End If
Exit Sub
label:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub peraddress_GotFocus()
Merlin "Enter Permanent Address Here"
End Sub

Private Sub phonumber_GotFocus()
Merlin "Enter Phone Number Of House"
End Sub

Private Sub PhysicalWelfare_LostFocus()
On Error Resume Next
If PhysicalWelfare.Text = "" Then
PhysicalWelfare.Text = 0
End If
End Sub

Private Sub rolnumber_GotFocus()
Merlin "Enter Roll Number Of The Student Here"
End Sub

Private Sub SaveBtn_Click()
On Error GoTo label

If saveflagfee = True Then

strcf = "INSERT INTO FeesInformation "
strcf = strcf & "(Reciept_Number, Admission_Number, Student_Name, Course_Name, Course_Year, Reciept_Date, Admission_Fee, Coaching_Fee, Library_Fee, Lab_Fee, Special_Fee, Development_Fund, Fine, Migration_Fee, Enrolment_Fee, Physical_Welfare_Fee, Computer_Fee, Caution_Deposit, Endowment, ICard, Other_Fees, Total_Amount, Total_Fee, Remaining_Fee)"
strcf = strcf & "VALUES(" & CDbl(RecieptNumber.Text) & ", "
strcf = strcf & CDbl(AdmissionNumber.Text) & ", "
strcf = strcf & "'" & Trim$(StudentName.Text) & "', "
strcf = strcf & "'" & Trim$(cmbcourse.Text) & "', "
strcf = strcf & "'" & Trim$(cmbyear.Text) & "', "
strcf = strcf & "'" & RecieptDate.Value & "', "
strcf = strcf & CDbl(AdmissionFee.Text) & ", "
strcf = strcf & CDbl(CoachingFee.Text) & ", "
strcf = strcf & CDbl(LibraryFee.Text) & ", "
strcf = strcf & CDbl(LabFee.Text) & ", "
strcf = strcf & CDbl(SpecialFee.Text) & ", "
strcf = strcf & CDbl(DevelopmentFund.Text) & ", "
strcf = strcf & CDbl(Fine.Text) & ", "
strcf = strcf & CDbl(MigrationFee.Text) & ", "
strcf = strcf & CDbl(EnrolmentFee.Text) & ", "
strcf = strcf & CDbl(PhysicalWelfare.Text) & ", "
strcf = strcf & CDbl(ComputerFee.Text) & ", "
strcf = strcf & CDbl(CautionDeposit.Text) & ", "
strcf = strcf & CDbl(Endowment.Text) & ", "
strcf = strcf & CDbl(ICard.Text) & ", "
strcf = strcf & CDbl(OtherFee.Text) & ", "
strcf = strcf & CDbl(TotalAmt.Text) & ", "
strcf = strcf & CDbl(TotalFee.Text) & ", "
strcf = strcf & CDbl(Remaining.Text) & ")"
feescon.Execute strcf

Else

strcf = "UPDATE FeesInformation SET "
strcf = strcf & "Reciept_Number =" & CDbl(RecieptNumber.Text) & ", "
strcf = strcf & "Admission_Number =" & CDbl(AdmissionNumber.Text) & ", "
strcf = strcf & "Student_Name ='" & Trim$(StudentName.Text) & "', "
strcf = strcf & "Course_Name ='" & Trim$(cmbcourse.Text) & "', "
strcf = strcf & "Course_Year ='" & Trim$(cmbyear.Text) & "', "
strcf = strcf & "Reciept_Date ='" & RecieptDate.Value & "', "
strcf = strcf & "Admission_Fee =" & CDbl(AdmissionFee.Text) & ", "
strcf = strcf & "Coaching_Fee =" & CDbl(CoachingFee.Text) & ", "
strcf = strcf & "Library_Fee =" & CDbl(LibraryFee.Text) & ", "
strcf = strcf & "Lab_Fee =" & CDbl(LabFee.Text) & ", "
strcf = strcf & "Special_Fee =" & CDbl(SpecialFee.Text) & ", "
strcf = strcf & "Development_Fund =" & CDbl(DevelopmentFund.Text) & ", "
strcf = strcf & "Fine =" & CDbl(Fine.Text) & ", "
strcf = strcf & "Migration_Fee =" & CDbl(MigrationFee.Text) & ", "
strcf = strcf & "Enrolment_Fee =" & CDbl(EnrolmentFee.Text) & ", "
strcf = strcf & "Physical_Welfare_Fee =" & CDbl(PhysicalWelfare.Text) & ", "
strcf = strcf & "Computer_Fee =" & CDbl(ComputerFee.Text) & ", "
strcf = strcf & "Caution_Deposit =" & CDbl(CautionDeposit.Text) & ", "
strcf = strcf & "Endowment =" & CDbl(Endowment.Text) & ", "
strcf = strcf & "ICard =" & CDbl(ICard.Text) & ", "
strcf = strcf & "Other_Fees =" & CDbl(OtherFee.Text) & ", "
strcf = strcf & "Total_Amount =" & CDbl(TotalAmt.Text) & ", "
strcf = strcf & "Total_Fee =" & CDbl(TotalFee.Text) & ", "
strcf = strcf & "Remaining_Fee =" & CDbl(Remaining.Text) & ""
strcf = strcf & " WHERE Reciept_Number=" & CDbl(RecieptNumber.Text)
feescon.Execute strcf
AddNew.Enabled = True

End If

feesrec.Requery
feesrec.Movelast

Call lockbtnfees(True)
Call showdatafee
Call checkbtnfee
Call lockbtnfee(False)
Call LockTxtFee(True)
XPButton6.Enabled = False
Merlin "Save Edited Or New Record"

MsgBox "Record Saved Successfully", vbInformation, "Save Info"

Exit Sub
label:

If Err.Number = -2147217900 Then
MsgBox strcf
MsgBox ("Reciept Number Already Exist, Please Enter Another Number"), vbCritical, "Admission Number Exist"
Else
MsgBox Err.Description, vbInformation, "Error Occured"
End If
End Sub

Private Sub Search_Click()
On Error Resume Next
Me.Hide
Load Frm_SearchFeeEntry
Frm_SearchFeeEntry.Show
Frm_SearchFeeEntry.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_SearchFeeEntry.Picture = MainMenu.ACPRibbon1.LoadBackground
End Sub

Private Sub SpecialFee_LostFocus()
On Error Resume Next
If SpecialFee.Text = "" Then
SpecialFee.Text = 0
End If
End Sub

Private Sub studetailtab_TabSwitch(ByVal iLastActiveTab As Integer)
On Error GoTo label

If studetailtab.ActiveTab = 1 And adnumber.Locked = False Then
Call locktxtfamily(False)
btnsavedetail.Enabled = True
Else
Call locktxtfamily(True)
btnsavedetail.Enabled = False
End If

If studetailtab.ActiveTab = 1 Then
Merlin "This Is Where Student Personal Information Is Entered"
ElseIf studetailtab.ActiveTab = 2 Then
Merlin "Enter Student Fee Detail Here, You Can Do Fee Entry Detaily From Here"
End If

Exit Sub
label:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Function checkfamilydata() As Boolean
Dim statfamily As Boolean
statfamily = False

If fname.Text = "" Then
MsgBox "Enter Father's Name", vbInformation, "Deficient Data"
ElseIf cmbfoccupation.Text = "" Then
MsgBox "Enter Father's Occupation", vbInformation, "Deficient Data"
ElseIf mname.Text = "" Then
MsgBox "Enter Mother's Name", vbInformation, "Deficient Data"
ElseIf moccupation.Text = "" Then
MsgBox "Enter Mother's Occupation", vbInformation, "Deficient Data"
Else
statfamily = True
End If

checkfamilydata = statfamily
End Function

Private Function LockTxtFee(lockt As Boolean)
RecieptNumber.Locked = lockt
AdmissionNumber.Locked = lockt
StudentName.Locked = lockt
cmbcourse.Locked = lockt
TotalFee.Locked = lockt
TotalAmt.Locked = lockt
Remaining.Locked = lockt
AdmissionFee.Locked = lockt
CoachingFee.Locked = lockt
LibraryFee.Locked = lockt
LabFee.Locked = lockt
SpecialFee.Locked = lockt
DevelopmentFund.Locked = lockt
Fine.Locked = lockt
MigrationFee.Locked = lockt
EnrolmentFee.Locked = lockt
PhysicalWelfare.Locked = lockt
ComputerFee.Locked = lockt
CautionDeposit.Locked = lockt
Endowment.Locked = lockt
OtherFee.Locked = lockt
ICard.Locked = lockt
cmbyear.Locked = lockt
End Function

Private Function lockbtnfee(st As Boolean)
SaveBtn.Enabled = st
RecieptDate.Enabled = st
CancelBtn.Enabled = st
End Function

Private Function checkbtnfee()
If feesrec.RecordCount = 0 Then
edit.Enabled = False
delete.Enabled = False
Search.Enabled = False
Movefirst.Enabled = False
Moveprevious.Enabled = False
Movenext.Enabled = False
Movelast.Enabled = False
Else
edit.Enabled = True
delete.Enabled = True
Search.Enabled = True
Movefirst.Enabled = True
Moveprevious.Enabled = True
Movenext.Enabled = True
Movelast.Enabled = True
End If
End Function

Private Function clearfee()
RecieptNumber.Text = ""
AdmissionNumber.Text = ""
StudentName.Text = ""
cmbcourse.Text = ""
cmbyear.Text = ""
TotalFee.Text = ""
TotalAmt.Text = ""
AdmissionFee.Text = ""
Remaining.Text = ""
CoachingFee.Text = ""
LibraryFee.Text = ""
LabFee.Text = ""
SpecialFee.Text = ""
DevelopmentFund.Text = ""
Fine.Text = ""
MigrationFee.Text = ""
EnrolmentFee.Text = ""
PhysicalWelfare.Text = ""
ComputerFee.Text = ""
CautionDeposit.Text = ""
Endowment.Text = ""
ICard.Text = ""
OtherFee.Text = ""
End Function

Private Function showdatafee()
If feesrec.EOF = False And feesrec.BOF = False Then
    RecieptNumber.Text = feesrec.Fields(0)
    AdmissionNumber.Text = feesrec.Fields(1)
    StudentName.Text = feesrec.Fields(2)
    cmbcourse.Text = feesrec.Fields(3)
    cmbyear.Text = feesrec.Fields(4)
    RecieptDate.Value = feesrec.Fields(5)
    AdmissionFee.Text = feesrec.Fields(6)
    CoachingFee.Text = feesrec.Fields(7)
    LibraryFee.Text = feesrec.Fields(8)
    LabFee.Text = feesrec.Fields(9)
    SpecialFee.Text = feesrec.Fields(10)
    DevelopmentFund.Text = feesrec.Fields(11)
    Fine.Text = feesrec.Fields(12)
    MigrationFee.Text = feesrec.Fields(13)
    EnrolmentFee.Text = feesrec.Fields(14)
    PhysicalWelfare.Text = feesrec.Fields(15)
    ComputerFee.Text = feesrec.Fields(16)
    CautionDeposit.Text = feesrec.Fields(17)
    Endowment.Text = feesrec.Fields(18)
    ICard.Text = feesrec.Fields(19)
    OtherFee.Text = feesrec.Fields(20)
    TotalAmt.Text = feesrec.Fields(21)
    TotalFee.Text = feesrec.Fields(22)
    Remaining.Text = feesrec.Fields(23)
End If
End Function

Private Sub stuname_GotFocus()
Merlin "Enter Name Of The Student Here", "Read"
End Sub

Private Sub temaddress_GotFocus()
Merlin "Enter Student Temporary Address Here"
End Sub

Private Sub TotalFee_LostFocus()
On Error GoTo label
Remaining.Text = Val(TotalFee.Text) - Val(TotalAmt.Text)
Exit Sub
label:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Function disbtnstu(va As Boolean)
btnedit.Enabled = va
btndelete.Enabled = va
btnmovefirst.Enabled = va
btnmoveprevious.Enabled = va
btnmovenext.Enabled = va
btnmovelast.Enabled = va
End Function

Private Function lockbtnfees(valueb As Boolean)
edit.Enabled = valueb
delete.Enabled = valueb
Search.Enabled = valueb
Movefirst.Enabled = valueb
Moveprevious.Enabled = valueb
Movenext.Enabled = valueb
Movelast.Enabled = valueb
End Function

Private Sub XPButton6_Click()
On Error Resume Next
Load Frm_SearchStudentInfo
Frm_SearchStudentInfo.Show
Frm_SearchStudentInfo.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_SearchStudentInfo.Picture = MainMenu.ACPRibbon1.LoadBackground
Merlin "Search For Student Detail"
End Sub

Private Sub yearclass_GotFocus()
Merlin "Select Course Year From Here"
End Sub
