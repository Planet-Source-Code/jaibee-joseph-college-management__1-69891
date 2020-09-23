VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_UniversityReciept 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New University Fee Reciept"
   ClientHeight    =   8160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7440
   Icon            =   "Frm_UniversityReciept.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   27
      Top             =   5400
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
      Caption         =   "Search Fee Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Candy.CandyButton btnprint 
         Height          =   255
         Left            =   3720
         TabIndex        =   36
         ToolTipText     =   "Print Searched Records"
         Top             =   2280
         Width           =   3255
         _ExtentX        =   5741
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
         Caption         =   "Print"
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
      Begin Candy.CandyButton btnsearch 
         Height          =   255
         Left            =   240
         TabIndex        =   35
         ToolTipText     =   "Search Selected Values"
         Top             =   2280
         Width           =   3255
         _ExtentX        =   5741
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
      Begin VB.ComboBox cmbtype 
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
         ItemData        =   "Frm_UniversityReciept.frx":076A
         Left            =   4680
         List            =   "Frm_UniversityReciept.frx":0777
         TabIndex        =   34
         ToolTipText     =   "Select Search Type From Here"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtval 
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
         Left            =   1320
         TabIndex        =   32
         ToolTipText     =   "Enter Search Value Here"
         Top             =   360
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Frm_UniversityReciept.frx":07A4
         Height          =   1335
         Left            =   240
         TabIndex        =   30
         ToolTipText     =   "Search Results"
         Top             =   840
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2355
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16508095
         HeadLines       =   1
         RowHeight       =   18
         FormatLocked    =   -1  'True
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
         Caption         =   "Search Results"
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "Reciept_Number"
            Caption         =   "Reciept Number"
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
            DataField       =   "Student_Name"
            Caption         =   "Student Name"
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
            DataField       =   "Date_Reciept"
            Caption         =   "Date Reciept"
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
            DataField       =   "Course_Name"
            Caption         =   "Course Name"
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
            DataField       =   "Year"
            Caption         =   "Year"
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
            DataField       =   "Form_Number"
            Caption         =   "Form Number"
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
            DataField       =   "Exam_Form_Fee"
            Caption         =   "Exam Form Fee"
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
            DataField       =   "Examination_Fee"
            Caption         =   "Examination Fee"
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
            DataField       =   "Total_Amount"
            Caption         =   "Total Amount"
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
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Type"
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
         Left            =   3720
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Value"
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
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8281
      Caption         =   "Add New Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Candy.CandyButton btnhelp 
         Height          =   375
         Left            =   4920
         TabIndex        =   15
         ToolTipText     =   "Help For The Form"
         Top             =   2760
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
      Begin VB.ComboBox cmbyear 
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
         ItemData        =   "Frm_UniversityReciept.frx":07B9
         Left            =   1920
         List            =   "Frm_UniversityReciept.frx":07DE
         TabIndex        =   5
         ToolTipText     =   "Select Course Year"
         Top             =   2280
         Width           =   2775
      End
      Begin VB.ComboBox cmbcourse 
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
         TabIndex        =   4
         ToolTipText     =   "Select Course Name"
         Top             =   1800
         Width           =   2775
      End
      Begin Candy.CandyButton btnmovelast 
         Height          =   375
         Left            =   6720
         TabIndex        =   19
         ToolTipText     =   "Move To The Last Record"
         Top             =   4200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
         Height          =   375
         Left            =   6360
         TabIndex        =   18
         ToolTipText     =   "Move To Next Record"
         Top             =   4200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
         Height          =   375
         Left            =   5280
         TabIndex        =   17
         ToolTipText     =   "Move To Previous Record"
         Top             =   4200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         ToolTipText     =   "Move To First Record"
         Top             =   4200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
      Begin Candy.CandyButton btncancel 
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         ToolTipText     =   "Cancel Add New Or Edit"
         Top             =   2280
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
         Left            =   4920
         TabIndex        =   13
         ToolTipText     =   "Delete Current Record"
         Top             =   1800
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
         Left            =   4920
         TabIndex        =   12
         ToolTipText     =   "Edit Current Record"
         Top             =   1320
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
         Left            =   4920
         TabIndex        =   11
         ToolTipText     =   "Save Edited Or New Record"
         Top             =   840
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
      Begin Candy.CandyButton btnadd 
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         ToolTipText     =   "Add New Record"
         Top             =   360
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
      Begin VB.TextBox txttotalamt 
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
         Left            =   1920
         TabIndex        =   9
         ToolTipText     =   "Enter Total Amount Here"
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox txtexaminationfee 
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
         Left            =   1920
         TabIndex        =   8
         ToolTipText     =   "Enter Examination Fee Here"
         Top             =   3720
         Width           =   2775
      End
      Begin VB.TextBox txtexamformfee 
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
         Left            =   1920
         TabIndex        =   7
         ToolTipText     =   "Enter Exam Form Fee Here"
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtformnumber 
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
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "Enter Form Number Here"
         Top             =   2760
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker datepicker 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         ToolTipText     =   "Enter Reciept Date Here"
         Top             =   1320
         Width           =   2775
         _ExtentX        =   4895
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
         Format          =   48889857
         CurrentDate     =   39430
      End
      Begin VB.TextBox txtstuname 
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
         Left            =   1920
         TabIndex        =   2
         ToolTipText     =   "Enter Student Name Here"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtreciept 
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
         Left            =   1920
         TabIndex        =   1
         ToolTipText     =   "Enter Reciept Number Here"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         TabIndex        =   26
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Examination Fee"
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
         TabIndex        =   25
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Form Fee"
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
         TabIndex        =   24
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Form Number"
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
         TabIndex        =   23
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   6960
      Picture         =   "Frm_UniversityReciept.frx":085A
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "University Fee Entry Can Be Done Here"
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
      TabIndex        =   37
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_UniversityReciept.frx":0FC4
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Frm_UniversityReciept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable declarations
Dim saveflag As Boolean
Dim struni As String

Private Function checkbtnuniversity()
' Function to check whether the buttons should be enabled or disabled
If universityrec.RecordCount = 0 Then
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

Private Function disablebtnuniversity(statb As Boolean)
' Function to disable buttons
btnedit.Enabled = statb
btndelete.Enabled = statb
btnmovefirst.Enabled = statb
btnmoveprevious.Enabled = statb
btnmovenext.Enabled = statb
btnmovelast.Enabled = statb
End Function

Private Function lockbtnuniversity(lockst As Boolean)
' Function to lock buttons and DTPicker
btnsave.Enabled = lockst
btnCancel.Enabled = lockst
datepicker.Enabled = lockst
End Function

Private Sub btnadd_Click()
' Code to add new record
On Error GoTo message

saveflag = True
Call locktextbox(False)
Call disablebtnuniversity(False)
Call lockbtnuniversity(True)
Call cleardatauniversity
datepicker.Value = Now
txtreciept.SetFocus
Merlin "Click Me To Add New Receipt"

Exit Sub
message:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnCancel_Click()
' Cancel add new or edit
On Error Resume Next

Call cleardatauniversity
Call locktextbox(True)
Call disablebtnuniversity(False)
Call lockbtnuniversity(False)
Call checkbtnuniversity
Merlin "Cancel Add New Or Edit"

If universityrec.BOF And universityrec.EOF Then
MsgBox "No Existing Record, Insert New Record", vbInformation, "No Record"
Else
universityrec.Movefirst
Call showdatauniversity
End If
End Sub

Private Sub btndelete_Click()
' Delete current record
On Error GoTo message

Merlin "Delete Current Record"

If MsgBox("Execution Of Command Will Delete Current Datarecord" & vbCrLf & "Are You Sure You Wan't To Delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then

struni = "DELETE FROM UniversityFeeInformation WHERE "
struni = struni & "Reciept_Number= "
struni = struni & CDbl(txtreciept.Text)
universitycon.Execute struni
universityrec.Requery

MsgBox "Record Deleted Sucessfully.", vbInformation, "Delete Record"

If universityrec.BOF And universityrec.EOF Then
Call cleardatauniversity
MsgBox ("The Previous Record Was Last Record."), vbInformation, "Last Record"
Call checkbtnuniversity
Else
universityrec.Movenext
If universityrec.EOF Then
universityrec.Movelast
End If
Call showdatauniversity
End If

End If

Exit Sub
message:
MsgBox "No Existing Record, Insert New Record", vbInformation, "Error Occured"
End Sub

Private Sub btnedit_Click()
' Code to edit a record
On Error GoTo mesa

saveflag = False
Call locktextbox(False)
Call disablebtnuniversity(False)
Call lockbtnuniversity(True)
Merlin "Click Me To Edit Current Record"

Exit Sub
mesa:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnhelp_Click()
On Error Resume Next
Call showhelpfile
Merlin "Application Help"
End Sub

Private Sub btnmovefirst_Click()
' Code to move to the first record of the table
On Error GoTo GoFirstError

universityrec.Movefirst
'Show the current data record
Call showdatauniversity
 
Exit Sub

GoFirstError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub btnmovelast_Click()
' Code to move to the last record of the table
On Error GoTo GoLastError

universityrec.Movelast
'Show the current data record
Call showdatauniversity
Exit Sub

GoLastError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub btnmovenext_Click()
' Code to move to the next record of the table
On Error GoTo GoNextError
  
If Not universityrec.EOF Then universityrec.Movenext
If universityrec.EOF And universityrec.RecordCount > 0 Then
' Moved off the end so go back
universityrec.Movelast
End If
' Show the current data record
Call showdatauniversity
  
Exit Sub
GoNextError:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnmoveprevious_Click()
' Code to move to the previous record of the table
On Error GoTo GoPrevError
  
If Not universityrec.BOF Then universityrec.Moveprevious
If universityrec.BOF And universityrec.RecordCount > 0 Then
    
' Moved off the end so go back
universityrec.Moveprevious
 
End If
' Show the current data record
Call showdatauniversity
Exit Sub

GoPrevError:
If Err.Number = 3021 Then
MsgBox ("This Is First Record."), vbInformation, "First Record"
universityrec.Movenext
ElseIf Err.Number <> 0 Then
MsgBox Err.Description, vbInformation, "Error Occured"
End If
End Sub

Private Sub btnprint_Click()
' Code to print the searched results
On Error GoTo message

Set UniversityFeeReport.DataSource = searuniversity
Load UniversityFeeReport
UniversityFeeReport.Show
Merlin "Print Search Value"

Exit Sub
message:
MsgBox "Search Again And Then Print", vbCritical, "Error Occured"
End Sub

Private Sub btnsave_Click()
' Code to save
On Error GoTo message

Merlin "Click Me To Save Record"

If checkalldata = True Then
If saveflag = True Then
struni = "INSERT INTO UniversityFeeInformation"
struni = struni & "(Reciept_Number, Student_Name, Date_Reciept, Course_Name, Year, Form_Number, Exam_Form_Fee, Examination_Fee, Total_Amount) "
struni = struni & "VALUES(" & CDbl(txtreciept.Text) & ", "
struni = struni & "'" & Trim$(txtstuname.Text) & "', "
struni = struni & "'" & datepicker.Value & "', "
struni = struni & "'" & Trim$(cmbcourse.Text) & "', "
struni = struni & "'" & Trim$(cmbyear.Text) & "', "
struni = struni & "'" & Trim$(txtformnumber.Text) & "', "
struni = struni & CDbl(txtexamformfee.Text) & ", "
struni = struni & CDbl(txtexaminationfee.Text) & ", "
struni = struni & CDbl(txttotalamt.Text) & ")"
universitycon.Execute struni
Else
struni = "UPDATE UniversityFeeInformation SET "
struni = struni & "Reciept_Number=" & CDbl(txtreciept.Text) & ","
struni = struni & "Student_Name='" & Trim$(txtstuname.Text) & "',"
struni = struni & "Date_Reciept='" & datepicker.Value & "',"
struni = struni & "Course_Name='" & Trim$(cmbcourse.Text) & "',"
struni = struni & "Year='" & Trim$(cmbyear.Text) & "',"
struni = struni & "Form_Number='" & Trim$(txtformnumber.Text) & "',"
struni = struni & "Exam_Form_Fee=" & CDbl(txtexamformfee.Text) & ","
struni = struni & "Examination_Fee=" & CDbl(txtexaminationfee.Text) & ","
struni = struni & "Total_Amount=" & CDbl(txttotalamt.Text) & ""
struni = struni & " WHERE Reciept_Number=" & CDbl(txtreciept.Text)
universitycon.Execute struni
End If

universityrec.Requery
universityrec.Movelast
Call showdatauniversity
Call lockbtnuniversity(False)
Call checkbtnuniversity

MsgBox "Record Has Been Successfully Saved", vbInformation, "Saved"
End If

Exit Sub
message:
If Err.Number = -2147217900 Then
MsgBox ("Reciept Number Already Exist,Please Enter Another Number"), vbCritical, "Reciept Number Exist"
Else
MsgBox Err.Description, vbCritical, "Error Occured"
End If
End Sub

Private Function locktextbox(statt As Boolean)
' Function to lock and unlock text boxes
txtreciept.Locked = statt
txtstuname.Locked = statt
txtformnumber.Locked = statt
txtexamformfee.Locked = statt
txtexaminationfee.Locked = statt
txttotalamt.Locked = statt
cmbcourse.Locked = statt
cmbyear.Locked = statt
End Function

Private Function cleardatauniversity()
' Function to clear data
txtreciept.Text = ""
txtstuname.Text = ""
txtformnumber.Text = ""
txtexamformfee.Text = ""
txtexaminationfee.Text = ""
txttotalamt.Text = ""
cmbcourse.Text = ""
cmbyear.Text = ""
End Function

Private Function showdatauniversity()
' Function to show data of the table
If universityrec.EOF = False And universityrec.BOF = False Then
txtreciept.Text = universityrec.Fields(0)
txtstuname.Text = universityrec.Fields(1)
datepicker.Value = universityrec.Fields(2)
cmbcourse.Text = universityrec.Fields(3)
cmbyear.Text = universityrec.Fields(4)
txtformnumber.Text = universityrec.Fields(5)
txtexamformfee.Text = universityrec.Fields(6)
txtexaminationfee.Text = universityrec.Fields(7)
txttotalamt.Text = universityrec.Fields(8)
End If
End Function

Private Function checkalldata() As Boolean
' Function to check whether all the data are properly entered
Dim stat As Boolean

stat = False

If txtreciept.Text = "" Or Not IsNumeric(txtreciept.Text) Then
MsgBox "Enter Reciept Number Correctly", vbInformation, "No Reciept Number"
txtreciept.SetFocus
ElseIf txtstuname.Text = "" Then
MsgBox "Enter Student Name", vbInformation, "No Student Name"
txtstuname.SetFocus
ElseIf cmbcourse.Text = "" Then
MsgBox "Enter Course Name", vbInformation, "Empty Field"
cmbcourse.SetFocus
ElseIf cmbyear.Text = "" Then
MsgBox "Enter Course Year", vbInformation, "Empty Field"
cmbyear.SetFocus
ElseIf Not IsNumeric(txtexamformfee.Text) Or txtexamformfee.Text = "" Then
MsgBox "Enter Exam Form Fee Correctly", vbInformation, "Type Mismatch"
txtexamformfee.SetFocus
ElseIf Not IsNumeric(txtexaminationfee.Text) Or txtexaminationfee.Text = "" Then
MsgBox "Enter Examination Fee Correctly", vbInformation, "Type Mismatch"
txtexaminationfee.SetFocus
ElseIf txttotalamt.Text = "" Or Not IsNumeric(txttotalamt.Text) Then
MsgBox "Enter Total Amount Correctly", vbInformation, "Total Amt"
txttotalamt.SetFocus
Else
stat = True
End If

checkalldata = stat
End Function

Private Sub btnsearch_Click()
' Code to search data
On Error GoTo message

Merlin "Search Value By Clicking Me"

Dim searstr As String
If cmbtype.Text = "All Records" And txtval.Text = "" Then
searstr = "Select * from UniversityFeeInformation Order By Reciept_Number"
ElseIf cmbtype = "By Name" And txtval.Text <> "" Then
searstr = "Select * from UniversityFeeInformation where Student_Name like '" & Trim$(txtval.Text) & "%'"
ElseIf cmbtype = "By Reciept Number" And txtval.Text <> "" Then
searstr = "select * from UniversityFeeInformation where Reciept_Number like '" & Trim$(txtval.Text) & "%'"
ElseIf cmbtype.Text = "" And txtval.Text = "" Then
MsgBox "Select Correct Configuration Options", vbCritical, "Error In Options"
Exit Sub
Else
MsgBox "Select Correct Configuration Options", vbCritical, "Error In Options"
Exit Sub
End If

Set searuniversity = New ADODB.Recordset
searuniversity.Open searstr, universitycon, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = searuniversity
DataGrid1.ReBind

Exit Sub
message:
MsgBox Err.Description, vbCritical, "Error Occured"
End Sub

Private Sub cmbcourse_GotFocus()
Merlin "Select Course Name From Here"
End Sub

Private Sub cmbtype_GotFocus()
Merlin "Select Search Type"
End Sub

Private Sub cmbyear_GotFocus()
Merlin "Select Course Year From Here"
End Sub

Private Sub datepicker_GotFocus()
Merlin "Select Date Of Receipt Issue"
End Sub

Private Sub Form_Load()
' Events to be happened as the form loads
On Error GoTo message

unifeeentrycou.Movefirst
Do While Not unifeeentrycou.BOF And Not unifeeentrycou.EOF
   cmbcourse.AddItem unifeeentrycou(1).Value
   unifeeentrycou.Movenext
Loop

Me.Top = 50
Me.Left = 50

Call locktextbox(True)
Call lockbtnuniversity(False)
Call checkbtnuniversity
Call showdatauniversity
Merlin "This Is Where You Enter University Fee Receipt"

Exit Sub
message:
MsgBox Err.Description, vbCritical, "Error Occured"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Ask For Save If Save Button Is Enabled
On Error Resume Next
If btnsave.Enabled = True Then
message = MsgBox("Exit WithOut Saving The Record", vbQuestion + vbYesNo, "Save")
If message = vbYes Then
Unload Me
Else
Cancel = 1
End If
End If
End Sub

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub txtexamformfee_GotFocus()
Merlin "Enter Exam Form Fee"
End Sub

Private Sub txtexaminationfee_GotFocus()
Merlin "Enter Examination Fee"
End Sub

Private Sub txtexaminationfee_LostFocus()
' Add Values To Get Total Amount
On Error Resume Next
txttotalamt.Text = Val(txtexamformfee.Text) + Val(txtexaminationfee.Text)
End Sub

Private Sub txtformnumber_GotFocus()
Merlin "Enter Form Number"
End Sub

Private Sub txtstuname_GotFocus()
Merlin "Enter Student Name Here", "DoMagic2"
End Sub

Private Sub txttotalamt_GotFocus()
Merlin "Total Amount Of Receipt"
End Sub

Private Sub txtval_GotFocus()
Merlin "Enter Search Value Here"
End Sub

Private Sub txtval_KeyPress(KeyAscii As Integer)
' Event to be happened when enter key is pressed
On Error Resume Next

If KeyAscii = 13 Then
btnsearch_Click
End If
End Sub

