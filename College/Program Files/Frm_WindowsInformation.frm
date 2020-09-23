VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Object = "{5E8FF3F9-2372-4C96-A258-479E142BF3EF}#1.0#0"; "XP_ProBar.ocx"
Begin VB.Form Frm_WindowsInformation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows Information"
   ClientHeight    =   5145
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8640
   Icon            =   "Frm_WindowsInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7646
      Caption         =   "OS Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XP_ProBar.UserControl1 membar 
         Height          =   255
         Left            =   7320
         TabIndex        =   18
         Top             =   3840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   16576
         Scrolling       =   9
         ShowText        =   -1  'True
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   7320
         Top             =   1560
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Label16"
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
         TabIndex        =   16
         Top             =   3840
         Width           =   3495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
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
         TabIndex        =   15
         Top             =   3360
         Width           =   3495
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
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
         TabIndex        =   14
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
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
         TabIndex        =   13
         Top             =   2400
         Width           =   3495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Label12"
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
         TabIndex        =   12
         Top             =   1920
         Width           =   3495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
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
         TabIndex        =   11
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Label10"
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
         TabIndex        =   10
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
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
         TabIndex        =   9
         Top             =   480
         Width           =   3495
      End
      Begin VB.Line Line1 
         X1              =   4080
         X2              =   4080
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
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
         TabIndex        =   8
         Top             =   3840
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
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
         TabIndex        =   7
         Top             =   3360
         Width           =   3495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
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
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
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
         Top             =   2400
         Width           =   3495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
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
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   8160
      Picture         =   "Frm_WindowsInformation.frx":08CA
      ToolTipText     =   "Application Help"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Microsoft Windows Information"
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
      Left            =   840
      TabIndex        =   17
      Top             =   240
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Frm_WindowsInformation.frx":1034
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Frm_WindowsInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
' Show Windows Information When Form Loads
On Error GoTo message

Me.Top = 50
Me.Left = 50

Merlin "You Can View All The Information About Your OS Here"

Set wininfo = New GetWindowsInformation
Set info = New COSInfo
Set cmpinfo = New CComputerInfo
Set keyboardin = New CKeyboardInfo
Set memsta = New CMemoryStatus

Label1.Caption = "Operating System :" & " " & wininfo.getwinOS
Label2.Caption = "User :" & " " & wininfo.UserName
Label3.Caption = "Computer Name :" & " " & info.ComputerName
Label4.Caption = "Windows Build :" & " " & info.OSBuild
Label5.Caption = "Language :" & " " & info.Language
Label6.Caption = "System Directory :" & " " & info.SystemDirectory
Label7.Caption = "Windows Directory :" & " " & info.WindowsDirectory
Label8.Caption = "Company Name :" & " " & info.DefaultCompanyName
Label9.Caption = "Hardware Profile :" & " " & info.CurrentHardwareProfile
Label10.Caption = "Processor Type :" & " " & cmpinfo.ProcessorType
Label11.Caption = "Keyboard Type :" & " " & keyboardin.KeyboardType
Label13.Caption = "Total Page File :" & " " & memsta.TotalPageFile
Label14.Caption = "Total Memory :" & " " & memsta.TotalPhysical
Label15.Caption = "Total Virtual :" & " " & memsta.TotalVirtual

Exit Sub
message:
MsgBox Err.Description, vbCritical, "Error Occured"
End Sub

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub Timer1_Timer()
' Show Memory Status
On Error Resume Next
Label16.Caption = "Available Memory :" & " " & (CDbl(memsta.AvailablePhysical)) & " " & "KB"
Label12.Caption = "Memory Used :" & " " & memsta.MemoryLoad & "%"
membar.Value = 100 - memsta.MemoryLoad
End Sub

