VERSION 5.00
Begin VB.Form Frm_Splash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4320
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Frm_Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      TabIndex        =   4
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
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
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   3015
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
      Left            =   4440
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Memory Status"
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
      Left            =   4440
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   0
      Picture         =   "Frm_Splash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "Frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error Resume Next
Set info = New COSInfo
Set wininfo = New GetWindowsInformation
Label1.Caption = "Windows Version :" & " " & wininfo.getwinOS
Label2.Caption = "Windows User :" & " " & wininfo.UserName
Label3.Caption = "Computer Name :" & " " & info.ComputerName
Call AlwaysOnTop(Me, True)

If keyvalc.checkkeyims = False Then
Label4.Caption = "Unregistered User"
Else
Label4.Caption = "Registered User"
End If

Label5.Caption = "Exp. Date" & ":" & " " & GetSetting("VRJ Soft", "Exp Date", "RegExp", dateexp)
End Sub
