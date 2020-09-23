VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#2.0#0"; "Candy.ocx"
Begin VB.Form Frm_ServerManager 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Server and Database Option"
   ClientHeight    =   1575
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4500
   Icon            =   "Frm_ServerManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
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
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "Enter Server Name Here"
      Top             =   240
      Width           =   2415
   End
   Begin Candy.CandyButton btnCancel 
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Unload Form"
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
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
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Click Me To Save Data"
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
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
   Begin VB.TextBox Text1 
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
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "Enter Database Name Here"
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Database Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_ServerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub btnOK_Click()
On Error Resume Next
If Text2.Text = "" Or Text1.Text = "" Then
MsgBox "Empty Server Or Database Name", vbCritical, "Invalid Data"
Exit Sub
Else
namesqlserver = "Server=" & Text2.Text & ";"
namesqldatabase = "DataBase=" & Text1.Text & ";"

SaveSetting App.CompanyName, "ServerSQLName", "ServerName", namesqlserver
SaveSetting App.CompanyName, "ServerDataBaseName", "DataBaseName", namesqldatabase

MsgBox "Please Restart Application For Settings To Be Changed", vbInformation, "Server Settings"
Unload Me
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then btnOK_Click
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then btnOK_Click
End Sub

