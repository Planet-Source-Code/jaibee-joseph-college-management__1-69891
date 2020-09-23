VERSION 5.00
Begin VB.Form Frm_UserMessage 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Dialog Caption"
   ClientHeight    =   1860
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   1440
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   1920
      Top             =   1320
   End
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   2400
      Top             =   1320
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   360
      Picture         =   "Frm_UserMessage.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   435
   End
End
Attribute VB_Name = "Frm_UserMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
On Error Resume Next
' Set that the form was  loaded once
Loaded = True
' Enable the Timer to hide the form
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.BackColor = MainMenu.ACPRibbon1.BackColor
Me.Picture = MainMenu.ACPRibbon1.LoadBackground
' Get the Picture
' Main Part
' Set The Position of the form
Me.Top = Screen.Height
Me.Left = Screen.Width - Me.Width

' Make the Form Transparent
Call Trans(Frm_UserMessage)

' Check if The form had been loaded , if loaded then unload it
If Loaded Then Unload Me: Loaded = False Else Loaded = True
End Sub

' This sub calculates the total height necessary for the form
' This sub resizes the form according to the message. This should me called before
' The form is shown
Public Sub Resize()
On Error Resume Next
Me.Height = Me.Height + Label2.Height
End Sub


Private Sub Form_Terminate()
On Error Resume Next
Unload Me ' Unload the form
End Sub

' Slide the form into view
Private Sub Timer1_Timer()
On Error Resume Next
' Check if the form has reached its maximum height & if Yes then stop timer
If Me.Top <= Screen.Height - (Me.Height + 250) Then Timer1.Enabled = False

' Else Move it position to 100 pix top
Me.Top = Me.Top - 100
End Sub
 
' Scroll out timer
Private Sub Timer2_Timer()
On Error Resume Next ' Will raise an error if the form is unloaded unexpectedly
' So keep an error trapper

' Check if the form has reached its minimum height & if Yes then stop timer
If Me.Top >= 11000 Then
Timer2.Enabled = False
' Unload the form
Unload Me
Load Frm_TipInformation
Frm_TipInformation.Show
Frm_TipInformation.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_TipInformation.Picture = MainMenu.ACPRibbon1.LoadBackground
Load Frm_Sidebar
Frm_Sidebar.Show

End If

' Else Move it position to 100 pix down
Me.Top = Me.Top + 100
End Sub


' Time out Timer
' Automatically close the box after 20 secs.
Private Sub Timer3_Timer()
On Error Resume Next
' Enable Scroll out timer
Timer2.Enabled = True

' Disable this timer
Timer3.Enabled = False

End Sub
