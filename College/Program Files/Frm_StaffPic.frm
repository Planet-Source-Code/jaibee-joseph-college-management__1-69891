VERSION 5.00
Object = "{99974422-393D-11D4-B1E7-00104C10C50F}#2.0#0"; "AbstractThumbPage.ocx"
Begin VB.Form Frm_StaffPic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Staff Picture Preview"
   ClientHeight    =   7590
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_StaffPic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin AbstractThumbPage.AbsThumbPage PicPreview 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Existing Staff Pictures"
      Top             =   600
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12091
      Caption         =   "Pictures Existing"
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   8760
      Picture         =   "Frm_StaffPic.frx":076A
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_StaffPic.frx":0ED4
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Of The Existing Staff Pictures From The Staff Image Folder."
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "Frm_StaffPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
' Events That Should Happen When Form Is Loaded
' Add Pictures Present To The Control
On Error Resume Next

Me.Top = 50
Me.Left = 50

Set fldr = FSO.GetFolder(App.Path & "\StaffImages\")

For Each f In fldr.Files
    PicPreview.AddThumb f.Path
Next
    
'Show 10 thumbs on each page
PicPreview.BuildPages 25
    
'Display the first page
PicPreview.ShowPage 1

Load Frm_StaffPictureFull
Frm_StaffPictureFull.Show

Merlin "This Is Where You Can View Staff Pictures", "Read"
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Unload The Maximiser Form When Closed
On Error Resume Next
Unload Frm_StaffPictureFull
End Sub

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub PicPreview_ThumbClick(ThumbPath As String, ThumbName As String)
' Load Picture To Maximiser Form When Clicked
On Error Resume Next
Frm_StaffPictureFull.Image1.Picture = LoadPicture(ThumbPath)
End Sub

