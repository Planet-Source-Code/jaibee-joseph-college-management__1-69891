VERSION 5.00
Begin VB.Form Frm_StudentPictureFull 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maximised Picture"
   ClientHeight    =   2310
   ClientLeft      =   9600
   ClientTop       =   330
   ClientWidth     =   2055
   ControlBox      =   0   'False
   Icon            =   "Frm_StudentPictureFull.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Frm_StudentPictureFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
' Events That Should Happen When Form Is Loaded
' Arrange Form Position
On Error Resume Next
Frm_StudentPictureFull.Left = Frm_StudentPic.Width + 20
Frm_StudentPictureFull.Top = Frm_StudentPic.Top
End Sub
