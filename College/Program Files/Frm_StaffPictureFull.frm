VERSION 5.00
Begin VB.Form Frm_StaffPictureFull 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maximised Picture"
   ClientHeight    =   2295
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
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
Attribute VB_Name = "Frm_StaffPictureFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
' Events That Should Happen When Form Is Loaded
' Arrange Form Position
On Error Resume Next
Frm_StaffPictureFull.Left = Frm_StaffPic.Width + 20
Frm_StaffPictureFull.Top = Frm_StaffPic.Top
End Sub
