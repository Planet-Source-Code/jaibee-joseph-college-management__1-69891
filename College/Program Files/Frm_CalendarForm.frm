VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form Frm_CalendarForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMS Calendar"
   ClientHeight    =   3645
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4515
   Icon            =   "Frm_CalendarForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4515
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   3660
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "IMS Calendar"
      Top             =   0
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   6456
      _Version        =   393216
      ForeColor       =   192
      BackColor       =   16508095
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16508095
      StartOfWeek     =   49086465
      TitleBackColor  =   4210816
      CurrentDate     =   39437
   End
End
Attribute VB_Name = "Frm_CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
On Error Resume Next
    Me.Top = 700
    Me.Left = (Screen.Width - Me.Width) / 2
    Merlin "IMS Calendar Control, Today Date Is" & "  " & Date, "Read"
End Sub
