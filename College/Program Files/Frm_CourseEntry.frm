VERSION 5.00
Object = "{79817FF7-38F3-446A-8956-C9E957F74576}#1.0#0"; "Candy.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form Frm_CourseEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Course Entry"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "Frm_CourseEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin Candy.CandyButton btncancel 
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      ToolTipText     =   "Cancel Entry"
      Top             =   6360
      Width           =   2175
      _ExtentX        =   3836
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
   Begin Candy.CandyButton delete 
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      ToolTipText     =   "Delete Record"
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
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
   Begin Candy.CandyButton edit 
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      ToolTipText     =   "Edit Current Record"
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
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
   Begin Candy.CandyButton save 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      ToolTipText     =   "Save Record"
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
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
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Add New Record"
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
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
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9128
      Caption         =   "Enter Course Information Here"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtcourseid 
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
         Left            =   2040
         TabIndex        =   1
         ToolTipText     =   "Enter Course ID Here"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtcoursename 
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
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Enter Course Name Here"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtuniversity 
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
         Left            =   2040
         TabIndex        =   5
         ToolTipText     =   "Enter University Name Here"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.ComboBox cmbcoursesystem 
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
         ItemData        =   "Frm_CourseEntry.frx":076A
         Left            =   2040
         List            =   "Frm_CourseEntry.frx":0774
         TabIndex        =   3
         ToolTipText     =   "Select Course System From Here"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txttotalseats 
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
         Left            =   2040
         TabIndex        =   6
         ToolTipText     =   "Total Number Of Seats Available"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txteligibility 
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
         Left            =   2040
         TabIndex        =   7
         ToolTipText     =   "Eligibility Needed For Admission"
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox txtcombination 
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
         Left            =   2040
         TabIndex        =   8
         ToolTipText     =   "Combination Of Subjects"
         Top             =   3960
         Width           =   2535
      End
      Begin VB.TextBox txtsubjects 
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
         Left            =   2040
         TabIndex        =   9
         ToolTipText     =   "Subjects In The Course"
         Top             =   4440
         Width           =   2535
      End
      Begin VB.ComboBox cmbcourseduration 
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
         ItemData        =   "Frm_CourseEntry.frx":078A
         Left            =   2040
         List            =   "Frm_CourseEntry.frx":07AF
         TabIndex        =   4
         ToolTipText     =   "Select Course Duration From Here"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Course ID"
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
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Course System"
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
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Seats Available"
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
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "University"
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
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Course Duration"
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
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Eligibility"
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
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Combination"
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
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Subjects"
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
         TabIndex        =   19
         Top             =   4440
         Width           =   1575
      End
   End
   Begin Candy.CandyButton Movenext 
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      ToolTipText     =   "Move To Next Record"
      Top             =   6360
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
      Left            =   4320
      TabIndex        =   18
      ToolTipText     =   "Move To Last Record"
      Top             =   6360
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
      Left            =   720
      TabIndex        =   15
      ToolTipText     =   "Move To Previous Record"
      Top             =   6360
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
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Move To First Record"
      Top             =   6360
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
   Begin VB.Image Image2 
      Height          =   360
      Left            =   4320
      Picture         =   "Frm_CourseEntry.frx":082B
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Details Of The Course Offred Here"
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
      TabIndex        =   27
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_CourseEntry.frx":0F95
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Frm_CourseEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variables Needed
Dim strc As String
Dim saveflagc As Boolean

Private Sub AddNew_Click()
' Add New Record
On Error GoTo message

saveflagc = True
Call LockTxtCourse(False)
Call disablebtn(False)
Call lockbtncourse(True)
Call cleardata
txtcourseid.SetFocus
Merlin "Add New Course"

Exit Sub
message:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub btnCancel_Click()
' Cancel Entry
On Error Resume Next

Call cleardata
Call LockTxtCourse(True)
Call disablebtn(False)
Call lockbtncourse(False)
Call checkbtncourse
Merlin "Cancel Uapdation"

If courserec.BOF And courserec.EOF Then
MsgBox "No Existing Record, Insert New Record", vbInformation, "No Record"
Else
courserec.Movefirst
Call showdatacourse
End If
End Sub

Private Sub cmbcourseduration_GotFocus()
Merlin "Enter Course Duration Here"
End Sub

Private Sub cmbcoursesystem_GotFocus()
Merlin "Select One Course System From Here"
End Sub

Private Sub delete_Click()
' Delete Record
On Error GoTo message

Merlin "Delete Current Record"
If MsgBox("Execution Of Command Will Delete Current Datarecord" & vbCrLf & "Are You Sure You Wan't To Delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then

strc = "DELETE FROM CourseInformation WHERE "
strc = strc & "Course_ID = "
strc = strc & CDbl(txtcourseid.Text)
coursecon.Execute strc
courserec.Requery

MsgBox "Record Deleted Sucessfully.", vbInformation, "Delete Record"

If courserec.BOF And courserec.EOF Then
Call cleardata
MsgBox ("The Previous Record Was Last Record."), vbInformation, "Last Record"
Call checkbtncourse
Else
courserec.Movenext
If courserec.EOF Then
courserec.Movelast
End If
Call showdatacourse
End If

End If

Exit Sub
message:
MsgBox "No Existing Record, Insert New Record", vbInformation, "Error Occured"

End Sub

Private Sub Edit_Click()
' Edit Record
On Error GoTo mesa

saveflagc = False
Call LockTxtCourse(False)
Call disablebtn(False)
Call lockbtncourse(True)
Merlin "Edit Current Record"

Exit Sub
mesa:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub Form_Load()
' When Form Is Loaded
On Error GoTo message

Call LockTxtCourse(True)
Call lockbtncourse(False)
Call checkbtncourse
Call showdatacourse
Merlin "This Is Where Information About The Course Is Entered", "DoMagic1"

Me.Top = 50
Me.Left = 50
Exit Sub
message:
MsgBox Err.Description, vbInformation, "Error In Connection"
End Sub

' Function To Lock And Unlock Text Box
Private Function LockTxtCourse(lockstat As Boolean)
txtcoursename.Locked = lockstat
cmbcoursesystem.Locked = lockstat
cmbcourseduration.Locked = lockstat
txtuniversity.Locked = lockstat
txttotalseats.Locked = lockstat
txteligibility.Locked = lockstat
txtcombination.Locked = lockstat
txtsubjects.Locked = lockstat
txtcourseid.Locked = lockstat
End Function

' Function To Lock And Unlock Button
Private Function lockbtncourse(lockst As Boolean)
save.Enabled = lockst
btnCancel.Enabled = lockst
End Function

' Check Whether Buttons Should Be Enabled Or Not
Private Function checkbtncourse()
If courserec.RecordCount = 0 Then
Call disablebtn(False)
Else
Call disablebtn(True)
End If
End Function

' Function To Disable Buttons
Private Function disablebtn(statb As Boolean)
edit.Enabled = statb
delete.Enabled = statb
Movefirst.Enabled = statb
Moveprevious.Enabled = statb
Movenext.Enabled = statb
Movelast.Enabled = statb
End Function

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub Movefirst_Click()
' Move To First Record
On Error GoTo GoFirstError

courserec.Movefirst
' Show the current data record
Call showdatacourse
 
Exit Sub

GoFirstError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub Movelast_Click()
' Move To Last Record
On Error GoTo GoLastError

courserec.Movelast
' show the current data record
Call showdatacourse
Exit Sub

GoLastError:
MsgBox "No Existing Records, Insert New Record", vbInformation, "No Records"
End Sub

Private Sub Movenext_Click()
' Move To Next Record
On Error GoTo GoNextError

If Not courserec.EOF Then courserec.Movenext
If courserec.EOF And courserec.RecordCount > 0 Then
' Moved off the end so go back
courserec.Movelast
End If
' Show the current data record
Call showdatacourse
  
Exit Sub
GoNextError:
MsgBox Err.Description, vbInformation, "Error Occured"
End Sub

Private Sub Moveprevious_Click()
' Move To Previous Record
On Error GoTo GoPrevError
  
If Not courserec.BOF Then courserec.Moveprevious
If courserec.BOF And courserec.RecordCount > 0 Then
    
' Moved off the end so go back
courserec.Moveprevious
 
End If
' Show the current data record
Call showdatacourse
Exit Sub

GoPrevError:
If Err.Number = 3021 Then
MsgBox ("This Is First Record."), vbInformation, "First Record"
courserec.Movenext
ElseIf Err.Number <> 0 Then
MsgBox Err.Description, vbInformation, "Error Occured"
End If
End Sub

Private Sub save_Click()
' Save Data To Database
On Error GoTo mes

Merlin "Save Course Information Here"
If saveflagc = True Then
strc = "INSERT INTO CourseInformation"
strc = strc & "(Course_ID, Course_Name, Course_System, Course_Duration, University, Seats_Available, Eligibility, Combination, Subjects) "
strc = strc & "VALUES(" & CDbl(txtcourseid.Text) & ","
strc = strc & "'" & Trim$(txtcoursename.Text) & "', "
strc = strc & "'" & Trim$(cmbcoursesystem.Text) & "', "
strc = strc & "'" & Trim$(cmbcourseduration.Text) & "', "
strc = strc & "'" & Trim$(txtuniversity.Text) & "', "
strc = strc & CDbl(txttotalseats.Text) & ", "
strc = strc & "'" & Trim$(txteligibility.Text) & "', "
strc = strc & "'" & Trim$(txtcombination.Text) & "', "
strc = strc & "'" & Trim$(txtsubjects.Text) & "')"
coursecon.Execute strc
Else
strc = "UPDATE CourseInformation SET "
strc = strc & "Course_ID=" & CDbl(txtcourseid.Text) & ","
strc = strc & "Course_Name='" & Trim$(txtcoursename.Text) & "',"
strc = strc & "Course_System='" & Trim$(cmbcoursesystem.Text) & "',"
strc = strc & "Course_Duration='" & Trim$(cmbcourseduration.Text) & "',"
strc = strc & "University='" & Trim$(txtuniversity.Text) & "',"
strc = strc & "Seats_Available=" & CDbl(txttotalseats.Text) & ","
strc = strc & "Eligibility='" & Trim$(txteligibility.Text) & "',"
strc = strc & "Combination='" & Trim$(txtcombination.Text) & "',"
strc = strc & "Subjects='" & Trim$(txtsubjects.Text) & "'"
strc = strc & " WHERE Course_ID=" & CDbl(txtcourseid.Text)
coursecon.Execute strc
End If

MsgBox "Record Has Been Successfully Saved", vbInformation, "Saved"
courserec.Requery
courserec.Movelast
Call showdatacourse
Call lockbtncourse(False)
Call checkbtncourse

Exit Sub
mes:
MsgBox Err.Description, vbInformation, "Error Occured"
MsgBox strc
End Sub

' Function To Display Data
Private Function showdatacourse()
If courserec.EOF = False And courserec.BOF = False Then
          txtcourseid.Text = courserec.Fields(0)
          txtcoursename.Text = courserec.Fields(1)
          cmbcoursesystem.Text = courserec.Fields(2)
          cmbcourseduration.Text = courserec.Fields(3)
          txtuniversity.Text = courserec.Fields(4)
          txttotalseats.Text = courserec.Fields(5)
          txteligibility.Text = courserec.Fields(6)
          txtcombination.Text = courserec.Fields(7)
          txtsubjects.Text = courserec.Fields(8)
End If
End Function

' Clear All Text Boxes
Private Function cleardata()
txtcourseid.Text = ""
txtcoursename.Text = ""
cmbcoursesystem.Text = ""
cmbcourseduration.Text = ""
txtuniversity.Text = ""
txttotalseats.Text = ""
txteligibility.Text = ""
txtcombination.Text = ""
txtsubjects.Text = ""
End Function

Private Sub txtcombination_GotFocus()
Merlin "Enter Combination Of Subjects Here"
End Sub

Private Sub txtcourseid_GotFocus()
Merlin "Enter Course ID Here"
End Sub

Private Sub txtcoursename_GotFocus()
Merlin "Enter Course Name Here"
End Sub

Private Sub txteligibility_GotFocus()
Merlin "Enter Eligibility For The Course Here"
End Sub

Private Sub txtsubjects_GotFocus()
Merlin "Enter Course Subjects Here"
End Sub

Private Sub txttotalseats_GotFocus()
Merlin "Enter Total Seats Available Here"
End Sub

Private Sub txtuniversity_GotFocus()
Merlin "Enter University Name Here"
End Sub
