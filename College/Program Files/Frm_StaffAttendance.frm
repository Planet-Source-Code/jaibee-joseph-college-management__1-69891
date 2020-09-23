VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Begin VB.Form Frm_StaffAttendance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mark Staff Attendance"
   ClientHeight    =   5655
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9600
   Icon            =   "Frm_StaffAttendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2520
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=College"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "College"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "StaffAttendanceInformation"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frm_StaffAttendance.frx":076A
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enter New Staff Attendance Entry"
      Top             =   720
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16508095
      HeadLines       =   1
      RowHeight       =   18
      TabAction       =   1
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   "Enter New Entry Here"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Serial_Number"
         Caption         =   "Serial Number"
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
         DataField       =   "Staff_Name"
         Caption         =   "Staff Name"
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
         DataField       =   "Department"
         Caption         =   "Department"
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
         DataField       =   "Atn_Date"
         Caption         =   "Attendance Date"
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
         DataField       =   "Attendance_Status"
         Caption         =   "Attendance Status"
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
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   9120
      Picture         =   "Frm_StaffAttendance.frx":077F
      ToolTipText     =   "Application Help"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mark Staff Daily Attendance Here"
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
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_StaffAttendance.frx":0EE9
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "Frm_StaffAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intHeight As Integer
Dim intCol As Integer
Dim intRow As Integer

Private Sub Form_Load()
' Events That Should Happen When Form Is Loaded
On Error Resume Next

List1.AddItem ""
List2.AddItem ""

staffstnen.Movefirst
Do While Not staffstnen.BOF And Not staffstnen.EOF
   List1.AddItem staffstnen(1).Value
   staffstnen.Movenext
Loop

departglo.Movefirst
Do While Not departglo.BOF And Not departglo.EOF
   List2.AddItem departglo(1).Value
   departglo.Movenext
Loop

Me.Top = 50
Me.Left = 50

DataGrid1.Columns(1).Button = True
DataGrid1.Columns(2).Button = True
Merlin "Staff Daily Attendance Is Entered Here"
End Sub
Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)
' Code To Show List Boxses When Button In Grid Is Clicked
Dim strItem As String
On Error Resume Next
With DataGrid1
strItem = .Text
' Set height, move, select item, make visible, and
' Give focus to list box
Select Case ColIndex
  Case 1
    List1.Height = (.Height / .RowHeight - (intRow - 1)) * .RowHeight
    List1.Move .Left + .Columns(1).Left, _
     .Top + .RowTop(.Row) + .RowHeight, _
     .Columns(1).Width
    If Len(strItem) Then
       List1 = strItem
    Else
       List1.ListIndex = 0
    End If
    List1.Visible = True
    List1.SetFocus
  Case 2
    If intRow > 4 Then ' Place above cell
       List2.Height = (intRow + 1) * .RowHeight
       List2.Move .Left + .Columns(2).Left, _
        .Top + .RowHeight + (intRow * 1.4), _
        .Columns(2).Width
    Else ' Place below cell
       List2.Height = (.Height / .RowHeight - (intRow + 1)) * .RowHeight
       List2.Move .Left + .Columns(2).Left, _
        .Top + .RowTop(.Row) + .RowHeight, _
        .Columns(2).Width
    End If
    If Len(strItem) Then
       ' Find match in Listbox
       Dim n As Integer
       For n = 0 To List2.ListCount - 1
           If strItem = List2.List(n) Then
              ' List2.Selected(n) = True
              List2.ListIndex = n
              Exit For
           End If
       Next
    Else
       List2.ListIndex = 0
    End If
    List2.Visible = True
    List2.SetFocus
End Select
End With
End Sub
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
' Code Executed When Col Or Row Is Changed
intCol = DataGrid1.Col
intRow = DataGrid1.Row
' Label1.Caption = DataGrid1.Row
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
End If
End Sub
Private Sub DataGrid1_Scroll(Cancel As Integer)
On Error Resume Next
' Code To Be Executed When Datagrid Is Scrolled
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
End If
End Sub

Private Sub Form_Click()
On Error Resume Next
' When Clicked On The Form
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
End If
End Sub

Private Sub Image2_Click()
On Error Resume Next
Call showhelpfile
End Sub

Private Sub List1_Click()
' When List Box Is Clicked
On Error Resume Next
DataGrid1.Text = List1.Text
List1.Visible = False
End Sub

Private Sub List1_LostFocus()
' When List Box Lost Its Focus
On Error Resume Next
List1.Visible = False
End Sub

Private Sub List2_Click()
' When List Box Is Clicked
On Error Resume Next
DataGrid1.Text = List2.Text
List2.Visible = False
End Sub

Private Sub List2_LostFocus()
' When List Box Lost Its Focus
On Error Resume Next
List2.Visible = False
End Sub


