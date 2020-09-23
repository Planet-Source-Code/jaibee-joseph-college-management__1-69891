VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Begin VB.Form Frm_StudentAttendance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mark Student Monthly Attendance"
   ClientHeight    =   5640
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9600
   Icon            =   "Frm_StudentAttendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   2400
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
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
      Height          =   1410
      ItemData        =   "Frm_StudentAttendance.frx":08CA
      Left            =   7200
      List            =   "Frm_StudentAttendance.frx":08F2
      TabIndex        =   0
      Top             =   1320
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
      Height          =   1410
      ItemData        =   "Frm_StudentAttendance.frx":0955
      Left            =   5400
      List            =   "Frm_StudentAttendance.frx":097D
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frm_StudentAttendance.frx":09FB
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Add New Record Here"
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8705
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
      Caption         =   "Records Present"
      ColumnCount     =   8
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
         DataField       =   "Student_Name"
         Caption         =   "Student Name"
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
         DataField       =   "Student_Class"
         Caption         =   "Student Class"
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
         DataField       =   "Class_Year"
         Caption         =   "Class Year"
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
         DataField       =   "Subject"
         Caption         =   "Subject"
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
      BeginProperty Column05 
         DataField       =   "Month"
         Caption         =   "Month"
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
      BeginProperty Column06 
         DataField       =   "Working_Days"
         Caption         =   "Working Days"
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
      BeginProperty Column07 
         DataField       =   "Days_Present"
         Caption         =   "Days Present"
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
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1440
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
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
      RecordSource    =   "StudentAttendanceInformation"
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
   Begin VB.Image Image2 
      Height          =   360
      Left            =   9120
      Picture         =   "Frm_StudentAttendance.frx":0A10
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Record Student Monthly Attendance"
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
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   120
      Picture         =   "Frm_StudentAttendance.frx":117A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "Frm_StudentAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intHeight As Integer
Dim intCol As Integer
Dim intRow As Integer

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
    List3.Height = (.Height / .RowHeight - (intRow - 1)) * .RowHeight
    List3.Move .Left + .Columns(1).Left, _
     .Top + .RowTop(.Row) + .RowHeight, _
     .Columns(1).Width
    If Len(strItem) Then
       List3 = strItem
    Else
       List3.ListIndex = 0
    End If
    List3.Visible = True
    List3.SetFocus
  Case 2
    List4.Height = (.Height / .RowHeight - (intRow - 1)) * .RowHeight
    List4.Move .Left + .Columns(2).Left, _
     .Top + .RowTop(.Row) + .RowHeight, _
     .Columns(2).Width
    If Len(strItem) Then
       List4 = strItem
    Else
       List4.ListIndex = 0
    End If
    List4.Visible = True
    List4.SetFocus
  Case 3
    List1.Height = (.Height / .RowHeight - (intRow - 1)) * .RowHeight
    List1.Move .Left + .Columns(3).Left, _
     .Top + .RowTop(.Row) + .RowHeight, _
     .Columns(3).Width
    If Len(strItem) Then
       List1 = strItem
    Else
       List1.ListIndex = 0
    End If
    List1.Visible = True
    List1.SetFocus
  Case 5
    If intRow > 4 Then ' Place above cell
       List2.Height = (intRow + 1) * .RowHeight
       List2.Move .Left + .Columns(5).Left, _
        .Top + .RowHeight + (intRow * 1.4), _
        .Columns(5).Width
    Else ' Place below cell
       List2.Height = (.Height / .RowHeight - (intRow + 1)) * .RowHeight
       List2.Move .Left + .Columns(5).Left, _
        .Top + .RowTop(.Row) + .RowHeight, _
        .Columns(5).Width
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

Private Sub Form_Load()
' Events That Should Happen When Form Is Loaded
On Error Resume Next

List3.AddItem ""
List4.AddItem ""

globalstuin.Movefirst
Do While Not globalstuin.BOF And Not globalstuin.EOF
   List3.AddItem globalstuin(3).Value
   globalstuin.Movenext
Loop

stuattencourse.Movefirst
Do While Not stuattencourse.BOF And Not stuattencourse.EOF
   List4.AddItem stuattencourse(1).Value
   stuattencourse.Movenext
Loop

DataGrid1.Columns(5).Button = True
DataGrid1.Columns(3).Button = True
DataGrid1.Columns(1).Button = True
DataGrid1.Columns(2).Button = True

Me.Top = 50
Me.Left = 50
Merlin "Student Monthly Attendance Entry Is Done Here", "Read"
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
ElseIf List3.Visible Then
   List3.Visible = False
ElseIf List4.Visible Then
   List4.Visible = False
End If
End Sub
Private Sub DataGrid1_Scroll(Cancel As Integer)
On Error Resume Next
' Code To Be Executed When Datagrid Is Scrolled
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
ElseIf List3.Visible Then
   List3.Visible = False
ElseIf List4.Visible Then
   List4.Visible = False
End If
End Sub

Private Sub Form_Click()
On Error Resume Next
' When Clicked On The Form
If List1.Visible Then
   List1.Visible = False
ElseIf List2.Visible Then
   List2.Visible = False
ElseIf List3.Visible Then
   List3.Visible = False
ElseIf List4.Visible Then
   List4.Visible = False
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
Private Sub List3_Click()
' When List Box Is Clicked
On Error Resume Next
DataGrid1.Text = List3.Text
List3.Visible = False
End Sub

Private Sub List3_LostFocus()
' When List Box Lost Its Focus
On Error Resume Next
List3.Visible = False
End Sub

Private Sub List4_Click()
' When List Box Is Clicked
On Error Resume Next
DataGrid1.Text = List4.Text
List4.Visible = False
End Sub

Private Sub List4_LostFocus()
' When List Box Lost Its Focus
On Error Resume Next
List4.Visible = False
End Sub

