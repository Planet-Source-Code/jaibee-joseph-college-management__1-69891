VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Begin VB.Form Frm_StudentMarkEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Mark Entry"
   ClientHeight    =   5340
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9615
   Icon            =   "Frm_StudentMarkEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9615
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
      Height          =   735
      Left            =   4440
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
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
      Height          =   735
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
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
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
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
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1080
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
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
      RecordSource    =   "StudentMarkInformation"
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
      Bindings        =   "Frm_StudentMarkEntry.frx":076A
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enter Student Exam Marks Here"
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8070
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
      Caption         =   "Existing Records"
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
         DataField       =   "Exam_Type"
         Caption         =   "Exam Type"
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
         DataField       =   "Exam_Date"
         Caption         =   "Exam Date"
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
         DataField       =   "Max_Mark"
         Caption         =   "Max Mark"
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
         DataField       =   "Min_Mark"
         Caption         =   "Min Mark"
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
         DataField       =   "Mark_Obtained"
         Caption         =   "Mark Obtained"
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
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   9120
      Picture         =   "Frm_StudentMarkEntry.frx":077F
      ToolTipText     =   "Application Help"
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Exam Mark Entry"
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
      Top             =   240
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Frm_StudentMarkEntry.frx":0EE9
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Frm_StudentMarkEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intHeight As Integer
Dim intCol As Integer
Dim intRow As Integer

Private Sub Form_Load()
On Error Resume Next

List1.AddItem ""
List2.AddItem ""
List3.AddItem ""
List4.AddItem ""

globalmark.Movefirst
Do While Not globalmark.BOF And Not globalmark.EOF
   List2.AddItem globalmark(1).Value
   List3.AddItem globalmark(2).Value
   List4.AddItem globalmark(3).Value
   globalmark.Movenext
Loop

stumastin.Movefirst
Do While Not stumastin.BOF And Not stumastin.EOF
   List1.AddItem stumastin(3).Value
   stumastin.Movenext
Loop

Me.Top = 50
Me.Left = 50

DataGrid1.Columns(1).Button = True
DataGrid1.Columns(2).Button = True
DataGrid1.Columns(5).Button = True
DataGrid1.Columns(6).Button = True

Merlin "This Is Where You Enter Student Exam Mark Entry", "Read"
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
  Case 5
    If intRow > 4 Then ' Place above cell
       List3.Height = (intRow + 1) * .RowHeight
       List3.Move .Left + .Columns(5).Left, _
        .Top + .RowHeight + (intRow * 1.4), _
        .Columns(5).Width
    Else ' Place below cell
       List3.Height = (.Height / .RowHeight - (intRow + 1)) * .RowHeight
       List3.Move .Left + .Columns(5).Left, _
        .Top + .RowTop(.Row) + .RowHeight, _
        .Columns(5).Width
    End If
    If Len(strItem) Then
       ' Find match in Listbox
       Dim m As Integer
       For m = 0 To List3.ListCount - 1
           If strItem = List3.List(m) Then
              ' List2.Selected(n) = True
              List3.ListIndex = m
              Exit For
           End If
       Next
    Else
       List3.ListIndex = 0
    End If
    List3.Visible = True
    List3.SetFocus
  Case 6
    If intRow > 4 Then ' Place above cell
       List4.Height = (intRow + 1) * .RowHeight
       List4.Move .Left + .Columns(6).Left, _
        .Top + .RowHeight + (intRow * 1.4), _
        .Columns(6).Width
    Else ' Place below cell
       List4.Height = (.Height / .RowHeight - (intRow + 1)) * .RowHeight
       List4.Move .Left + .Columns(6).Left, _
        .Top + .RowTop(.Row) + .RowHeight, _
        .Columns(6).Width
    End If
    If Len(strItem) Then
       ' Find match in Listbox
       Dim o As Integer
       For o = 0 To List4.ListCount - 1
           If strItem = List4.List(o) Then
              ' List2.Selected(n) = True
              List4.ListIndex = o
              Exit For
           End If
       Next
    Else
       List4.ListIndex = 0
    End If
    List4.Visible = True
    List4.SetFocus

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

