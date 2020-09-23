VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkUserControlsXP.ocx"
Begin VB.MDIForm MainMenu 
   BackColor       =   &H8000000C&
   Caption         =   "Information Management System Main Menu"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   -30
   ClientWidth     =   9285
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin vkUserContolsXP.vkSysTray vkSysTray1 
      Left            =   2880
      Top             =   3960
      _ExtentX        =   794
      _ExtentY        =   794
      BalloonTipString=   "Information Management System 1.0.0"
      Icon            =   "MainMenu.frx":076A
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5865
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Show CAPS Enabled Or Not"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Shows NUM Enabled Or Not"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
            Object.ToolTipText     =   "Shows INS Enabled Or Not"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "3:25 PM"
            Object.ToolTipText     =   "Shows System Time"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "1/11/2008"
            Object.ToolTipText     =   "Shows System Date"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
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
   End
   Begin CollegeManagement.ACPRibbon ACPRibbon1 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _extentx        =   16378
      _extenty        =   3069
      backcolor       =   4210752
      forecolor       =   -2147483630
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   48
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":0EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":165E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":2552
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":2CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":3446
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":3BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":433A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":4AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":522E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":59A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":6122
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":689C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":7016
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":7790
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":7F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":8684
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":8DFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":9578
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":9CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":A46C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":ABE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":B360
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":BADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":C254
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":C9CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":D148
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":D8C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":E03C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":E7B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":EF30
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":F6AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":FE24
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1059E
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":10D18
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":11492
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":11C0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":12386
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":12B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1327A
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":139F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1416E
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":148E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":15062
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":157DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":15F56
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":166D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":16E4A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   2880
      Top             =   3360
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ACPRibbon1_ButtonClick(ByVal ID As String, ByVal Caption As String)
' When A Button In The Ribbon Is Clicked
On Error GoTo errmessage

If ID = 0 Then
Call loadstuentry

ElseIf ID = 1 Then
Call loadcourseentry

ElseIf ID = 2 Then
Call loadsearchstudent

ElseIf ID = 3 Then
Call loadstudentpic

ElseIf ID = 4 Then
Call loadsearchstudentfamily

ElseIf ID = 5 Then
Call loadglobalmain

ElseIf ID = 6 Then
Call loadcourseinforeport

ElseIf ID = 7 Then
Call loadstaffentry

ElseIf ID = 8 Then
Call shownotepad

ElseIf ID = 9 Then
Call showcalculator

ElseIf ID = 10 Then
Call loadabout

ElseIf ID = 11 Then
Call loadcredit

ElseIf ID = 12 Then
Call loadwininfo

ElseIf ID = 13 Then
Call loadcreateuser

ElseIf ID = 14 Then
Call loaddeleteuser

ElseIf ID = 15 Then
Call loadstaffpic

ElseIf ID = 16 Then
Call loadsearchstaff

ElseIf ID = 17 Then
Call loadglobalstaff

ElseIf ID = 18 Then
Call loadchangetheme

ElseIf ID = 19 Then
Call loaddefaulttheme

ElseIf ID = 20 Then
Call loadbackupsoft

ElseIf ID = 21 Then
Call loadprintfeereport

ElseIf ID = 22 Then
Call loadsearchfeeentry

ElseIf ID = 23 Then
Call loaduniversityreciept

ElseIf ID = 24 Then
Call loadprintunireciept

ElseIf ID = 25 Then
Call loadexporttoexcel

ElseIf ID = 26 Then
Call loadshowtooltipstart

ElseIf ID = 27 Then
Call loadstaffatten

ElseIf ID = 28 Then
Call loadsearchstaffatten

ElseIf ID = 29 Then
Call loadstaffattenprint

ElseIf ID = 30 Then
Call loadstuatten

ElseIf ID = 31 Then
Call loadsearchstuatten

ElseIf ID = 32 Then
Call loadprintstuatten

ElseIf ID = 33 Then
Call loadsearchsalentry

ElseIf ID = 34 Then
Call loadprintsal

ElseIf ID = 35 Then
Call loadchangeadmin

ElseIf ID = 36 Then
Call loadglobalexam

ElseIf ID = 37 Then
Call loadstumark

ElseIf ID = 38 Then
Call loadsearchmark

ElseIf ID = 39 Then
Call loadstumarkrep

ElseIf ID = 40 Then
Call loadcalendarcon

ElseIf ID = 41 Then
Call loadprintstudetail

ElseIf ID = 42 Then
Call showhelpfile

ElseIf ID = 43 Then
Call showuserdemo

ElseIf ID = 44 Then
Call loadprintstaffdetail

ElseIf ID = 46 Then
Call loadquerybuilder

ElseIf ID = 47 Then
Call disablemerlin

ElseIf ID = 48 Then
Call loadaboutinstitution

ElseIf ID = 49 Then
Call changecharload

End If

Exit Sub
errmessage:
MsgBox Err.Description, vbCritical, "Error Occured"
End Sub

Private Sub ACPRibbon1_TabClick(ByVal ID As String, ByVal Caption As String)
' Code For Checking Whether Administrator Is Login Or User
On Error Resume Next
If ID = "3" Then
If admcheck <> "Administrator" Then
MsgBox "Administrator Area Is Not Allowed For Users", vbInformation, "User Error"
End
End If
End If

If ID = 1 Then
Merlin "This Is The Area Where You Can Manage All The Student Information", "Read"
ElseIf ID = 2 Then
Merlin "This Is The Area Where You Can Add All The Staff Information", "DoMagic1"
ElseIf ID = 3 Then
Merlin "This Is The Administrator's Area Only The Administrator Can Do The Desired Changes", "DoMagic2"
ElseIf ID = 4 Then
Merlin "This Is The Area For All Sorts Of Printing Works", "Read"
ElseIf ID = 5 Then
Merlin "This Is The Area Where You Can Access To Other Utilities Of The Software", "DoMagic1"
ElseIf ID = 6 Then
Merlin "This Is The Area Where You Can Do All The Global Enteries", "DoMagic2"
ElseIf ID = 7 Then
Merlin "This Is The Area Where You Can Acess To The Help Of Software", "Read"
End If

End Sub

Private Sub MDIForm_Load()
' Code To Be Executed When Form Is Loaded
On Error GoTo label

vkSysTray1.AddToTray

Theme = GetSetting(App.CompanyName, "ThemeSettings", "ThemeSet", Theme)

'# SET Theme
ACPRibbon1.Theme = Theme
                        
'# OPTIONAL - Load Background for Form.
MainMenu.Picture = ACPRibbon1.LoadBackground

'# OPTIONAL - Load Background for Form
MainMenu.BackColor = ACPRibbon1.BackColor

'# Set ImageList to use for icons
ACPRibbon1.ImageList = ImageList1

'# Set Buttons on Center verticaly    (True = Center, False(Default) = Align on Top)
ACPRibbon1.ButtonCenter = False

'# Add Tabs ---   ID - Caption
ACPRibbon1.AddTab "1", "Student Management"
ACPRibbon1.AddTab "2", "Staff Management"
ACPRibbon1.AddTab "3", "Administrator Area"
ACPRibbon1.AddTab "4", "Printing Area"
ACPRibbon1.AddTab "5", "Other Utilities"
ACPRibbon1.AddTab "6", "Global Entries"
ACPRibbon1.AddTab "7", "Software Help"


'# Add Cats ---   ID - Tab - Caption - ShowDialogButton
ACPRibbon1.AddCat "1", "1", "Student Management Area", False
ACPRibbon1.AddCat "2", "6", "Global Insertions", False
ACPRibbon1.AddCat "3", "4", "Printing Selections", False
ACPRibbon1.AddCat "4", "2", "Staff Management Area", False
ACPRibbon1.AddCat "5", "5", "Other Utilities Area", False
ACPRibbon1.AddCat "6", "7", "Application Help Area", False
ACPRibbon1.AddCat "7", "3", "Application Admin Area", False


'# Add Button ---    ID - Cat - Capt. - Icons -   More Arrow   - ToolTip
ACPRibbon1.AddButton "0", "1", "New Student And" & vbCrLf & "Fee Entry", 1, False, "Enter New Student Information"
ACPRibbon1.AddButton "1", "2", "Add Information Of" & vbCrLf & "New Course", 2, False, "Enter New Course Information"
ACPRibbon1.AddButton "2", "1", "Search Information" & vbCrLf & "Of Student", 3, False, "Search Student Information"
ACPRibbon1.AddButton "3", "1", "View Preview Of" & vbCrLf & "Student Pictures", 4, False, "View Student Picture Preview"
ACPRibbon1.AddButton "4", "1", "Search Student" & vbCrLf & "Family Information", 5, False, "Search Student Family Information"
ACPRibbon1.AddButton "5", "2", "Add New Global" & vbCrLf & "Entry (Student Table)", 6, False, "All Main Global Entry Form For Student Form"
ACPRibbon1.AddButton "6", "3", "Print Course" & vbCrLf & "Information Report", 7, False, "Print All Course Information"
ACPRibbon1.AddButton "7", "4", "New Staff Info And" & vbCrLf & "Salary Entry", 8, False, "Enter New Staff Information"
ACPRibbon1.AddButton "8", "5", "Open Notepad" & vbCrLf & "Utility", 9, False, "Open Windows Notepad Utility"
ACPRibbon1.AddButton "9", "5", "Open Calculator" & vbCrLf & "Utility", 10, False, "Open Windows Calculator Utility"
ACPRibbon1.AddButton "10", "6", "About Information" & vbCrLf & "Management System", 11, False, "About Information Management System"
ACPRibbon1.AddButton "11", "6", "Credits Information" & vbCrLf & "Management System", 12, False, "Credits Information Management System"
ACPRibbon1.AddButton "12", "5", "Operating System" & vbCrLf & "Information", 13, False, "Operating System Information"
ACPRibbon1.AddButton "13", "7", "Create New Software" & vbCrLf & "User Information", 14, False, "Create New Software User"
ACPRibbon1.AddButton "14", "7", "Delete Software" & vbCrLf & "User Information", 15, False, "Delete Software User"
ACPRibbon1.AddButton "15", "4", "View Preview Of" & vbCrLf & "Staff Pictures", 4, False, "View Preview Of Staff Pictures"
ACPRibbon1.AddButton "16", "4", "Search Staff" & vbCrLf & "Information", 16, False, "Staff Information Search"
ACPRibbon1.AddButton "17", "2", "Add New Global" & vbCrLf & "Entry (Staff Table)", 17, False, "All Main Global Entry Form For Staff Form"
ACPRibbon1.AddButton "18", "7", "Change Software" & vbCrLf & "Theme", 18, False, "To Change Software Theme"
ACPRibbon1.AddButton "19", "7", "Change Software" & vbCrLf & "Theme To Default", 19, False, "Change Software Theme To Default"
ACPRibbon1.AddButton "20", "7", "Make Software" & vbCrLf & "DB BackUp", 20, False, "Software DB BackUp"
ACPRibbon1.AddButton "21", "3", "Print Fee" & vbCrLf & "Information Report", 21, False, "Print Fee Information Report"
ACPRibbon1.AddButton "22", "1", "Search Student" & vbCrLf & "Fee Information", 16, False, "Search Fee Information "
ACPRibbon1.AddButton "23", "1", "University Fee" & vbCrLf & "Reciept Entry", 22, False, "University Fee Information Entry"
ACPRibbon1.AddButton "24", "3", "Print University Fee" & vbCrLf & "Reciept Entry", 23, False, "University Fee Information Print"
ACPRibbon1.AddButton "25", "5", "Export Tables" & vbCrLf & "To Excel Database", 24, False, "Export Tables To Excel Database"
ACPRibbon1.AddButton "26", "7", "Enable Or Disable" & vbCrLf & "Software ToolTip", 25, False, "Enable Or Disable ToolTip"
ACPRibbon1.AddButton "27", "4", "Insert Staff" & vbCrLf & "Daily Attendance", 26, False, "Mark Staff Daily Attendance"
ACPRibbon1.AddButton "28", "4", "Search Staff" & vbCrLf & "Daily Attendance", 27, False, "Search Staff Daily Attendance"
ACPRibbon1.AddButton "29", "3", "Staff Daily" & vbCrLf & "Attendance Report", 28, False, "Print Staff Daily Attendance"
ACPRibbon1.AddButton "30", "1", "Mark Student" & vbCrLf & "Monthly Attendance", 29, False, "Mark Student Monthly Attendance"
ACPRibbon1.AddButton "31", "1", "Search Student" & vbCrLf & "Monthly Attendance", 30, False, "Search Student Monthly Attendance"
ACPRibbon1.AddButton "32", "3", "Print Student" & vbCrLf & "Monthly Attendance", 31, False, "Print Student Monthly Attendance"
ACPRibbon1.AddButton "33", "4", "Search Staff" & vbCrLf & "Salary/Month Entry", 32, False, "Search Staff Salary Entry"
ACPRibbon1.AddButton "34", "3", "Print Staff" & vbCrLf & "Salary/Month Entry", 33, False, "Print Staff Salary Entry"
ACPRibbon1.AddButton "35", "7", "Change Admin User" & vbCrLf & "And Password", 34, False, "Change Administrator UserName And Password"
ACPRibbon1.AddButton "36", "2", "Global Exam Type" & vbCrLf & "Entry (Mark Form)", 35, False, "Global Examination Type Entry"
ACPRibbon1.AddButton "37", "1", "Student Exam Mark" & vbCrLf & "Entry (Mark Form)", 36, False, "Student Exam Mark Entry"
ACPRibbon1.AddButton "38", "1", "Search Exam Mark" & vbCrLf & "Entry (Mark Form)", 37, False, "Search Exam Mark Entry"
ACPRibbon1.AddButton "39", "3", "Print Exam Mark" & vbCrLf & "Entry (All Mark)", 38, False, "Print Exam Mark Entry"
ACPRibbon1.AddButton "40", "5", "Show IMS" & vbCrLf & "Calendar Control", 39, False, "Show Calendar"
ACPRibbon1.AddButton "41", "3", "Print Student" & vbCrLf & "Information Detail", 40, False, "Print Student Information Detail"
ACPRibbon1.AddButton "42", "6", "User Help For" & vbCrLf & "IMS 1.0.0", 41, False, "User Help For IMS"
ACPRibbon1.AddButton "43", "6", "User Tutorial For" & vbCrLf & "IMS 1.0.0", 42, False, "User Tutorial For IMS"
ACPRibbon1.AddButton "44", "3", "Print Staff Entry" & vbCrLf & "Information", 43, False, "Print Staff Entry Information"
ACPRibbon1.AddButton "45", "5", namesqlserver & vbCrLf & namesqldatabase, 45, False, "Name Of SQL Server and Database"
ACPRibbon1.AddButton "46", "5", "Build SQL 2005" & vbCrLf & "Queries", 46, False, "SQL Server Query Builder"
ACPRibbon1.AddButton "47", "6", "Enable Or Disable" & vbCrLf & "Help Assistant", 47, False, "Enable Or Disable Help Assistant"
ACPRibbon1.AddButton "48", "6", "About The Current" & vbCrLf & "Institution", 48, False, "About The Institution"
ACPRibbon1.AddButton "49", "5", "Select Your" & vbCrLf & "Help Assistant", 42, False, "Help Assistant Selection"


'# Repaint Ribbon
ACPRibbon1.Refresh

characterop = GetSetting(App.CompanyName, "Character", "MyAssist", characterop)
If characterop = "" Then
characterop = "Enabled"
SaveSetting App.CompanyName, "Character", "MyAssist", characterop
End If

'Initialize Agent
If characterop = "Enabled" Then
MyAgent.Characters.Load charactername, charfile
Set mycharacter = MyAgent.Characters(charactername)
mycharacter.SoundEffectsOn = True
mycharacter.Show
mycharacter.MoveTo 850, 550
End If

Exit Sub
label:
MsgBox Err.Description, vbCritical, "Error Occured"
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Code To Show PopUp Menu When Mouse Button Is Clicked
On Error GoTo message

If Button = 2 Then
Me.PopupMenu Frm_PopMenu.MainMenuFrm
End If

Exit Sub
message:
MsgBox Err.Description, vbCritical, "Error Occured"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
' When Main Form Is Unloaded
On Error GoTo label

Call MainConClose
Call backuppicture
Unload Frm_PopMenu
Unload Frm_UserMessage
vkSysTray1.RemoveFromTray

Exit Sub
label:
MsgBox Err.Description, vbCritical, "Error Occured"
End Sub

