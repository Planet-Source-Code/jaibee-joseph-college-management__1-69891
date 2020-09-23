VERSION 5.00
Begin VB.Form Frm_PopMenu 
   BorderStyle     =   0  'None
   Caption         =   "Dialog Caption"
   ClientHeight    =   30
   ClientLeft      =   2715
   ClientTop       =   3705
   ClientWidth     =   2310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Begin VB.Menu MainMenuFrm 
      Caption         =   "Main Menu"
      Begin VB.Menu StuManage 
         Caption         =   "Student Management"
         Begin VB.Menu newstufee 
            Caption         =   "New Student and Fee Entry"
         End
         Begin VB.Menu seastu 
            Caption         =   "Search Student Information"
         End
         Begin VB.Menu viestupic 
            Caption         =   "View Preview of Student Pic"
         End
         Begin VB.Menu seastufam 
            Caption         =   "Search Student Family Info"
         End
         Begin VB.Menu seastufeein 
            Caption         =   "Search Student Fee Info"
         End
         Begin VB.Menu unfeeen 
            Caption         =   "University Fee Entry"
         End
         Begin VB.Menu stumarken 
            Caption         =   "Student Mark Entry"
         End
         Begin VB.Menu searmarkentry 
            Caption         =   "Search Student Mark Entry"
         End
         Begin VB.Menu studentmonatten 
            Caption         =   "Student Mon Atten Entry"
         End
         Begin VB.Menu searattenentry 
            Caption         =   "Search Attendance Entry"
         End
      End
      Begin VB.Menu stamana 
         Caption         =   "Staff Management"
         Begin VB.Menu newstaffsal 
            Caption         =   "New Staff and Salary Entry"
         End
         Begin VB.Menu viewstapic 
            Caption         =   "View Staff Picture Preview"
         End
         Begin VB.Menu searstaffinfo 
            Caption         =   "Search Staff Information"
         End
         Begin VB.Menu staffattnentry 
            Caption         =   "Staff Attendance Entry"
         End
         Begin VB.Menu seastaffatten 
            Caption         =   "Search Staff Attendance"
         End
         Begin VB.Menu staffsalentry 
            Caption         =   "Search Staff Salary Entry"
         End
      End
      Begin VB.Menu adminmanage 
         Caption         =   "Admin Management"
         Begin VB.Menu createuser 
            Caption         =   "Create User Info"
         End
         Begin VB.Menu Deluser 
            Caption         =   "Delete User Info"
         End
         Begin VB.Menu chasoftthe 
            Caption         =   "Change Software Theme"
         End
         Begin VB.Menu defthe 
            Caption         =   "Default Theme"
         End
         Begin VB.Menu dbback 
            Caption         =   "Database BackUp"
         End
         Begin VB.Menu enatool 
            Caption         =   "Enable or Disable ToolTip"
         End
      End
      Begin VB.Menu printing 
         Caption         =   "Printing Management"
         Begin VB.Menu pricourse 
            Caption         =   "Print Course Information"
         End
         Begin VB.Menu prifeeinfo 
            Caption         =   "Print Fee Information"
         End
         Begin VB.Menu priunfeeen 
            Caption         =   "Print University Fee Entry"
         End
         Begin VB.Menu ptistaffdai 
            Caption         =   "Print Staff Daily Attendance"
         End
         Begin VB.Menu primonstu 
            Caption         =   "Print Student Mon Attendance"
         End
         Begin VB.Menu pristaffsal 
            Caption         =   "Print Staff Salary Entry"
         End
         Begin VB.Menu ptistumarken 
            Caption         =   "Print Student Mark Entry"
         End
         Begin VB.Menu printstuinfodetail 
            Caption         =   "Print Student Information"
         End
         Begin VB.Menu printstaffinfo 
            Caption         =   "Print Staff Information"
         End
      End
      Begin VB.Menu oherutili 
         Caption         =   "Other Utility Management"
         Begin VB.Menu notepa 
            Caption         =   "Open Notepad"
         End
         Begin VB.Menu opcal 
            Caption         =   "Open Calculator"
         End
         Begin VB.Menu osinfo 
            Caption         =   "OS Information"
         End
         Begin VB.Menu exporttoexcel 
            Caption         =   "Export Table To Excel"
         End
         Begin VB.Menu calendarims 
            Caption         =   "IMS Calendar"
         End
         Begin VB.Menu sqlquerybuild 
            Caption         =   "SQL Query Builder"
         End
         Begin VB.Menu ChangeHelpAssis 
            Caption         =   "Change Your Help Assistant"
         End
      End
      Begin VB.Menu globalen 
         Caption         =   "Global Entry Management"
         Begin VB.Menu newcou 
            Caption         =   "New Course Entry"
         End
         Begin VB.Menu stuglobal 
            Caption         =   "Student Global Information"
         End
         Begin VB.Menu staffGlobal 
            Caption         =   "Staff Global Entries"
         End
         Begin VB.Menu stumarkglo 
            Caption         =   "Student Mark Global"
         End
      End
      Begin VB.Menu applihelp 
         Caption         =   "Application Help"
         Begin VB.Menu usermanu 
            Caption         =   "User Manual"
         End
         Begin VB.Menu appdemo 
            Caption         =   "Application Demo"
         End
         Begin VB.Menu dishelpassis 
            Caption         =   "Disable/Enable Help Assistant"
         End
      End
      Begin VB.Menu seperator5 
         Caption         =   "-"
      End
      Begin VB.Menu refdatainfo 
         Caption         =   "Refresh Database Information"
      End
      Begin VB.Menu seperator 
         Caption         =   "-"
      End
      Begin VB.Menu abtinfomana 
         Caption         =   "About Information Management System"
      End
      Begin VB.Menu creditsinfo 
         Caption         =   "Credits Information Management System"
      End
      Begin VB.Menu AboutInstitute 
         Caption         =   "About The Current Institution"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu exitmenu 
         Caption         =   "Exit From Software"
      End
   End
End
Attribute VB_Name = "Frm_PopMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub AboutInstitute_Click()
Call loadaboutinstitution
End Sub

Private Sub abtinfomana_Click()
Call loadabout
End Sub

Private Sub appdemo_Click()
Call showuserdemo
End Sub

Private Sub calendarims_Click()
Call loadcalendarcon
End Sub

Private Sub ChangeHelpAssis_Click()
Call changecharload
End Sub

Private Sub chasoftthe_Click()
Call loadchangetheme
End Sub

Private Sub createuser_Click()
Call loadcreateuser
End Sub

Private Sub creditsinfo_Click()
Call loadcredit
End Sub

Private Sub dbback_Click()
Call loadbackupsoft
End Sub

Private Sub defthe_Click()
Call loaddefaulttheme
End Sub

Private Sub Deluser_Click()
Call loaddeleteuser
End Sub

Private Sub dishelpassis_Click()
Call disablemerlin
End Sub

Private Sub enatool_Click()
Call loadshowtooltipstart
End Sub

Private Sub exitmenu_Click()
On Error Resume Next
Unload MainMenu
End Sub

Private Sub exporttoexcel_Click()
Call loadexporttoexcel
End Sub

Private Sub Form_Load()
On Error Resume Next
If admcheck <> "Administrator" Then
Frm_PopMenu.adminmanage.Enabled = False
End If
End Sub

Private Sub newcou_Click()
Call loadcourseentry
End Sub

Private Sub newstaffsal_Click()
Call loadstaffentry
End Sub

Private Sub newstufee_Click()
Call loadstuentry
End Sub

Private Sub notepa_Click()
Call shownotepad
End Sub

Private Sub opcal_Click()
Call showcalculator
End Sub

Private Sub osinfo_Click()
Call loadwininfo
End Sub

Private Sub pricourse_Click()
Call loadcourseinforeport
End Sub

Private Sub prifeeinfo_Click()
Call loadprintfeereport
End Sub

Private Sub primonstu_Click()
Call loadprintstuatten
End Sub

Private Sub printstaffinfo_Click()
Call loadprintstaffdetail
End Sub

Private Sub printstuinfodetail_Click()
Call loadprintstudetail
End Sub

Private Sub pristaffsal_Click()
Call loadprintsal
End Sub

Private Sub priunfeeen_Click()
Call loadprintunireciept
End Sub

Private Sub ptistaffdai_Click()
Call loadstaffattenprint
End Sub

Private Sub ptistumarken_Click()
Call loadstumarkrep
End Sub

Private Sub refdatainfo_Click()
On Error Resume Next
Call MainConClose
Call MainConEstablish
End Sub

Private Sub searattenentry_Click()
Call loadsearchstuatten
End Sub

Private Sub searmarkentry_Click()
Call loadsearchmark
End Sub

Private Sub searstaffinfo_Click()
Call loadsearchstaff
End Sub

Private Sub seastaffatten_Click()
Call loadsearchstaffatten
End Sub

Private Sub seastu_Click()
Call loadsearchstudent
End Sub

Private Sub seastufam_Click()
Call loadsearchstudentfamily
End Sub

Private Sub seastufeein_Click()
Call loadsearchfeeentry
End Sub

Private Sub sqlquerybuild_Click()
Call loadquerybuilder
End Sub

Private Sub staffattnentry_Click()
Call loadstaffatten
End Sub

Private Sub staffGlobal_Click()
Call loadglobalstaff
End Sub

Private Sub staffsalentry_Click()
Call loadsearchsalentry
End Sub

Private Sub studentmonatten_Click()
Call loadstuatten
End Sub

Private Sub stuglobal_Click()
Call loadglobalmain
End Sub

Private Sub stumarken_Click()
Call loadstumark
End Sub

Private Sub stumarkglo_Click()
Call loadglobalexam
End Sub

Private Sub unfeeen_Click()
Call loaduniversityreciept
End Sub

Private Sub usermanu_Click()
Call showhelpfile
End Sub

Private Sub viestupic_Click()
Call loadstudentpic
End Sub

Private Sub viewstapic_Click()
Call loadstaffpic
End Sub
