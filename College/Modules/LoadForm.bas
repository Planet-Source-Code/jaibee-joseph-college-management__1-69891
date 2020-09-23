Attribute VB_Name = "LoadForm"
' This Is The Module That Contain The Functions To Load Froms
' Every Form In The Application Is Loaded With The Functions Here
' These Functions Are Called By The Call Method To Load The Desired Form

Public Function loadstuentry()
Load Frm_StudentEntry
Frm_StudentEntry.Show
Frm_StudentEntry.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_StudentEntry.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_StudentEntry.studetailtab.TabStripBackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadcourseentry()
Load Frm_CourseEntry
Frm_CourseEntry.Show
Frm_CourseEntry.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_CourseEntry.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadsearchstudent()
Load Frm_SearchStudent
Frm_SearchStudent.Show
Frm_SearchStudent.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_SearchStudent.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadstudentpic()
Load Frm_StudentPic
Frm_StudentPic.Show
Frm_StudentPic.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_StudentPic.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadsearchstudentfamily()
Load Frm_SearchStudentFamily
Frm_SearchStudentFamily.Show
Frm_SearchStudentFamily.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_SearchStudentFamily.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadglobalmain()
Load Frm_GlobalMain
Frm_GlobalMain.Show
Frm_GlobalMain.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_GlobalMain.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadcourseinforeport()
Set CourseInfoReport.DataSource = courserec
Load CourseInfoReport
CourseInfoReport.Show
Merlin "Print Course Information Report"
End Function

Public Function loadstaffentry()
Load Frm_StaffEntry
Frm_StaffEntry.Show
Frm_StaffEntry.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_StaffEntry.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_StaffEntry.stfdetailtab.TabStripBackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Sub shownotepad()
On Error GoTo errcode
a = Shell("notepad.exe", vbNormalFocus)
Exit Sub
errcode:
MsgBox "Unable To Run Notepad Utility On Your Computer", vbInformation, "Error in opening!!!"
Resume Next
End Sub

Public Sub showcalculator()
On Error GoTo errHandle
a = Shell("calc.exe", vbNormalFocus)
Exit Sub
errHandle:
MsgBox "Unable To Run Calculator Utility On Your Computer", vbInformation, "Error in opening!!!"
Resume Next
End Sub

Public Function loadabout()
Load Frm_AboutInformationManager
Frm_AboutInformationManager.Show
End Function

Public Function loadcredit()
Load Frm_InformationManagerCredit
Frm_InformationManagerCredit.Show
End Function

Public Function loadwininfo()
Load Frm_WindowsInformation
Frm_WindowsInformation.Show
Frm_WindowsInformation.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_WindowsInformation.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadcreateuser()
Load Frm_CreateUser
Frm_CreateUser.Show
Frm_CreateUser.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_CreateUser.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loaddeleteuser()
Load Frm_DeleteUser
Frm_DeleteUser.Show
Frm_DeleteUser.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_DeleteUser.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadstaffpic()
Load Frm_StaffPic
Frm_StaffPic.Show
Frm_StaffPic.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_StaffPic.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadsearchstaff()
Load Frm_SearchStaff
Frm_SearchStaff.Show
Frm_SearchStaff.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_SearchStaff.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadglobalstaff()
Load Frm_GlobalMainStaff
Frm_GlobalMainStaff.Show
Frm_GlobalMainStaff.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_GlobalMainStaff.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadchangetheme()
Theme = Theme + 1
If Theme = 3 Then Theme = 0
MainMenu.ACPRibbon1.Theme = Theme
MainMenu.ACPRibbon1.Refresh
MainMenu.Picture = MainMenu.ACPRibbon1.LoadBackground
MainMenu.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_Sidebar.BackColor = MainMenu.ACPRibbon1.BackColor
Call checkthemeside
SaveSetting App.CompanyName, "ThemeSettings", "ThemeSet", Theme
End Function

Public Function loaddefaulttheme()
Theme = 1
MainMenu.ACPRibbon1.Theme = Theme
MainMenu.ACPRibbon1.Refresh
MainMenu.Picture = MainMenu.ACPRibbon1.LoadBackground
MainMenu.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_Sidebar.BackColor = MainMenu.ACPRibbon1.BackColor
Call checkthemeside
SaveSetting App.CompanyName, "ThemeSettings", "ThemeSet", Theme
End Function

Public Function loadbackupsoft()
If copyf.FileExists(App.Path & "\BackUp Software\BackRest.exe") Then
Shell App.Path & "\BackUp Software\BackRest.exe", vbNormalFocus
Else
MsgBox "Cannot Find BackUp Software File", vbInformation, "No BackUp File"
End If
End Function

Public Function loadprintfeereport()
Load Frm_PrintFeeReport
Frm_PrintFeeReport.Show
Frm_PrintFeeReport.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_PrintFeeReport.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadsearchfeeentry()
Load Frm_SearchFeeEntry
Frm_SearchFeeEntry.Show
Frm_SearchFeeEntry.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_SearchFeeEntry.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loaduniversityreciept()
Load Frm_UniversityReciept
Frm_UniversityReciept.Show
Frm_UniversityReciept.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_UniversityReciept.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadprintunireciept()
Set UniversityFeeReport.DataSource = universityrec
Load UniversityFeeReport
UniversityFeeReport.Show
Merlin "Print University Fee Receipt"
End Function

Public Function loadexporttoexcel()
Load Frm_ExportToExcel
Frm_ExportToExcel.Show
Frm_ExportToExcel.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_ExportToExcel.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadshowtooltipstart()
ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
If ShowAtStartup = 0 Then
ShowAtStartup = 1
MsgBox "Tool Tip Enabled", vbInformation, "Enabled"
Else
ShowAtStartup = 0
MsgBox "Tool Tip Disabled", vbInformation, "Disabled"
End If
SaveSetting App.EXEName, "Options", "Show Tips at Startup", ShowAtStartup
End Function

Public Function loadstaffatten()
Load Frm_StaffAttendance
Frm_StaffAttendance.Show
Frm_StaffAttendance.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_StaffAttendance.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadsearchstaffatten()
Load Frm_SearchStaffAttendance
Frm_SearchStaffAttendance.Show
Frm_SearchStaffAttendance.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_SearchStaffAttendance.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadstaffattenprint()
Load Frm_StaffAtnPrint
Frm_StaffAtnPrint.Show
Frm_StaffAtnPrint.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_StaffAtnPrint.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadstuatten()
Load Frm_StudentAttendance
Frm_StudentAttendance.Show
Frm_StudentAttendance.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_StudentAttendance.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadsearchstuatten()
Load Frm_SearchStudentAttendance
Frm_SearchStudentAttendance.Show
Frm_SearchStudentAttendance.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_SearchStudentAttendance.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadprintstuatten()
Load Frm_PrintStudentAttendance
Frm_PrintStudentAttendance.Show
Frm_PrintStudentAttendance.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_PrintStudentAttendance.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadsearchsalentry()
Load Frm_SearchSalaryEntry
Frm_SearchSalaryEntry.Show
Frm_SearchSalaryEntry.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_SearchSalaryEntry.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadprintsal()
Set StaffSalaryReport.DataSource = staffsalary
Load StaffSalaryReport
StaffSalaryReport.Show
Merlin "Print Staff Salary Entry"
End Function

Public Function loadchangeadmin()
Load Frm_ChangeAdministrator
Frm_ChangeAdministrator.Show
Frm_ChangeAdministrator.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_ChangeAdministrator.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadglobalexam()
Load Frm_GlobalExamEntry
Frm_GlobalExamEntry.Show
Frm_GlobalExamEntry.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_GlobalExamEntry.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadstumark()
Load Frm_StudentMarkEntry
Frm_StudentMarkEntry.Show
Frm_StudentMarkEntry.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_StudentMarkEntry.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadsearchmark()
Load Frm_SearchMarkEntry
Frm_SearchMarkEntry.Show
Frm_SearchMarkEntry.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_SearchMarkEntry.BackColor = MainMenu.ACPRibbon1.BackColor
End Function

Public Function loadstumarkrep()
Set StudentMarkReport.DataSource = markprint
Load StudentMarkReport
StudentMarkReport.Show
Merlin "Print Student Mark Entry From Here"
End Function

Public Function loadcalendarcon()
Load Frm_CalendarForm
Frm_CalendarForm.Show
End Function

Public Function loadprintstudetail()
Load Frm_PrintStudentReport
Frm_PrintStudentReport.Show
Frm_PrintStudentReport.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_PrintStudentReport.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadprintstaffdetail()
Set StaffDetailReport.DataSource = staffrec
Load StaffDetailReport
StaffDetailReport.Show
Merlin "Print Staff Detail Here"
End Function

Public Function showhelpfile()
  lngResult = ShellExecute.LaunchDocument( _
    MainMenu.hwnd, _
    App.Path & "\Help Files\Help File.pdf", _
    CurDir, sesSW_SHOWDEFAULT)
  
  If lngResult <> seeNoError Then
    MsgBox "Error on LaunchDocument: " & lngResult
  End If
End Function

Public Function showuserdemo()
If copyf.FileExists(App.Path & "\User Demo\IMSDEMO.exe") = True Then
Shell App.Path & "\User Demo\IMSDEMO.exe", vbNormalFocus
Else
MsgBox "Cannot Found File", vbCritical, "File Error"
End If
End Function

Public Function loadquerybuilder()
Load Frm_QueryBuilder
Frm_QueryBuilder.Show
Frm_QueryBuilder.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_QueryBuilder.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function loadaboutinstitution()
Load Frm_AboutInstitute
Frm_AboutInstitute.Show
Frm_AboutInstitute.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_AboutInstitute.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function

Public Function changecharload()
Load Frm_ChangeCharacter
Frm_ChangeCharacter.Show
Frm_ChangeCharacter.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_ChangeCharacter.Picture = MainMenu.ACPRibbon1.LoadBackground
End Function
