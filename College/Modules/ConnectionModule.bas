Attribute VB_Name = "ConnectionModule"
' Main Database Connection Declarations
Public studentrec As ADODB.Recordset
Public studentcon As ADODB.Connection
Public familycon As ADODB.Connection
Public familyrec As ADODB.Recordset
Public coursecon As ADODB.Connection
Public courserec As ADODB.Recordset
Public feescon As ADODB.Connection
Public feesrec As ADODB.Recordset
Public staffcon As ADODB.Connection
Public staffrec As ADODB.Recordset
Public GlobalCon As ADODB.Connection
Public GlobalCaste As ADODB.Recordset
Public GlobalReligion As ADODB.Recordset
Public GlobalNationality As ADODB.Recordset
Public GlobalOccupationF As ADODB.Recordset
Public GlobalOccupationM As ADODB.Recordset
Public GlobalDesignation As ADODB.Recordset
Public GlobalDepartment As ADODB.Recordset
Public universitycon As ADODB.Connection
Public universityrec As ADODB.Recordset
Public staffatnrec As ADODB.Recordset
Public staffsalary As ADODB.Recordset
Public globalmark As ADODB.Recordset
Public markprint As ADODB.Recordset
Public globalstuin As New ADODB.Recordset
Public staffstnen As New ADODB.Recordset
Public departglo As New ADODB.Recordset
Public stumastin As New ADODB.Recordset
Public searuniversity As ADODB.Recordset
Public unifeeentrycou As New ADODB.Recordset
Public exportdepart As New ADODB.Recordset
Public exportcourse As New ADODB.Recordset
Public coursestudent As New ADODB.Recordset
Public queryobjectims As New ADODB.Recordset
Public stuattencourse As New ADODB.Recordset

' Function To Open All Connections
Public Function MainConEstablish()
   
Set studentcon = New ADODB.Connection
studentcon.CursorLocation = adUseClient
studentcon.Open "Driver={SQL Server};" & _
namesqlserver & _
namesqldatabase

slct = "select * from StudentInformation Order by Admission_Number"
Set studentrec = New ADODB.Recordset
studentrec.Open slct, studentcon, adOpenStatic, adLockOptimistic

Set familycon = New ADODB.Connection
familycon.CursorLocation = adUseClient
familycon.Open "Driver={SQL Server};" & _
namesqlserver & _
namesqldatabase

slcf = "select * from FamilyInformation Order by Admission_Number"
Set familyrec = New ADODB.Recordset
familyrec.Open slcf, familycon, adOpenStatic, adLockOptimistic


Set coursecon = New ADODB.Connection
coursecon.CursorLocation = adUseClient
coursecon.Open "Driver={SQL Server};" & _
namesqlserver & _
namesqldatabase

slcc = "select * from CourseInformation"
Set courserec = New ADODB.Recordset
courserec.Open slcc, coursecon, adOpenStatic, adLockOptimistic


Set feescon = New ADODB.Connection
feescon.CursorLocation = adUseClient
feescon.Open "Driver={SQL Server};" & _
namesqlserver & _
namesqldatabase

slfc = "select * from FeesInformation Order by Admission_Number"
Set feesrec = New ADODB.Recordset
feesrec.Open slfc, feescon, adOpenStatic, adLockOptimistic

Set GlobalCon = New ADODB.Connection
GlobalCon.CursorLocation = adUseClient
GlobalCon.Open "Driver={SQL Server};" & _
namesqlserver & _
namesqldatabase

slccg = "select * from CasteGlobal"
Set GlobalCaste = New ADODB.Recordset
GlobalCaste.Open slccg, GlobalCon, adOpenStatic, adLockOptimistic

slcrg = "select * from ReligionGlobal"
Set GlobalReligion = New ADODB.Recordset
GlobalReligion.Open slcrg, GlobalCon, adOpenStatic, adLockOptimistic

slcng = "select * from NationalityGlobal"
Set GlobalNationality = New ADODB.Recordset
GlobalNationality.Open slcng, GlobalCon, adOpenStatic, adLockOptimistic

slcfg = "select * from OccupationGlobalF"
Set GlobalOccupationF = New ADODB.Recordset
GlobalOccupationF.Open slcfg, GlobalCon, adOpenStatic, adLockOptimistic

slcmg = "select * from OccupationGlobalM"
Set GlobalOccupationM = New ADODB.Recordset
GlobalOccupationM.Open slcmg, GlobalCon, adOpenStatic, adLockOptimistic

slsdg = "select * from DesignationGlobal"
Set GlobalDesignation = New ADODB.Recordset
GlobalDesignation.Open slsdg, GlobalCon, adOpenStatic, adLockOptimistic

slsgg = "select * from DepartmentGlobal"
Set GlobalDepartment = New ADODB.Recordset
GlobalDepartment.Open slsgg, GlobalCon, adOpenStatic, adLockOptimistic

Set staffcon = New ADODB.Connection
staffcon.CursorLocation = adUseClient
staffcon.Open "Driver={SQL Server};" & _
namesqlserver & _
namesqldatabase

slcff = "select * from StaffInformation Order by Staff_ID"
Set staffrec = New ADODB.Recordset
staffrec.Open slcff, staffcon, adOpenStatic, adLockOptimistic

Set universitycon = New ADODB.Connection
universitycon.CursorLocation = adUseClient
universitycon.Open "Driver={SQL Server};" & _
namesqlserver & _
namesqldatabase

slcug = "select * from UniversityFeeInformation Order by Reciept_Number"
Set universityrec = New ADODB.Recordset
universityrec.Open slcug, universitycon, adOpenStatic, adLockOptimistic

slsts = "select * from StaffSalaryInformation Order by Reciept_Number"
Set staffsalary = New ADODB.Recordset
staffsalary.Open slsts, staffcon, adOpenStatic, adLockOptimistic

slmge = "Select * from MarkGlobalInformation Order by Serial_Number"
Set globalmark = New ADODB.Recordset
globalmark.Open slmge, GlobalCon, adOpenStatic, adLockOptimistic

slmrp = "Select * from StudentMarkInformation order by Serial_Number"
Set markprint = New ADODB.Recordset
markprint.Open slmrp, GlobalCon, adOpenStatic, adLockOptimistic

globalstuin.Open "Select * from StudentInformation Order by Admission_Number", GlobalCon, adOpenStatic, adLockOptimistic
staffstnen.Open "Select * from StaffInformation Order by Staff_ID", GlobalCon, adOpenStatic, adLockOptimistic
departglo.Open "Select * from DepartmentGlobal order by Serial_Number", GlobalCon, adOpenStatic, adLockOptimistic
stumastin.Open "Select * from StudentInformation Order by Admission_Number", GlobalCon, adOpenStatic, adLockOptimistic
unifeeentrycou.Open "Select * from CourseInformation order by Course_ID", GlobalCon, adOpenStatic, adLockOptimistic
exportdepart.Open "Select * from DepartmentGlobal order by Serial_Number", GlobalCon, adOpenStatic, adLockOptimistic
exportcourse.Open "Select * from CourseInformation order by Course_ID", GlobalCon, adOpenStatic, adLockOptimistic
coursestudent.Open "Select * from CourseInformation order by Course_ID", GlobalCon, adOpenStatic, adLockOptimistic
stuattencourse.Open "Select * from CourseInformation order by Course_ID", coursecon, adOpenStatic, adLockOptimistic

End Function

' Function To Close All Connections
Public Function MainConClose()
If studentcon.State = adStateOpen Then
studentcon.Close
End If
If familycon.State = adStateOpen Then
familycon.Close
End If
If coursecon.State = adStateOpen Then
coursecon.Close
End If
If feescon.State = adStateOpen Then
feescon.Close
End If
If GlobalCon.State = adStateOpen Then
GlobalCon.Close
End If
If staffcon.State = adStateOpen Then
staffcon.Close
End If
If universitycon.State = adStateOpen Then
universitycon.Close
End If
End Function
