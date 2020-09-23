Attribute VB_Name = "FunctionModule"
' API Function Declarations For Setting Windows Position
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' Constants Needed For API
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const LWA_COLORKEY = &H3
Public Const LWA_ALPHA = &H3
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const SWP_NOOWNERZORDER = &H200

' Function To Validate E-Mail Address
Public Function ValidateEmail(n As String) As Boolean
Dim j As Integer
If n = "" Then GoTo ByPass1                  ' allow zero length

For j = 1 To Len(n)
   If Mid$(n, j, 1) = "@" Then               ' check if address has @ sign
      If Mid$(n, j + 1, 1) = "." Then        ' if @ sign is followed by period(.)
         ValidateEmail = False
         Exit For
      End If
      If j = 1 Then
           ValidateEmail = False
         Else
           If ValidateExt(j, n) = True Then  ' validate extention
ByPass1:
              ValidateEmail = True
                 Else
              ValidateEmail = False
           End If
      End If
      Exit For
         Else
      ValidateEmail = False
   End If
Next j

End Function

' Part Of E-Mail Checking Function For Checking Extensions
Public Function ValidateExt(Start As Integer, n As String) As Boolean
Dim j As Integer
Dim ext As String
Dim depot As String

For j = Start To Len(n)
    depot = Mid$(n, j, 1)
    If depot = "." Then
      ext = depot
      ext = ext + Mid$(n, j + 1, Len(n) - j)
        If ext = ".com" Or ext = ".org" Or ext = ".net" Or ext = ".co.in" Or ext = ".com.ph" Then
           ValidateExt = True
               Else
           ValidateExt = False
        End If
        Exit For
End If
Next j
End Function

' Function To Backup Existing Student Pictures
Public Sub backuppicture()
On Error Resume Next
Set copyf = New Scripting.FileSystemObject
If copyf.FolderExists(App.Path & "\Images") Then
Dim FolPath As String
FolPath = copyf.GetSpecialFolder(SystemFolder) & "\"
copyf.CopyFolder App.Path & "\Images", FolPath, True
End If
End Sub

' Function For Checking Admin Password And UserName
Public Function readadminpass()
key = Split(SeCheck.ReadSecurityFile, "//")
If Frm_LoginMain.txtname.Text = crypt.DeCode(key(0)) And Frm_LoginMain.txtpass.Text = crypt.DeCode(key(1)) Then
   MainMenu.StatusBar1.Panels(6).Text = "Login As :" & " " & crypt.DeCode(key(0)) & " " & Format(Time)
   MainMenu.Enabled = True
   Unload Frm_LoginMain
   Frm_UserMessage.Resize
   Frm_UserMessage.Show
   Frm_UserMessage.Label1.Caption = "Login As :" & " " & crypt.DeCode(key(0))
   Frm_UserMessage.Label2.Caption = Format(Time)
Else
   MsgBox "Sorry, Incompatible Username and Password.", vbCritical, "Error Occured"
   Frm_LoginMain.txtpass.SetFocus
End If
End Function

' Function For Checking UserName And Password
Public Sub readuserpass()
If copyf.FileExists(App.Path & "\User\" & Frm_LoginMain.txtusername.Text & ".SecurityFile") Then
userex = True
keyuser = Split(SeCheck.ReadSecurityFileUser(App.Path & "\User\" & Frm_LoginMain.txtusername.Text), "//")
If Frm_LoginMain.txtusername.Text = crypt.DeCode(keyuser(0)) And Frm_LoginMain.txtpass.Text = crypt.DeCode(keyuser(1)) Then
   MainMenu.StatusBar1.Panels(6).Text = "Login As :" & " " & crypt.DeCode(keyuser(0)) & " " & Format(Time)
   MainMenu.Enabled = True
   Unload Frm_LoginMain
   Frm_UserMessage.Resize
   Frm_UserMessage.Show
   Frm_UserMessage.Label1.Caption = "Login As :" & " " & crypt.DeCode(keyuser(0))
   Frm_UserMessage.Label2.Caption = Format(Time)
Else
   MsgBox "Sorry, Incompatible Username and Password.", vbCritical, "Error Occured"
   Frm_LoginMain.txtpass.SetFocus
End If
Else
MsgBox "No Such User Exist, Login Using Admin and" & vbCrLf & "Create New User", vbInformation, "No User Exist"
Frm_LoginMain.txtusername.Text = ""
Frm_LoginMain.txtpass.Text = ""
userex = False
Exit Sub
End If
End Sub

' Function To Show One Form Over All The Other Forms
Public Sub AlwaysOnTop(formname As Form, SetOnTop As Boolean)
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If

    SetWindowPos formname.hwnd, lFlag, _
    formname.Left / Screen.TwipsPerPixelX, _
    formname.Top / Screen.TwipsPerPixelY, _
    formname.Width / Screen.TwipsPerPixelX, _
    formname.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

' Function To Export Data To Excel
Public Sub exporttoexcel(ByRef CommDial1 As CommonDialog, ByRef Recordset1 As ADODB.Recordset)
On Error GoTo Err:
    Dim i, j As Integer
    Dim str1 As String
    i = 0
    j = 0
    
    With CommDial1
    
        .CancelError = True
        .Filter = "Excel File Format (*.xls)|*.xls;|"
        .ShowSave
         str1 = .FileName
         
    End With
    
    Dim createExcel As New Excel.Application
    Dim Wbook As Excel.Workbook
    Dim Wsheet As Excel.Worksheet
    Set Wbook = createExcel.Workbooks.Add
    Set Wsheet = Wbook.Worksheets.Add
       
    For i = 0 To Recordset1.Fields.Count - 1
        Wsheet.Cells(1, i + 1).Value = Recordset1.Fields(i).Name
    Next i
    
    If (Recordset1.RecordCount > 0) Then
        Recordset1.Movefirst
        For i = 0 To Recordset1.RecordCount - 1
            For j = 0 To Recordset1.Fields.Count - 1
                Wsheet.Cells(i + 2, j + 1).Value = Recordset1(j).Value
            Next j
            Recordset1.Movenext
        Next i
    End If
    
    Wbook.SaveAs str1
    Wbook.Close True
    
    Set createExcel = Nothing
    Set Wbook = Nothing
    Set Wsheet = Nothing
    Set Wbook = createExcel.Workbooks.Open(str1)
    
    createExcel.Visible = True
    Exit Sub
    
    
Err:
 Select Case Err.Number
    Case 32755
             MsgBox "Press Cacel Button", vbInformation, "Cancelled Export"
    Case 1004
             MsgBox "OverWrite Cancel", vbInformation, "Over Write"
             Wbook.Close False
          
    Case Else
         MsgBox Err.Number & " " & Err.Description, vbInformation, "Error Occured"
         
 End Select

End Sub

' Function To Check Whether Family Information Is Entered Or Not
Public Function checkfamilydatavalid() As Boolean
On Error Resume Next
Dim checkdataf As ADODB.Recordset
Dim checkdatafstr As String
Set checkdataf = New ADODB.Recordset
If checkdataf.State = adStateOpen Then checkdataf.Close
checkdatafstr = "Select * from FamilyInformation where Admission_Number = '" & CDbl(Frm_StudentEntry.adnumber.Text) & "'"
checkdataf.Open checkdatafstr, GlobalCon, adOpenStatic, adLockOptimistic
If checkdataf.RecordCount <> 0 Then
checkfamilydatavalid = True
checkdataf.Close
Else
checkfamilydatavalid = False
checkdataf.Close
End If
End Function

' Transparent Making area
Public Sub Trans(frmname As Form)
On Error Resume Next
ret = GetWindowLong(frmname.hwnd, GWL_EXSTYLE)
ret = ret Or WS_EX_LAYERED
SetWindowLong frmname.hwnd, GWL_EXSTYLE, ret
SetLayeredWindowAttributes frmname.hwnd, 0, 200, LWA_ALPHA
End Sub

' Read in the tips file and display a tip at random.
Public Function loadtipsall(labelname As label)
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        labelname.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If
End Function

' Load Tips To Memory
Public Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

' Move To Next Tip
Public Sub DoNextTip()
' Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd) + 1)

' Show it.
    Call DisplayCurrentTip(Frm_TipInformation.lblTipText)
End Sub

' Display Current
Public Sub DisplayCurrentTip(labelname As label)
    If Tips.Count > 0 Then
        labelname.Caption = Tips.Item(CurrentTip)
    End If
End Sub

' Move To Next Tip
Public Sub DoNextTipSide()
' Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd) + 1)

' Show it.
    Call DisplayCurrentTip(Frm_Sidebar.lblTipText)
End Sub

' Check For Theme And Chenge Label Forecolor With That
' For The Side Bar Only
Public Function checkthemeside()
If MainMenu.ACPRibbon1.Theme = 0 Then
Frm_Sidebar.Label1.ForeColor = &HFFFFFF
Frm_Sidebar.Label16.ForeColor = &HFFFFFF
ElseIf MainMenu.ACPRibbon1.Theme = 1 Then
Frm_Sidebar.Label1.ForeColor = &H80&
Frm_Sidebar.Label16.ForeColor = &H80&
ElseIf MainMenu.ACPRibbon1.Theme = 2 Then
Frm_Sidebar.Label1.ForeColor = &H80&
Frm_Sidebar.Label16.ForeColor = &H80&
End If
End Function

' Check For Registration
Public Function imsreg()
Dim dateexp As String
dateexp = GetSetting("VRJ Soft", "Exp Date", "RegExp", dateexp)
Date = GetSetting("VRJ Soft", "Install Date", "RegInstall", Date)
If dateexp = "" Then
MsgBox "Cannot Find Registration Expire Date", vbInformation, "Reg Error"
End
Else
If DateDiff("d", dateexp, Date) >= 0 Then
MsgBox "Your Key Is Expired, Application Will Exit", vbInformation, "Key Expired"
End
End If
End If
End Function

' Show Message Of Assistant
Public Sub Merlin(Optional ByVal Msg As String, Optional ByVal Animation As String = "Explain")
    On Error Resume Next
    If mycharacter.Visible Then
        mycharacter.StopAll
        mycharacter.Play Animation
        
        If Not Msg = "" Then mycharacter.Speak Msg
    End If
End Sub

' Disable Help Assistant
Public Function disablemerlin()
On Error Resume Next
characterop = GetSetting(App.CompanyName, "Character", "MyAssist", characterop)
If characterop = "Enabled" Then
characterop = "Disabled"
MainMenu.MyAgent.Characters.Unload charactername
MsgBox "Help Assistant Is Disabled", vbInformation, "Help Assistant"
SaveSetting App.CompanyName, "Character", "MyAssist", characterop
ElseIf characterop = "Disabled" Then
characterop = "Enabled"
MainMenu.MyAgent.Characters.Load charactername
MsgBox "Help Assistant Is Enabled", vbInformation, "Help Assistant"
SaveSetting App.CompanyName, "Character", "MyAssist", characterop
Set mycharacter = MainMenu.MyAgent.Characters(charactername)
mycharacter.SoundEffectsOn = True
mycharacter.Show
mycharacter.MoveTo 850, 550
End If
End Function

Public Function browseimage(diagname As CommonDialog, imagename As Image, Optional ByVal textname As TextBox)
Dim strfn As String
diagname.CancelError = True   ' Catch Cancel Error
On Error GoTo ErrHandler   ' Catch On Any Other Error
   
diagname.DefaultExt = "jpg" ' Default Picture Extension
diagname.DialogTitle = "Select Picture Files" ' Common Dialog Title
diagname.Filter = "Picture File JPEG Format (*.jpg,*.jpe,*.jpeg)|*.jpg;*.jpe;*.jpeg|Bitmap Files (*.bmp)|*.bmp"  ' Only Show Included Extensions
diagname.ShowOpen ' Show Open Dialog
strfn = diagname.FileName ' Store Image Name To String
textname.Text = strfn
imagename.Picture = LoadPicture(strfn) ' Load Picture To Image Box
Exit Function ' Exit Picture Open Code
ErrHandler:
'User pressed the Cancel button
If Err.Number = 32755 Then
MsgBox "You Have Cancelled The Picture" & vbCrLf & "         File Selection", vbInformation + vbOKOnly, "Action Cancelled"
Else
MsgBox Err.Description, vbInformation, "Error Information"
End If
End Function
