Attribute VB_Name = "MainModule"
' This Is The Module Where All The Software Load Functions Are Done
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Main Loading Code
Sub Main()
On Error Resume Next
  
  charactername = GetSetting(App.CompanyName, "Character", "CharName", charactername)
  charfile = GetSetting(App.CompanyName, "Character", "CharFile", charfile)
  If charactername = "" And charfile = "" Then
  charactername = "Merlin"
  charfile = App.Path & "\Characters\merlin.acs"
  End If
  
  If SeCheck.FindSecurityFile = False Then
  Load Frm_CreateAdmin
  Frm_CreateAdmin.Show
  Exit Sub
  End If
  
  Call imsreg
  
  'Get The ServerName and Database Name
  namesqlserver = GetSetting(App.CompanyName, "ServerSQLName", "ServerName", nameserver)
  namesqldatabase = GetSetting(App.CompanyName, "ServerDataBaseName", "DataBaseName", namedatabase)
  
  If splashtime = 0 Then
  splashtime = 8000
  End If
  
  'Load set server name form if the server name is empty
  If namesqlserver <> "" And namesqldatabase <> "" Then
  Call globalload
  Else
  Load Frm_ServerManager
  Frm_ServerManager.Show
  End If
  
  Set copyf = New Scripting.FileSystemObject
  If copyf.FolderExists(App.Path & "\Images") = False Then
  copyf.CreateFolder App.Path & "\Images"
  End If
    
  If copyf.FolderExists(App.Path & "\User") = False Then
  copyf.CreateFolder App.Path & "\User"
  End If
  
  If copyf.FolderExists(App.Path & "\StaffImages") = False Then
  copyf.CreateFolder App.Path & "\StaffImages"
  End If

End Sub

' Function that loads the other forms if the server name is set
Public Function globalload()
On Error Resume Next

' Loads the application flash form
Load Frm_Splash
Frm_Splash.Show
DoEvents

Sleep splashtime
Call MainConEstablish

Unload Frm_Splash
DoEvents
Load MainMenu
Load Frm_LoginMain
MainMenu.Show
MainMenu.Enabled = False
Frm_LoginMain.BackColor = MainMenu.ACPRibbon1.BackColor
Frm_LoginMain.Picture = MainMenu.ACPRibbon1.LoadBackground
Frm_LoginMain.Show

End Function
