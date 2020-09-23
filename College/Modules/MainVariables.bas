Attribute VB_Name = "MainVariables"
' This Is The Module Where All The Public
' Variable Declarations For The Software
' Are Done

Public Const n = vbNewLine
Public Const TIP_FILE = "TIPOFDAY.TXT"
Public mycharacter As IAgentCtlCharacter
Public charactername As String
Public charfile As String

Public Tips As New Collection

Public a As Double

Public userex As Boolean
Public Loaded As Boolean
Public changeas As Boolean

Public Theme As Integer

Public characterop As String
Public key() As String
Public keyuser() As String
Public admcheck As String
Public namesqlserver As String
Public namesqldatabase As String
Public slct As String
Public slcf As String
Public slcc As String
Public slfc As String
Public slccg As String
Public slcrg As String
Public slcng As String
Public slcfg As String
Public slcmg As String
Public slcff As String
Public slsdg As String
Public slsgg As String
Public slcug As String
Public slcas As String
Public slsts As String
Public slmge As String
Public slmrp As String
Public dateexp As String

Public fol As Folder

Public copyf As Scripting.FileSystemObject
Public Folder As New FolderSysObject
Public File As New FileSysObject
Public SeCheck As New SecurityClass
Public crypt As New cSimpleCrypt
Public ShellExecute As New CShellExecute
Public lngResult As EnumShellExecuteErrors
Public FSO As New Scripting.FileSystemObject
Public fldr As Scripting.Folder
Public f As Scripting.File

Public wininfo As New GetWindowsInformation
Public memsta As New CMemoryStatus
Public info As New COSInfo
Public mousein As New CMouseInfo
Public keyboardin As New CKeyboardInfo
Public cmpinfo As New CComputerInfo
Public keyvalc As New CClassValidate

Public ShowAtStartup As Long
Public splashtime As Long
Public ret As Long
Public CurrentTip As Long

