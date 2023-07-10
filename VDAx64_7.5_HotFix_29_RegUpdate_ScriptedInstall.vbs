Option Explicit

Dim sScriptVer : sScriptVer = "1.6"

Sub SetCustomizations()
	sPkgName = "Citrix_VirtualDeliveryAgentx64_7.5_HotFix_29_RegUpdate" ' Application Friendly Name
	sProductVendor = "Citrix Systems, Inc" ' Product vendor
	sProductName = "VirtualDeliveryAgent"  ' Product name
	sProductVersion	= "7.5" ' Product version
	sProductCode = "Citrix Virtual Desktop Agent" ' Product code
	sInstallFile = "VDAx64_7.5\x64\XenDesktop Setup\XenDesktopVdaSetup.exe" ' Install file name
	sInstallFolderPath = sProgFiles64 & "\Citrix\XenDesktopVdaSetup" ' Install folder path
	sProductFilePath = sProgFiles64 & "\Citrix\XenDesktopVdaSetup\XenDesktopVdaSetup.exe" ' Main application executable
	sProductFileVersion = "7.5.0.4529" ' Main application executable product 
	sWestController = "(DC1CTXXD01V.corp.firstrepublic.com DC1CTXXD02V.corp.firstrepublic.com) (DC3CTXXD01V.corp.firstrepublic.com DC3CTXXD02V.corp.firstrepublic.com)"
	sEastController = "(DC2CTXXD01V.corp.firstrepublic.com DC2CTXXD02V.corp.firstrepublic.com) (NYCTXXD01V.corp.firstrepublic.com NYCTXXD02V.corp.firstrepublic.com)"
	sCompanyName = "FRB" ' Variable used to create log folder & registry key path
    
    sRegAuditPath = sCompanyName & "\Applications"
    sLogPath = sSystemDrive & "\" & sCompanyName & "\Logs"

    bFileOverwrite = True ' Set to True to overwrite existing files if they exist
End Sub

' SetCustomizations
Dim sPkgName,sProductVendor,sProductName,sProductVersion,sProductCode,sProductFilePath_PreReq,sWestController,sEastController
Dim sInstallFile,sInstallFolderPath,sProductFilePath,sProductFileVersion
Dim sCompanyName,sRegAuditPath,sLogPath,bFileOverwrite

' Init
Dim oShell,oShellApp,oFSO,oNetwork
Dim sProcessArchitectureX86,sProcessArchitectureW6432
Dim iOSArch,iScriptArch
Dim sSysDir32,sSysDir64
Dim sProgFiles32,sProgFiles64
Dim sRegKeyRoot32,sRegKeyRoot64
Dim sAllUsersStartMenu,sAllUsersStartPrograms
Dim sAllUsersDesktop,sAllUsersAppData
Dim sWinDir,sSystemDrive,sTempDir,sScriptDir
Dim sUserProfile,bUninstall,Quotes,sInstallDate
Dim sTempGuid

Const HKEY_CLASSES_ROOT =&H80000000
Const HKEY_CURRENT_USER =&H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003

' Logging
Dim sVerb,sLogText

' GetRegPath
Dim sRegPath,sTattooPath

'========================================================================
' Main Script Logic
'========================================================================

Init
SetCustomizations
Dim oReg
Set oReg = GetObject("winmgmts://./root/default:StdRegProv")
Dim sKeyPath,sValueName,sValue,sTSKeyPath,sTSValueName,sTSValue,sTSName,sTSNameData
sKeyPath = "SYSTEM\CurrentControlSet\Control\TimeZoneInformation"
sValueName = "TimeZoneKeyName"
oReg.GetStringValue HKEY_LOCAL_MACHINE,sKeyPath,sValueName,sValue
sTSKeyPath = "SOFTWARE\FRB\OS Deployment"
sTSName = "Task Sequence Name"
oReg.GetStringValue HKEY_LOCAL_MACHINE,sTSKeyPath,sTSName,sTSNameData	
sTSValueName = "Task Sequence Version"
oReg.GetStringValue HKEY_LOCAL_MACHINE,sTSKeyPath,sTSValueName,sTSValue
BeginLog
If bUninstall Then
	If Not IsProductInstalled Then
		LogItem "Deployment script will now exit!",True,False
		QuitScript(0)
	Else
		' Uninstall Logic
        GetRegPath
        LogItem "Deployment script will now proceed with: " & sVerb & "!",True,False
        CopyFile sAllUsersAppData & "\Citrix\picaser\picaser.sys", sWinDir & "\system32\DRIVERS\picaser.sys" 
        Uninstall sSysDir64 & "\msiexec.exe","/uninstall {A4EA9750-2FDE-5A09-91CD-71A4546F304A} /package {BFFC42EE-9788-4E25-8CD2-8C3E703698F0} /qb-! REBOOT=ReallySuppress /l*xv " & Quotes & sLogPath & "\" & sPkgName & "_(Patch)_Uninstall.log" & Quotes 
		If Not (IsRegKeyExist("HKLM","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ICAWS750WX64029")) Then
			LogItem "Validate Uninstall: Able to validate non-existence of Target Reg Key at: HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ICAWS750WX64029",True,False
			LogItem "Re-applying ORIGINAL value for - ListOfDDCs - Registry Key.",True,False
			LogItem "Time Zone for this system is: " & sValue,True,False
			If sValue = "Eastern Standard Time" Then
				LogItem "Applying East Coast Controller parameter for - ListOfDDCs.",True,False
				SetRegVal "HKLM","SOFTWARE\Citrix\VirtualDesktopAgent","ListOfDDCs","REG_SZ",sEastController
			Else
				LogItem "Applying West Coast Controller parameter for - ListOfDDCs.",True,False
				SetRegVal "HKLM","SOFTWARE\Citrix\VirtualDesktopAgent","ListOfDDCs","REG_SZ",sWestController
			End If
		Else
			LogItem "Validate Uninstall: Target Reg Key still exits at: HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ICAWS750WX64029",True,False
			LogItem "Validate Uninstall: Hotfix - ICAWS750WX64029 - was not uninstalled successfully.",True,False
			QuitScript(1603)
		End If
		'RemoveTattoo
		DelRegValName "HKLM","SOFTWARE\Policies\Citrix","EnforceAutoLogon"
		DeleteKey "HKLM","SOFTWARE\FRB\Applications\" & sPkgName
		DeleteFolder sAllUsersAppData & "\Citrix\picaser"
		LogItem "Deployment script will now exit!",True,False
	End If
Else
	If IsProductInstalled Then
		LogItem "Deployment script will now exit!",True,False
		QuitScript(0)
	Else
		'Install Logic
		LogItem "Deployment script will now proceed with: " & sVerb & "!",True,False	
		If IsRegKeyExist("HKLM","SOFTWARE\FRB\Applications\Citrix_VirtualDeliveryAgentx64_7.5") Then
			LogItem "Dependency Package has found - Citrix_VirtualDeliveryAgentx64_7.5",True,False
			oShell.Run "robocopy " & Quotes & sProgFiles64 & "\Citrix\ICAService\DeviceDrivers\picaser" & Quotes & " " & Quotes & sAllUsersAppData & "\Citrix\picaser" & Quotes & " /E",0,True
			LogItem "Proceeding with installing Hotfix (ICAWS750WX64029)",True,False
			InstallMSP "Hotfix_29\ICAWS750WX64029.msp","/qb-! REINSTALL=ALL REINSTALLMODE=omus REBOOT=ReallySuppress /l*xv " & Quotes & sLogPath & "\" & sPkgName & "_(Patch)_Install.log" & Quotes
			LogItem "Time Zone for this system is: " & sValue,True,False
			If sValue = "Eastern Standard Time" Then
				LogItem "Applying East Coast Controller parameter for - ListOfDDCs.",True,False
				SetRegVal "HKLM","SOFTWARE\Citrix\VirtualDesktopAgent","ListOfDDCs","REG_SZ","DC2CTXXD01V.corp.firstrepublic.com DC2CTXXD02V.corp.firstrepublic.com"
			Else
				LogItem "Applying West Coast Controller parameter for - ListOfDDCs.",True,False
				SetRegVal "HKLM","SOFTWARE\Citrix\VirtualDesktopAgent","ListOfDDCs","REG_SZ","DC1CTXXD01V.corp.firstrepublic.com DC1CTXXD02V.corp.firstrepublic.com"
			End If			
		Else
			LogItem "Dependency Package has NOT found - Citrix_VirtualDeliveryAgentx64_7.5",True,False
			LogItem "Install Script will Terminiate, without applying any changes to your system.",True,False
			QuitScript(0)
		End If	
		   					
		'ValidateInstall	
		If IsRegKeyExist("HKLM","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ICAWS750WX64029") Then
			LogItem "Validate Install: Able to validate Target Reg Key at: HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ICAWS750WX64029",True,False
		Else
			LogItem "Validate Install: Unable to validate Target Reg Key.",True,False
			QuitScript(1603)
		End If 		
		GetRegPath				
		LogItem "Applying Add/Remove Programs Customizations.",True,False
		'ARPCustomization
		'SetRegVal "HKLM","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ICAWS750WX64029","Comments","REG_SZ","Hotfix 29 for Citrix VDA 7.5"
		SetRegVal "HKLM","SOFTWARE\Policies\Citrix","EnforceAutoLogon","REG_DWORD","0"
		DelRegValName "HKLM","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ICAWS750WX64029","HelpLink"
		DelRegValName "HKLM","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ICAWS750WX64029","URLInfoAbout"
		DelRegValName "HKLM","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ICAWS750WX64029","URLUpdateInfo"		
		LogItem "Creating Audit Registry Keys!",True,False
		TattooRegistry
		LogItem "Deployment script will now exit!",True,False
		If FileExist (sSysDir64 & "\Tasks\Del_SCCMCache_CitrixVDA") Then
			Uninstall sSysDir64 & "\schtasks.exe", "/delete /f /tn " & Quotes & "Del_SCCMCache_CitrixVDA" & Quotes
		End If
		If FileExist (sAllUsersAppData & "\Citrix\Del_SCCMCache_CitrixVDA.xml") Then
			LogItem "Found file - " & sAllUsersAppData & "\Citrix\Del_SCCMCache_CitrixVDA.xml",True,False 
			oShell.Run "cmd /c del /F /Q " & Quotes & sAllUsersAppData & "\Citrix\Del_SCCMCache_CitrixVDA.xml" & Quotes,0,True
			LogItem "Deleted file - " & sAllUsersAppData & "\Citrix\Del_SCCMCache_CitrixVDA.xml",True,False
		End If
		'TattooInstallFolder(sInstallFolderPath & "" & sPkgName & "")
	End If
End If

'Success (No Reboot)
'QuitScript(0)

'Success (No Reboot) 
'QuitScript(1707)

'Soft Reboot 
QuitScript(3010)

'Hard Reboot
'QuitScript(1641)

'========================================================================
' Do not edit bellow this section
'========================================================================

Sub Init
    Set oShell = CreateObject("WScript.Shell")
    Set oShellApp = CreateObject("Shell.Application")
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oNetwork = CreateObject("WScript.Network")

	' Determine OS Architecture
	sProcessArchitectureX86 = oShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
	If oShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%") = "%PROCESSOR_ARCHITEW6432%" Then
		sProcessArchitectureW6432 = "Not Defined"
	End If

	If sProcessArchitectureX86 = "x86" And sProcessArchitectureW6432 = "Not Defined" Then
		' Windows 32-bit
		iOSArch = 32
		iScriptArch = 32
		sSysDir32 = oShellApp.NameSpace(37).Self.Path
		sProgFiles32 = oShellApp.NameSpace(38).Self.Path
		sRegKeyRoot64 = "SOFTWARE"
		sRegKeyRoot32 = "SOFTWARE"
	Else
		' Windows 64-bit
		iOSArch = 64
		iScriptArch = 64
		sSysDir64 = oShellApp.NameSpace(37).Self.Path
		sProgFiles64 = oShellApp.NameSpace(38).Self.Path
		sSysDir32 = oShellApp.NameSpace(41).Self.Path
		sProgFiles32 = oShellApp.NameSpace(42).Self.Path
		sRegKeyRoot32 = "SOFTWARE\Wow6432Node"
		sRegKeyRoot64 = "SOFTWARE"
	End If

	' %ProgramData%\Microsoft\Windows\Start Menu
	sAllUsersStartMenu = oShellApp.NameSpace(22).Self.Path
		
	' %ProgramData%\Microsoft\Windows\Start Menu\Programs
	sAllUsersStartPrograms = oShellApp.NameSpace(23).Self.Path
	
	' %SystemDrive%\Users\Public\Desktop
	sAllUsersDesktop = oShellApp.NameSpace(25).Self.Path
	
	' %SYSTEMDRIVE%\ProgramData
	sAllUsersAppData = oShellApp.NameSpace(35).Self.Path
	
	' %WINDIR%
	sWinDir = oShellApp.NameSpace(36).Self.Path
	
	' %SYSTEMDRIVE%
	sSystemDrive = oShell.ExpandEnvironmentStrings("%SystemDrive%")
			
	' %WINDIR%\Temp - System Account
	sTempDir = oShell.ExpandEnvironmentStrings("%TEMP%")
		
	' Get script directory without trailing slash
	sScriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\")-1)
	
	'If root of drive, strip trailing backslash
	If Len(sScriptDir) = 3 Then sScriptDir = Left(sScriptDir, 2)
	
	' %SYSTEMDRIVE%\Users\%USERNAME%
	sUserProfile = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
	
	' Check if /uninstall was passed to the script
	bUninstall = False

	If WScript.Arguments.Named.Exists("UNINSTALL") Then bUninstall = True
		
	' used to encapsute paths
	Quotes = Chr(34)
		
	' Convert Now() to String
	sInstallDate = CStr(Now())
	
	' Generate GUID
	sTempGuid = CreateGuid

End Sub

'========================================================================
' Logging
'========================================================================

Sub BeginLog
	If bUninstall Then
		sVerb = "Uninstall"
	Else
		sVerb = "Install"
	End If
	
	' Create log folder path if needed
	If Not oFSO.FolderExists(sLogPath) Then
		CreateFolderIfNeeded(sLogPath)
	End If

	LogItem "******************** Begin " & sVerb & " ********************", True, False
	LogItem "Deployment Script Version: " & sScriptVer,True,False
	LogItem "OS Architecture: " & iOSArch & " / Script Architecture: " & iScriptArch, True, False
	LogItem "Domain Information: " & oNetwork.UserDomain, True, False
	LogItem "Computer Name: " & oNetwork.ComputerName, True, False
	LogItem "Task Sequence Name: " & sTSNameData, True, False
	LogItem "Task Sequence Version: " & sTSValue, True, False
	LogItem "User Executing Script: " & oNetwork.UserName, True, False
	LogItem "Product Vendor: " & sProductVendor, True, False
	LogItem "Product Name: " & sProductName, True, False
	LogItem "Product Version: " & sProductVersion, True, False
	LogItem "Product Code: " & sProductCode, True, False
	LogItem "Date: " & Now(), True, False
End Sub

Sub QuitScript(iRetVal)
	LogItem "Exiting Script (" & iRetVal & ")", True, True
	LogItem "******************** End " & sVerb & " ********************", True, False
	sLogText = Trim(sLogText)
	Select Case iRetVal
	Case 0, 3010
		oShell.Logevent 4, sLogText
	Case Else
		oShell.LogEvent 1, sLogText
	End Select
	WScript.Quit(iRetVal)
End Sub

Sub LogItem(sMessage, bLogFile, bEventLog)
	If bEventLog Then
		sLogText = sLogText & sMessage & vbCrLf
	End If

	If bLogFile Then 
		On Error Resume Next
		Dim tsLog
		Set tsLog = oFSO.OpenTextFile(sLogPath & "\" & sPkgName & "_(Script).log", 8, True)
		tsLog.WriteLine Now() & " - " & sMessage
		On Error GoTo 0
	End If
End Sub

'========================================================================
' Install & Uninstall Routines
'========================================================================

Sub InstallMSI(sFileName,sParameters)
	Dim sSubroutine
	sSubroutine = "Install MSI: "

	LogItem sSubroutine & "Started",True,False

	Dim sFilePath
	sFilePath = sScriptDir & "\" & sFileName
	
	Dim sCmd
	sCmd = Quotes & sWinDir & "\System32\MsiExec.exe" & Quotes & " /I " & Quotes & sFilePath & Quotes & " " & sParameters
	
	LogItem sSubroutine & "About to execute " & sCmd,True,False
	
	Dim iRetVal
	iRetVal = oShell.Run(sCmd,0,True)
	
	Select Case iRetVal
	Case 0, 3010
		'Success!
		LogItem sSubroutine & "Install was successful.  (Returned " & iRetVal & ")",True,True
	Case 1618
		LogItem sSubroutine & "Install was uncessessful (Returned 1618 -- Another installation is in progress)", True,True
		QuitScript(1618)
	Case Else
		'Failure
		LogItem sSubroutine & "Install was unsuccessful.  (Returned " & iRetVal & ")",True,True
		QuitScript(iRetVal)
	End Select
	LogItem sSubroutine & "Process finished. (Returned " & iRetVal & ")",True,True
	LogItem sSubroutine & "Finished",True,False
End Sub

Sub InstallMSP(sFileName,sParameters)
	Dim sSubroutine
	sSubroutine = "Install MSP: " 

	LogItem sSubroutine & "Started",True,False

	Dim sFilePath
	sFilePath = sScriptDir & "\" & sFileName

	Dim sCmd
	sCmd = Quotes & sWinDir & "\System32\MsiExec.exe" & Quotes &  " /p " & Quotes & sFilePath & Quotes & " " & sParameters

   	LogItem sSubroutine & "About to execute " & sCmd,True,False
   	
   	Dim iRetVal
	iRetVal = oShell.Run(sCmd,0,True)
	Select Case iRetVal

	Case 0, 3010
		'Success!
		LogItem sSubroutine & "Install was successful.  (Returned " & iRetVal & ")",True,True
	Case 1618
		LogItem sSubroutine & "Install was uncessessful (Returned 1618 -- Another installation is in progress)", True,True
		QuitScript(1618)
	Case Else
		'Failure
		LogItem sSubroutine & "Install was unsuccessful.  (Returned " & iRetVal & ")",True,True
		QuitScript(iRetVal)
	End Select
	LogItem sSubroutine & "Process finished. (Returned " & iRetVal & ")",True,True
	LogItem sSubroutine & "Finished",True,False
End Sub

Sub UninstallMSI(sGUID,sParameters)
	Dim sSubroutine
	sSubroutine = "Uninstall MSI: "
	
	LogItem sSubroutine & "Started",True,False
	
	Dim sCmd,iRetVal
	sCmd = Quotes & sWinDir & "\System32\MsiExec.exe" & Quotes & " /X " & sGUID & " " & sParameters
   	
   	LogItem sSubroutine & "About to execute " & sCmd, True, False

	iRetVal = oShell.Run(sCmd,0,True)
	Select Case iRetVal
		Case 0, 3010
			'Success!
			LogItem sSubroutine & "Uninstall was successful.  (Returned " & iRetVal & ")",True,True
		Case 1618
			LogItem sSubroutine & "Uninstall was uncessessful (Returned 1618 -- Another installation is in progress)", True,True
			QuitScript(1618)
	Case Else
	
		'Failure
		LogItem sSubroutine & "Uninstall was unsuccessful.  (Returned " & iRetVal & ")",True,True
		QuitScript(iRetVal)
	End Select
	
	LogItem sSubroutine & "Process finished. (Returned " & iRetVal & ")", True, False
	LogItem sSubroutine & "Finished",True,False
End Sub

Sub Install(sFilePath,sParameters)
	Dim sSubroutine
	sSubroutine = "Install Setup: " 
	
	LogItem sSubroutine & "Started",True,False
		
	Dim sCmd
	sCmd =  Quotes & sFilePath & Quotes & " " & sParameters
   	
   	LogItem sSubroutine & "Running: " & sCmd,True,False
	
	Dim iRetVal
	iRetVal = oShell.Run(sCmd,0,True)
	LogItem sSubroutine & "Process finished. (Returned " & iRetVal & ")",True,True
	LogItem sSubroutine & "Finished",True,False
End Sub

Sub Uninstall(sFilePath,sParameters)
	Dim sSubroutine
	sSubroutine = "Uninstall Setup: " 

	LogItem sSubroutine & "Started",True,False

	Dim sCmd
	sCmd = Quotes & sFilePath & Quotes & " " & sParameters
	
   	LogItem sSubroutine & "Running: " & sCmd,True,False
   	
	Dim iRetVal
	iRetVal = oShell.Run(sCmd,0,True)
	
	LogItem sSubroutine & "Process finished. (Returned " & iRetVal & ")",True,False
	LogItem sSubroutine & "Finished",True,False
End Sub

'========================================================================
' Registry Routines
'========================================================================

Function IsRegKeyExist(sRootKey,sSubKey)
	Dim sKeyName,iRetVal
	sKeyName = sRootKey & "\" & sSubKey
    iRetVal = oShell.Run("REG QUERY" & " " & Quotes &  sKeyName & Quotes,0,True)
    If iRetVal <> 0 Then
        IsRegKeyExist = False
    Else
        IsRegKeyExist = True
    End If
End Function

Function IsRegValNameExist(sRootKey,sSubKey,sValueName)
	Dim sKeyName,iRetVal
	sKeyName = sRootKey & "\" & sSubKey
    iRetVal = oShell.Run("REG QUERY " & Quotes & sKeyName & Quotes & " /v " & Quotes & sValueName & Quotes,0,True)
    If iRetVal <> 0 Then
        IsRegValNameExist = False
    Else
        IsRegValNameExist = True
    End If
End Function

Sub CreateKey(sRootKey,sSubKey)
	Dim sKeyName,iRetVal
	sKeyName = sRootKey & "\" & sSubKey
    If Not IsRegKeyExist(sRootKey,sSubKey) Then
        LogItem "About to create: " & Quotes & sKeyName & Quotes,True,False
        iRetVal = oShell.Run("REG ADD " & Quotes &  sKeyName & Quotes & " /f",0,True)
        If iRetVal <> 0 Then
            LogItem Quotes & sKeyName & Quotes & " was not created",True,False
        Else
            LogItem Quotes & sKeyName & Quotes & " has been created",True,False
        End If
    Else
        LogItem Quotes & sKeyName & Quotes & " already exists",True,False
    End If
End Sub

Sub DeleteKey(sRootKey,sSubKey)
	Dim sKeyName,iRetVal
	sKeyName = sRootKey & "\" & sSubKey
    If IsRegKeyExist(sRootKey,sSubKey) Then
        LogItem "About to delete: " & Quotes & sKeyName & Quotes,True,False
        iRetVal = oShell.Run("REG DELETE " & Quotes &  sKeyName & Quotes & " /f",0,True)
        If iRetVal <> 0 Then
            LogItem Quotes & sKeyName & Quotes & " was not deleted",True,False
        Else
            LogItem Quotes & sKeyName & Quotes & " has been deleted",True,False
        End If
    Else
        LogItem Quotes & sKeyName & Quotes & " does not exist.",True,False
    End If
End Sub

Sub SetRegVal(sRootKey,sSubKey,sValueName,sDataType,sValue)
    Dim sKeyName
	sKeyName = sRootKey & "\" & sSubKey
    
    If Not IsRegKeyExist(sRootKey,sSubKey) Then
        CreateKey sRootKey,sSubKey
    End If
    
    If Right(sValue,1) = "\" Then
    	sValue = sValue & "\"
    End If

    Dim iRetVal
    iRetVal = oShell.Run("REG ADD " & Quotes & sKeyName & Quotes & " /v " & Quotes & sValueName & Quotes & " /t " & sDataType & " /d " & Quotes &  sValue & Quotes & " /f",0,True)
    If iRetVal <> 0 Then
        LogItem "The value of " & Quotes & sValueName & Quotes & " under " & sKeyName & " was not set to " & Quotes & sValue & Quotes & " as " &  sDataType,True,False
        LogItem "The process returned: " & iRetVal,True,False
    Else
        LogItem "The value of " & Quotes & sValueName & Quotes & " under " & sKeyName & " was set to " & Quotes & sValue & Quotes & " as " &  sDataType,True,False
    End If
End Sub

Sub DelRegValName(sRootKey,sSubKey,sValueName)
	Dim iRetVal,sKeyName
	sKeyName = sRootKey & "\" & sSubKey
    If IsRegValNameExist(sRootKey,sSubKey,sValueName) Then
        LogItem "About to delete " & Quotes & sValueName & Quotes & " under " & Quotes & sRootKey & "\" & sSubKey & Quotes,True,False
        iRetVal = oShell.Run("REG DELETE " & Quotes &  sKeyName & Quotes & " /v " & sValueName & " /f",0,True)
        If iRetVal <> 0 Then
            LogItem Quotes & sKeyName & "\" & sValueName & Quotes & " was not deleted",True,False
        Else
            LogItem Quotes & sKeyName & "\" & sValueName & Quotes & " has been deleted",True,False
        End If
    Else
        LogItem Quotes & sKeyName & "\" & sValueName & Quotes & " does not exist.",True,False
    End If
End Sub

'========================================================================
' Process and Serices Manipulation Workers
'========================================================================

Sub ProcessStatus(sProcess)
	LogItem "Process Status Check: Started",True,False
	Dim bRunning, oProcesses, oProcess
	Dim oWMI
	
	Set oWMI = GetObject("winmgmts://./root/cimv2")
	Set oProcesses = oWMI.ExecQuery("Select * From Win32_Process",,48)
	
	On Error Resume Next
	For Each oProcess in oProcesses
		If sProcess = oProcess.Name Then
			LogItem "Process Status Check: Found " & sProcess & " is running.",True,False
			LogItem "<----> System Reboot required <---->. ",True,False
			LogItem "Process Status Check: Finished",True,False
			QuitScript(3010)	
		End If
	Next
	On Error GoTo 0
	LogItem "Process Status Check: Finished",True,False
End Sub

Sub TaskKill(sProcess)
	LogItem "Task Kill: Started",True,False
	Dim bRunning, oProcesses, oProcess
	Dim oWMI
	
	Set oWMI = GetObject("winmgmts://./root/cimv2")
	Set oProcesses = oWMI.ExecQuery("Select * From Win32_Process",,48)
	
	bRunning = False	
	On Error Resume Next
	For Each oProcess in oProcesses
		If sProcess = oProcess.Name Then
			bRunning = True
			LogItem "Task Kill: Found " & sProcess & " and will now be terminated",True,False
			oProcess.Terminate()			
		End If
	Next
	On Error GoTo 0

	If bRunning Then
		' Wait and make sure the process is terminated.
		LogItem "Task Kill: Validating if process is terminated",True,False
		Do Until Not bRunning
			Set oProcesses = oWMI.ExecQuery("Select * From Win32_Process Where Name = '" & sProcess & "'")
			'Wait for 100 MilliSeconds
			WScript.Sleep 100
			'If no more processes are running, exit Loop
			If oProcesses.Count = 0 Then 
				bRunning = False
			End If
		Loop
	Else
		LogItem "Task Kill: " & sProcess & " Is not running",True,False
	End If
	LogItem "Task Kill: Finished",True,False
End Sub

Sub WaitForProcess(sProcessName)
	Dim oWMI,oProcess 
	Set oWMI = GetObject("winmgmts://./root/cimv2")
	
	Do
		set oProcess = oWMI.ExecQuery("Select * From Win32_Process Where Name='" & sProcessName & "'")
		If oProcess.Count = 0 Then
			Exit Do
		End If
		WScript.Sleep 1000
	Loop
End Sub

Sub WaitAndKill(sProcessName)
	Dim oWMI, oProcess
	Set oWMI = GetObject("winmgmts://./root/cimv2")

	Do
		Set oProcess = oWMI.ExecQuery("Select * From Win32_Process Where Name='" & sProcessName & "'")
		If oProcess.Count <> 0 Then
			TaskKill(sProcessName)
			Exit Do
		End If
		WScript.Sleep 1000
	Loop
End Sub

'========================================================================
' Service Manipulation Workers
'========================================================================

Sub StopService(sService)
  LogItem "Stop Service: Started",True,False
    Dim sCmd,iRetVal
    sCmd = "sc stop " & sService
    On Error Resume Next
        LogItem "Stop Service: About to run: " & sCmd,True,False
        iRetVal = oShell.Run(sCmd,0,True)
        If iRetVal <> 0 Then
            LogItem "Stop Service: " & sService & " " & "Exit Code: " & iRetVal & " Error Number: " & Err.number,True,False
        Else
            LogItem "Stop Service: " & sService & " " & "Exit Code: " & iRetVal & " Return Number: " & Err.number,True,False
        End If
    On Error GoTo 0
    LogItem "Stop Service: Finished",True,False
End Sub

Sub StartService(sService)
 	LogItem "Start Service: Started",True,False
    Dim sCmd,iRetVal
    sCmd = "sc start " & sService
    On Error Resume Next
        LogItem "Start Service: About to run: " & sCmd,True,False
        iRetVal = oShell.Run(sCmd,0,True)
        If iRetVal <> 0 Then
            LogItem "Start Service: " & sService & " " & "Exit Code: " & iRetVal & " Error Number: " & Err.number,True,False
        Else
            LogItem "Start Service: " & sService & " " & "Exit Code: " & iRetVal & " Return Number: " & Err.number,True,False
        End If
    On Error GoTo 0
    LogItem "Start Service: Finished",True,False
End Sub

Sub RestartService(sService)
	LogItem "Restart Service: Started",True,False
	StopService sService
	StartService sService
	LogItem "Restart Service: Finished",True,False
End Sub 

Sub ConfigService(sService,sState)
 	LogItem "Configure Service: Started",True,False
    Dim sCmd,iRetVal
    sCmd = "sc config " & sService & " start= " & sState
    LogItem "Configure Service: About to configure" & sService & " to start as " & sState,True,False
    On Error Resume Next
        iRetVal = oShell.Run(sCmd,0,True)
        If iRetVal <> 0 Then
            LogItem "Configure Service: " & sService & " " & "Exit Code: " & iRetVal & " Error Number: " & Err.number,True,False
        Else
            LogItem "Configure Service: " & sService & " " & "Exit Code: " & iRetVal & " Return Number: " & Err.number,True,False
        End If
    On Error GoTo 0
    LogItem "Configure Service: Finished",True,False
End Sub

'========================================================================
' File and Folder Workers
'========================================================================

Function FileExist(sFilePath)
	If oFSO.FileExists(sFilePath) Then
		FileExist = True
	Else
		FileExist = False
	End If
End Function

Sub CopyFile(sSourceFile, sDest)
	On Error Resume Next
	Dim oFile, sFile
	If Right(sDest,1) = "\" Then
		CreateFolderIfNeeded sDest
		sFile = Right(sSourceFile, Len(sSourceFile) - InStrRev(sSourceFile, "\"))
		If oFSO.fileExists(sDest & sFile) Then
			If bFileOverwrite Then
				Set oFile = oFSO.GetFile(sDest & sFile)
				oFile.Attributes = 0
				Set oFile = Nothing
			End If
		End If
	Else
		CreateFolderIfNeeded Left(sDest, InStrRev(sDest, "\"))
		If oFSO.FileExists(sDest) Then
			If bFileOverwrite Then
				Set oFile = oFSO.GetFile(sDest)
				oFile.Attributes = 0
				Set oFile = Nothing
			End If
		End If
	End If
	oFSO.CopyFile sSourceFile, sDest, bFileOverwrite
	On Error GoTo 0
End Sub

Sub CopyFolder(sSourceFolder, sDestFolder)
	On Error Resume Next
	Dim oFolder, oFile, oSubFolder
	CreateFolderIfNeeded sDestFolder
	If Right(sDestFolder,1) <> "\" Then
		sDestFolder = sDestFolder & "\"
	End If
	Set oFolder = oFSO.GetFolder(sSourceFolder)
	For Each oFile In oFolder.Files
		CopyFile oFile.Path, sDestFolder
	Next

	For Each oSubFolder In oFolder.SubFolders
		CopyFolder oSubFolder.Path, sDestFolder & oSubFolder.Name
	Next
	On Error GoTo 0
End Sub

Sub CreateFolderIfNeeded(sTarget)
	On Error Resume Next
	Dim sFolder, aFolders, sWorkingFolder, i
	sFolder = sTarget
	If Right(sFolder, 1) = "\" Then
		sFolder = Left(sFolder, Len(sFolder) - 1)
	End If
	aFolders = Split(sFolder, "\")
	sWorkingFolder = aFolders(0)
	For i = 1 To UBound(aFolders)
		sWorkingFolder = sWorkingFolder & "\" & aFolders(i)
		If oFSO.FolderExists(sWorkingFolder) = False Then
			LogItem "Create Folder if Needed: Creating Folder " & Quotes & sWorkingFolder & Quotes,True,False
			oFSO.CreateFolder sWorkingFolder
		End If
	Next
	On Error GoTo 0
End Sub

Sub DeleteFile(sFile)
	On Error Resume Next
	Dim iRetVal,oFile
	If oFSO.FileExists(sFile) Then
		Set oFile = oFSO.GetFile(sFile)
		iRetVal = oFile.Delete
	Else
		LogItem sFile & " does not exist.",True,False
	End If
	On Error GoTo 0
End Sub

Sub MoveFile(sSource,sDest)
	On Error Resume Next
	Dim iRetVal
	If oFSO.FileExists(sSource) Then
		CreateFolderIfNeeded(sDest)
		iRetVal = oFSO.MoveFile(sSource,sDest & "\")
	Else
		LogItem sSource & " does not exist.",True,False
	End If
	On Error GoTo 0
End Sub

Sub DeleteFolder(sFolder)
	On Error Resume Next
	Dim iRetVal
	If oFSO.FolderExists(sFolder) Then
		Dim oFolder
		Set oFolder = oFSO.GetFolder(sFolder)
		iRetVal = oFolder.Delete
	Else
		LogItem sFolder & " does not exist.",True,False
	End If
	On Error GoTo 0
 End Sub
 
 Sub DeleteFileFromProfile(sFilePath)
	Dim oReg
	Set oReg = GetObject("winmgmts://./root/default:StdRegProv")
	
	Dim sKeyPath,aSubKeys,oSubKey
	sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	oReg.EnumKey HKEY_LOCAL_MACHINE,sKeyPath,aSubKeys
	
	Dim sProfile,sProfileName
	For Each oSubKey In aSubKeys
		Dim sValueName,sValue,sSubPath
		sValueName = "ProfileImagePath"
		sSubPath = sKeyPath & "\" & oSubKey
		oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,sSubPath,sValueName,sValue
		
		sProfile = Split(sValue, "\")
        sProfileName = sProfile(2)
		
		' filter out unnecessary profiles 
	    If (sProfileName <> "system32") And (sProfileName <> "ServiceProfiles") And (sProfileName <> "UpdatusUser") And (sProfileName <> "Administrator") And (sProfileName <> "z_cseappinstall") Then
        	Dim oFolder,sPath
			Set oFolder = oFSO.GetFolder(sValue)
			sPath =  oFolder & "\" & sFilePath
			If oFSO.FileExists(sPath) Then
				DeleteFile sPath
			End If
		End If
	Next
End Sub

Sub DeleteFolderFromProfile(sFolderPath)
	Dim oReg
	Set oReg = GetObject("winmgmts://./root/default:StdRegProv")
	
	Dim sKeyPath,aSubKeys,oSubKey
	sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	oReg.EnumKey HKEY_LOCAL_MACHINE,sKeyPath,aSubKeys
	
	Dim sProfile,sProfileName
	For Each oSubKey In aSubKeys
		Dim sValueName,sValue,sSubPath
		sValueName = "ProfileImagePath"
		sSubPath = sKeyPath & "\" & oSubKey
		oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,sSubPath,sValueName,sValue
		
		sProfile = Split(sValue, "\")
        sProfileName = sProfile(2)
		
	    ' filter out unnecessary profiles 
	    If (sProfileName <> "system32") And (sProfileName <> "ServiceProfiles") And (sProfileName <> "UpdatusUser") And (sProfileName <> "Administrator") And (sProfileName <> "z_cseappinstall") Then
        	Dim oFolder,sPath
			Set oFolder = oFSO.GetFolder(sValue)
			sPath =  oFolder & "\" & sFolderPath
			If oFSO.FolderExists(sPath) Then
				DeleteFolder sPath
			End If
		End If
	Next
End Sub

Sub CopyFolderToProfile(sSourceFolder,sTargetFolder)
	Dim oReg
	Set oReg = GetObject("winmgmts://./root/default:StdRegProv")
	
	Dim sKeyPath,aSubKeys,oSubKey
	sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	oReg.EnumKey HKEY_LOCAL_MACHINE,sKeyPath,aSubKeys
	
	Dim sProfile,sProfileName
	For Each oSubKey In aSubKeys
		Dim sValueName,sValue,sSubPath
		sValueName = "ProfileImagePath"
		sSubPath = sKeyPath & "\" & oSubKey
		oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,sSubPath,sValueName,sValue
		
		sProfile = Split(sValue, "\")
        sProfileName = sProfile(2)
		
	    ' filter out unnecessary profiles 
	    If (sProfileName <> "system32") And (sProfileName <> "ServiceProfiles") And (sProfileName <> "UpdatusUser") And (sProfileName <> "Administrator") And (sProfileName <> "z_cseappinstall") Then
			Dim iRetVal,sRobocopyCmd
			If oFSO.FolderExists(sSourceFolder) Then
				sRobocopyCmd = "Robocopy.exe " & Quotes & sSourceFolder & Quotes & " " & Quotes & sValue & "\" & sTargetFolder & Quotes & " /E /Z"
				LogItem "About to Execute: " & sRobocopyCmd,True,False
				iRetVal = oShell.Run(sRobocopyCmd,0,True)
				Select Case iRetVal
					Case 0
						LogItem "No files were copied. No failure was encountered. No files were mismatched. The files already exist in the destination directory; therefore, the copy operation was skipped.",True,False
					Case 1
						LogItem "All files were copied successfully.",True,False
					Case 2
						LogItem "There are some additional files in the destination directory that are not present in the source directory. No files were copied.",True,False
					Case 3
						LogItem "Some files were copied. Additional files were present. No failure was encountered.",True,False 
					Case 5
						LogItem "Some files were copied. Some files were mismatched. No failure was encountered.",True,False
					Case 6
						LogItem "Additional files and mismatched files exist. No files were copied and no failures were encountered. This means that the files already exist in the destination directory.",True,False
					Case 7
						LogItem "Files were copied, a file mismatch was present, and additional files were present.",True,True
						QuitScript(1603)
					Case 8
						LogItem "Several files did not copy.",True,True
						QuitScript(1603)
					Case Else
						LogItem "Unknown exception. Error number was " & Err.Number & " and the description was: " & Err.Description,True,True
						QuitScript(1603)
				End Select
			Else
				LogItem "Source folder does not exist. Please validate " & sSourceFolder & " path exists.",True,False
				QuitScript(1603)
			End If
		End If
	Next
End Sub

Sub CopyFileToProfile(sSourceFolder,sTargetFolder,sSourceFileName)
	Dim oReg
	Set oReg = GetObject("winmgmts://./root/default:StdRegProv")
	
	Dim sKeyPath,aSubKeys,oSubKey
	sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	oReg.EnumKey HKEY_LOCAL_MACHINE,sKeyPath,aSubKeys
	
	Dim sProfile,sProfileName
	For Each oSubKey In aSubKeys
		Dim sValueName,sValue,sSubPath
		sValueName = "ProfileImagePath"
		sSubPath = sKeyPath & "\" & oSubKey
		oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,sSubPath,sValueName,sValue
		
		sProfile = Split(sValue, "\")
        sProfileName = sProfile(2)
		
	    ' filter out unnecessary profiles 
	    If (sProfileName <> "system32") And (sProfileName <> "ServiceProfiles") And (sProfileName <> "UpdatusUser") And (sProfileName <> "Administrator") And (sProfileName <> "z_cseappinstall") Then
			Dim iRetVal,sRobocopyCmd
			If oFSO.FolderExists(sSourceFolder) Then
				sRobocopyCmd = "Robocopy.exe " & Quotes & sSourceFolder & Quotes & " " & Quotes & sValue & "\" & sTargetFolder & Quotes & " " & Quotes & sSourceFileName & Quotes & " /E /Z"
				LogItem "About to Execute: " & sRobocopyCmd,True,False
				iRetVal = oShell.Run(sRobocopyCmd,0,True)
				Select Case iRetVal
					Case 0
						LogItem "No files were copied. No failure was encountered. No files were mismatched. The files already exist in the destination directory; therefore, the copy operation was skipped.",True,False
					Case 1
						LogItem "All files were copied successfully.",True,False
					Case 2
						LogItem "There are some additional files in the destination directory that are not present in the source directory. No files were copied.",True,False
					Case 3
						LogItem "Some files were copied. Additional files were present. No failure was encountered.",True,False 
					Case 5
						LogItem "Some files were copied. Some files were mismatched. No failure was encountered.",True,False
					Case 6
						LogItem "Additional files and mismatched files exist. No files were copied and no failures were encountered. This means that the files already exist in the destination directory.",True,False
					Case 7
						LogItem "Files were copied, a file mismatch was present, and additional files were present.",True,True
						QuitScript(1603)
					Case 8
						LogItem "Several files did not copy.",True,True
						QuitScript(1603)
					Case Else
						LogItem "Unknown exception. Error number was " & Err.Number & " and the description was: " & Err.Description,True,True
						QuitScript(1603)
				End Select
			Else
				LogItem "Source folder does not exist. Please validate " & sSourceFolder & " path exists.",True,False
				QuitScript(1603)
			End If
		End If
	Next
End Sub


'========================================================================
' Product Discovery Workers
'========================================================================

Function IsProductInstalled
	If IsRegKeyExist("HKLM","SOFTWARE\FRB\Applications\" & sPkgName) Then
		IsProductInstalled = True
		LogItem sPkgName & " is installed.",True,False
	Else
		IsProductInstalled = False
		LogItem sPkgName & " is not installed.",True,False
	End If
End Function

'========================================================================
' Install Validation and Tattoo Workers
'========================================================================

Sub ValidateInstall
	Dim sInstalledFileVer
	sInstalledFileVer = oFSO.GetFileVersion(sProductFilePath)
	
	On Error Resume Next
	If sInstalledFileVer <> sProductFileVersion Then 
		LogItem "Unable to validate install." & sInstalledFileVer,True,False 
		QuitScript(1603)
	Else
		LogItem "Validated installation.",True,False
	End If
	On Error GoTo 0
End Sub

Sub ValidateInstall_File()
	LogItem "Validate Install: Started",True,False
	'Validate installation
	If (oFSO.FileExists(sProductFilePath)) Then 
		LogItem "Validate Install: Able to validate Target File existence. Target File path is: " & sProductFilePath,True,False 
	Else
		LogItem "Validate Install: Unable to validate Target File.",True,False
		QuitScript(1603)
	End If
	LogItem "Validate Install: Finished",True,False
End Sub

Sub ValidateUninstall_File()
	LogItem "Validate Uninstall: Started",True,False
	'Validate Uninstallation
	If Not (oFSO.FileExists(sProgFiles64 & "\Citrix\Virtual Desktop Agent\BrokerAgent.exe")) Then 
		LogItem "Validate Uninstall: Able to validate non-existence of Target File.",True,False 
	Else
		LogItem "Validate Uninstall: Target File still exits at: " & sProgFiles64 & "\Citrix\Virtual Desktop Agent\BrokerAgent.exe",True,False
		QuitScript(1603)
	End If
	LogItem "Validate Uninstall: Finished",True,False
End Sub

Function GetRegPath
	If IsRegKeyExist("HKLM",sRegKeyRoot64 & "\Microsoft\Windows\CurrentVersion\Uninstall\" & sProductCode) Then
		sRegPath = sRegKeyRoot64 & "\Microsoft\Windows\CurrentVersion\Uninstall\" & sProductCode
		sTattooPath = sRegKeyRoot64 & "\" & sRegAuditPath
	ElseIf IsRegKeyExist("HKLM",sRegKeyRoot32 & "\Microsoft\Windows\CurrentVersion\Uninstall\" & sProductCode) Then	
		sRegPath = sRegKeyRoot32 & "\Microsoft\Windows\CurrentVersion\Uninstall\" & sProductCode
		sTattooPath = sRegKeyRoot32 & "\" & sRegAuditPath
	End If
End Function

Sub TattooRegistry
    SetRegVal "HKLM",sTattooPath & "\" & sPkgName,"Install Date","REG_SZ",sInstallDate
    SetRegVal "HKLM",sTattooPath & "\" & sPkgName,"Vendor Name","REG_SZ",sProductVendor
    SetRegVal "HKLM",sTattooPath & "\" & sPkgName,"Product Version","REG_SZ",sProductVersion
    SetRegVal "HKLM",sTattooPath & "\" & sPkgName,"Product Code","REG_SZ",sProductCode
    SetRegVal "HKLM",sTattooPath & "\" & sPkgName,"Target Directory","REG_SZ",sInstallFolderPath
    SetRegVal "HKLM",sTattooPath & "\" & sPkgName,"Install Source","REG_SZ",sScriptDir
End Sub

Sub RemoveTattoo
	DeleteKey "HKLM",sTattooPath & "\" & sPkgName
	'DeleteKey "HKLM","SOFTWARE\Wow6432Node\FRB\Applications\" & sPkgName
End Sub

Sub ARPCustomization
	SetRegVal "HKLM",sRegPath,"DisplayName","REG_SZ",sPkgName
    SetRegVal "HKLM",sRegPath,"NoModify","REG_DWORD","1"
    SetRegVal "HKLM",sRegPath,"NoRepair","REG_DWORD","0"
    SetRegVal "HKLM",sRegPath,"Comments","REG_SZ","Script Package"
    SetRegVal "HKLM",sRegPath,"Contact","REG_SZ",""
    SetRegVal "HKLM",sRegPath,"HelpLink","REG_SZ",""
    SetRegVal "HKLM",sRegPath,"Readme","REG_EXPAND_SZ",""
    SetRegVal "HKLM",sRegPath,"URLUpdateInfo","REG_SZ",""
    SetRegVal "HKLM",sRegPath,"URLInfoAbout","REG_SZ",""
End Sub

Function TattooInstallFolder(sInstallDir)
	Dim sTattooFilePath
	oFSO.CreateTextFile(sInstallDir)
End Function 

'=======================================================================
' Per-User Registry Workers
'========================================================================

Sub AddToUserHives(sUserRegPath,sValueName,sType,sValue)
    Dim oReg
    Set oReg = GetObject("winmgmts://./root/default:StdRegProv")

    Dim sProListRegPath
    sProListRegPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

    ' Enumerate profile list from registry
    Dim oProfile,oProfiles,sProfileDir
	oReg.EnumKey HKEY_LOCAL_MACHINE, sProListRegPath, oProfiles

    Dim sProfile,sProfileName
    Dim iRetVal
    For Each oProfile In oProfiles
        oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, sProListRegPath & "\" & oProfile, "ProfileImagePath", sProfileDir

        sProfile = Split(sProfileDir, "\")
        sProfileName = sProfile(2)

        ' filter out unnecessary profiles 
	    If (sProfileName <> "system32") And (sProfileName <> "ServiceProfiles") And (sProfileName <> "UpdatusUser") And (sProfileName <> "Administrator") And (sProfileName <> "z_cseappinstall") Then
			If FileExist(sProfileDir & "\NTuser.dat") Then
				LogItem "Will try to mount user hive: " & sProfileDir & "\NTuser.dat",True,False
				If MountHive(sProfileDir & "\NTuser.dat") Then
					SetRegVal "HKU","CUSTOM\" & sUserRegPath,sValueName,sType,sValue
					UnmountHive
				Else
					Dim oWMI
					Set oWMI = GetObject("winmgmts://./root/cimv2")
					
					Dim sDomain
					sDomain = oNetwork.UserDomain
					
					Dim oAccount
					Set oAccount = oWMI.Get("Win32_UserAccount.Name='" & sProfileName & "',Domain='" & sDomain & "'")
					
					Dim sSID 
					sSID = oAccount.SID
					
					SetRegVal "HKU",sSID & "\" & sUserRegPath,sValueName,sType,sValue
				End If
			End If
		End If
	Next

    MountDefaultHive
    LogItem "Will now update default user hive for all users.",True,False
	SetRegVal "HKU","CUSTOM\" & sUserRegPath,sValueName,sType,sValue
    LogItem "Default user hive is now updated for all users.",True,False
	UnmountHive
End Sub

Sub DeleteFromUserHives(sUserRegPath)
    Dim oReg
    Set oReg = GetObject("winmgmts://./root/default:StdRegProv")

    Dim sProListRegPath
    sProListRegPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

    ' Enumerate profile list from registry
    Dim oProfile,oProfiles,sProfileDir
	oReg.EnumKey HKEY_LOCAL_MACHINE, sProListRegPath, oProfiles

    Dim sProfile,sProfileName
    Dim iRetVal
    For Each oProfile In oProfiles
        oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, sProListRegPath & "\" & oProfile, "ProfileImagePath", sProfileDir

        sProfile = Split(sProfileDir, "\")
        sProfileName = sProfile(2)

        ' filter out unnecessary profiles 
	    If (sProfileName <> "system32") And (sProfileName <> "ServiceProfiles") And (sProfileName <> "UpdatusUser") And (sProfileName <> "Administrator") And (sProfileName <> "z_cseappinstall") Then
            If FileExist(sProfileDir & "\NTuser.dat") Then
                LogItem "Will try to mount user hive: " & sProfileDir & "\NTuser.dat",True,False
                If MountHive(sProfileDir & "\NTuser.dat") Then
                    DeleteKey "HKU","CUSTOM\" & sUserRegPath
                    UnmountHive
                Else
                    Dim oWMI
					Set oWMI = GetObject("winmgmts://./root/cimv2")
                    
                    Dim sDomain
                    sDomain = oNetwork.UserDomain

                    Dim oAccount
                    Set oAccount = oWMI.Get("Win32_UserAccount.Name='" & sProfileName & "',Domain='" & sDomain & "'")

                    Dim sSID 
                    sSID = oAccount.SID

                    DeleteKey "HKU",sSID & "\" & sUserRegPath
                End If
            End If
        End If
    Next

    MountDefaultHive
    LogItem "Will now update default user hive for all users.",True,False
	DeleteKey "HKU","CUSTOM\" & sUserRegPath
    LogItem "Default user hive is now updated for all users.",True,False
	UnmountHive
End Sub
 
Function MountHive(sHivePath)
	Dim sCmd,iRetVal
	sCmd = "REG.EXE LOAD HKEY_USERS\CUSTOM " & Quotes & sHivePath & Quotes
	LogItem "About to run: " & sCmd,True,False
	iRetVal = oShell.Run(sCmd,0,True)
    If iRetVal <> 0 Then
        MountHive = False
        LogItem Quotes & sHivePath & Quotes & " is currently in use.",True,False
    Else
        MountHive = True
        LogItem Quotes & sHivePath & Quotes & " is now mounted.",True,False
    End If
End Function

Sub UnmountHive
	Dim sCmd
	sCmd = "REG UNLOAD HKEY_USERS\CUSTOM"
	Dim iRetVal
	iRetVal = oShell.Run(sCmd,0,True)
    If iRetVal <> 0 Then
        LogItem "Unable to unmount user hive. Exit code: " & iRetVal,True,False
        QuitScript(1603)
    Else
        LogItem "Unmounted user hive. Exit code: " & iRetVal,True,False
    End If
End Sub

Sub MountDefaultHive
	Dim sCmd
	sCmd = "REG LOAD HKEY_USERS\CUSTOM " & Quotes & "%SYSTEMDRIVE%\Users\Default\NTUSER.DAT" & Quotes
	Dim iRetVal
	iRetVal = oShell.Run(sCmd,0,True)
End Sub

Sub RemoveExcelAddIn(sAddInPath)
	Dim oReg
    Set oReg = GetObject("winmgmts://./root/default:StdRegProv")

	Dim sRegValue,i
	Dim sExcelRegPath,aExcelRegSubKeys,sExcelRegType
	
	sExcelRegPath = "SOFTWARE\Microsoft\Office\14.0\Excel\Options"
	
	Dim sProListRegPath
	sProListRegPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

	' Enumerate profile list from registry
	Dim oProfile,oProfiles,sProfileDir
	oReg.EnumKey HKEY_LOCAL_MACHINE, sProListRegPath, oProfiles

	Dim sProfile,sProfileName
	Dim arrRegValue,intIndex
	Dim tempString
	Dim iRetVal
	For Each oProfile In oProfiles
		oReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, sProListRegPath & "\" & oProfile, "ProfileImagePath", sProfileDir

		sProfile = Split(sProfileDir, "\")
		sProfileName = sProfile(2)

		' filter out unnecessary profiles 
	    If (sProfileName <> "system32") And (sProfileName <> "ServiceProfiles") And (sProfileName <> "UpdatusUser") And (sProfileName <> "Administrator") And (sProfileName <> "z_cseappinstall") Then
			If FileExist(sProfileDir & "\NTuser.dat") Then
				LogItem "Will try to mount user hive: " & sProfileDir & "\NTuser.dat",True,False
				If MountHive(sProfileDir & "\NTuser.dat") Then
					' Hive not in use
					oReg.EnumValues HKEY_USERS,"CUSTOM\" & sExcelRegPath,aExcelRegSubKeys,sExcelRegType
					If IsNull(aExcelRegSubKeys) = False Then
						For i=0 To UBound(aExcelRegSubKeys)
							If (sExcelRegType(i)=1) Or (sExcelRegType(i)=2) Then
								oReg.GetStringValue HKEY_USERS,"CUSTOM\" & sExcelRegPath, aExcelRegSubKeys(i),sRegValue
								If Left(sRegValue,1) = Chr(34) Then
									tempString = Replace(sRegValue,Quotes,"")
									arrRegValue = Split(tempString,"\")
									intIndex = UBound(arrRegValue)
									If LCase(arrRegValue(intIndex)) = LCase(sAddInPath) Then
										LogItem "About to delete HKU" & sSID & "\" & sExcelRegPath & "\" & aExcelRegSubKeys(i),True,False
										DelRegValName "HKU","CUSTOM\" & sExcelRegPath,aExcelRegSubKeys(i)
									End If
								End If
							End If
						Next
					End If
					UnmountHive
				Else
					' Hive in use
					Dim oWMI
					Set oWMI = GetObject("winmgmts://./root/cimv2")
                
                	Dim sDomain
                	sDomain = oNetwork.UserDomain

                	Dim oAccount
                	Set oAccount = oWMI.Get("Win32_UserAccount.Name='" & sProfileName & "',Domain='" & sDomain & "'")

	                Dim sSID 
    	            sSID = oAccount.SID
    	            
					oReg.EnumValues HKEY_USERS,sSID & "\" & sExcelRegPath,aExcelRegSubKeys,sExcelRegType
					If IsNull(aExcelRegSubKeys) = False Then
						For i=0 To UBound(aExcelRegSubKeys)
							If (sExcelRegType(i)=1) Or (sExcelRegType(i)=2) Then
								oReg.GetStringValue HKEY_USERS,sSID & "\" & sExcelRegPath,aExcelRegSubKeys(i),sRegValue
								If Left(sRegValue,1) = Chr(34) Then
									tempString = Replace(sRegValue,Quotes,"")
									arrRegValue = Split(tempString,"\")
									intIndex = UBound(arrRegValue)
									If LCase(arrRegValue(intIndex)) = LCase(sAddInPath) Then
										LogItem "About to delete HKU" & sSID & "\" & sExcelRegPath & "\" & aExcelRegSubKeys(i),True,False
										DelRegValName "HKU",sSID & "\" & sExcelRegPath,aExcelRegSubKeys(i)
									End If
								End If
							End If
						Next
					End If
            	End If
        	End If
    	End If
	Next
End Sub

Function CreateGuid
    CreateGuid = Left(CreateObject("Scriptlet.TypeLib").Guid,38)
End Function
