Option Explicit
' Set Registry Constants
Const HKEY_CLASSES_ROOT =&H80000000
Const HKEY_CURRENT_USER =&H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003

Dim Quotes,oShell,oShellApp,oFSO,oNetwork
Quotes = Chr(34)
Set oShell = CreateObject("WScript.Shell")
Set oShellApp = CreateObject("Shell.Application")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oNetwork = CreateObject("WScript.Network")

DeleteFromUserHives "Software\Wow6432Node\Microsoft\Active Setup\Installed Components\{B00463FF-74E0-4EA8-AF81-F62F99EED251}"
DeleteFromUserHives "Software\Microsoft\Internet Explorer\Main\HangRecovery"


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
	   If (sProfileName <> "config") And (sProfileName <> "system32") And (sProfileName <> "ServiceProfiles") And (sProfileName <> "UpdatusUser") And (sProfileName <> "Administrator") And (sProfileName <> "Administrator."  & oNetwork.ComputerName) And (sProfileName <> "z_cseappinstall") And (sProfileName <> "ctx_cpsvcuser") And (sProfileName <> "MsDtsServer110") And (sProfileName <> "ReportServer") And (sProfileName <> "MSSQLFDLauncher") And (sProfileName <> "SQLSERVERAGENT") And (sProfileName <> "MSSQLSERVER") And (sProfileName <> "QBDataServiceUser26") Then
	         If FileExist(sProfileDir & "\NTuser.dat") Then                
                If MountHive(sProfileDir & "\NTuser.dat") Then
                    DeleteKey "HKU","CUSTOM\" & sUserRegPath
                    UnmountHive
                Else
                    Dim oWMI
					Set oWMI = GetObject("winmgmts://./root/cimv2")
                    
                    Dim sDomain
                    sDomain = oNetwork.UserDomain

                    Dim oAccount
                    'Set oAccount = oWMI.Get("Win32_UserAccount.Name='" & sProfileName & "',Domain='" & sDomain & "'")

                    Dim sSID 
                    'sSID = oAccount.SID
					sSID = oProfile
                    DeleteKey "HKU",sSID & "\" & sUserRegPath
                End If
            End If
        End If
    Next
    MountDefaultHive   
	DeleteKey "HKU","CUSTOM\" & sUserRegPath
	UnmountHive
End Sub

Function MountHive(sHivePath)
	Dim sCmd,iRetVal
	sCmd = "REG.EXE LOAD HKEY_USERS\CUSTOM " & Quotes & sHivePath & Quotes
	iRetVal = oShell.Run(sCmd,0,True)
    If iRetVal <> 0 Then
        MountHive = False
    Else
        MountHive = True
    End If
End Function

Sub UnmountHive
	Dim sCmd
	sCmd = "REG UNLOAD HKEY_USERS\CUSTOM"
	Dim iRetVal
	iRetVal = oShell.Run(sCmd,0,True)
    If iRetVal <> 0 Then
        Wscript.Quit
    Else
    End If
End Sub

Sub MountDefaultHive
	Dim sCmd
	sCmd = "REG LOAD HKEY_USERS\CUSTOM " & Quotes & "%SYSTEMDRIVE%\Users\Default\NTUSER.DAT" & Quotes
	Dim iRetVal
	iRetVal = oShell.Run(sCmd,0,True)
End Sub

Sub DeleteKey(sRootKey,sSubKey)
	Dim sKeyName,iRetVal
	sKeyName = sRootKey & "\" & sSubKey
    If IsRegKeyExist(sRootKey,sSubKey) Then        
        iRetVal = oShell.Run("REG DELETE " & Quotes &  sKeyName & Quotes & " /f",0,True)
        If iRetVal <> 0 Then
            'LogItem Quotes & sKeyName & Quotes & " was not deleted",True,False
        Else
            'LogItem Quotes & sKeyName & Quotes & " has been deleted",True,False
        End If
    Else
        'LogItem Quotes & sKeyName & Quotes & " does not exist.",True,False
    End If
End Sub

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

Function FileExist(sFilePath)
	If oFSO.FileExists(sFilePath) Then
		FileExist = True
	Else
		FileExist = False
	End If
End Function