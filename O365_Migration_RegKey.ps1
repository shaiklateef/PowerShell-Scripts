param([switch]$Uninstall)

if((Test-Path Variable:\PSScriptRoot) -eq $false){
    New-Variable -Name PSScriptRoot -Value $(Split-Path $MyInvocation.MyCommand.Path -Parent) -Option Constant -Scope script
} 
$ErrorActionPreference="Stop"
Trap {"Error: $_"; Break;}
#Set-StrictMode -Version Latest
#Set-StrictMode -Version 2.0

New-Variable -Name sScriptVer -Value "1.7" -Scope script

New-Variable -Name HKEY_CLASSES_ROOT -Value "&H80000000" -Option Constant -Scope script
New-Variable -Name HKEY_CURRENT_USER -Value "&H80000001" -Option Constant -Scope script
New-Variable -Name HKEY_LOCAL_MACHINE -Value "&H80000002" -Option Constant -Scope script
New-Variable -Name HKEY_USERS -Value "&H80000003" -Option Constant -Scope script


# Set Registry Constants
# SetCustomizations
$sPkgName = $sProductVendor = $sProductName = $sProductVersion = $sProductCode = $sPreviousVersionProductCode = $sProductFilePath_PreReq = $null
$sInstallFile = $sInstallFolderPath = $sInstallFolderPath2 = $sProductFilePath = $sProductFileVersion = $sScriptFileName = $sSystemLockStatus = $null
$sCompanyName = $sRegAuditPath = $sLogPath = $bFileOverwrite = $null

# Init
$oShell = $oShellApp = $oFSO = $oNetwork = $oSMSClient = $oSysInfo = $null
$oClass = $colSettings = $oOS = $colComputer = $oComputer = $colItems = $oItem = $null
$colTimeZone = $oTimeZone = $oProcessor = $oBIOS = $sOSName = $null
$sProcessArchitectureX86 = $sProcessArchitectureW6432 = $null
$iOSArch = $iScriptArch = $null
$sSysDir32 = $sSysDir64 = $null
$sProgFiles32 = $sProgFiles64 = $null
$sRegKeyRoot32 = $sRegKeyRoot64 = $null
$sAllUsersStartMenu = $sAllUsersStartPrograms = $null
$sAllUsersDesktop = $sAllUsersAppData = $null
$sWinDir = $sSystemDrive = $sTempDir = $sScriptDir = $null
$sUserProfile = $bUninstall = $Quotes = $sInstallDate = $null
$sTempGuid = $sAppData = $null
$sWestController = $sEastController = $null

$sSystemTimeZone = $null

# Logging
$sVerb = $sLogText = $null

# GetRegPath
$sRegPath = $sTattooPath = $null

# Set Task Sequence Variables
$sTSKeyPath = $sTSValueName = $sTSValue = $sTSName = $sTSNameData = $sBIOSKeyPath = $sBIOReleaseDate = $sBIOSValue = $null
$sBaseImageKeyPath = $sBaseImageName = $sBaseImageNameData = $sBaseImageVersion = $sBaseImageVersionData = $null
function SetCustomizations{
try{
	   $script:sPkgName = "Google_Chrome_74.0.3729.157" # Application Friendly Name
       $script:sProductVendor = "Google LLC" # Product vendor
       $script:sProductName = "Chrome"  # Product name
       $script:sProductVersion = "74.0.3729.157" # Product version
       $script:sProductCode = "{934334B7-4870-3796-B00B-BB14FA736926}" # Product code
       $script:sPreviousVersionProductCode = "{6D86A294-19C9-3B64-A92E-E1BD29883C66}" # Product code
       $script:sInstallFile = "Chrome_74.0\GoogleChromeStandaloneEnterprise.msi" # Install file name
       $script:sInstallFolderPath = "$sProgFiles32\Google\Chrome" # Install folder path
       $script:sProductFilePath = "$sInstallFolderPath\Application\chrome.exe" # Main application executable
       $script:sProductFileVersion = "74.0.3729.157" # Main application executable product 
       $script:sCompanyName = "FRB" # Variable used to create log folder + registry key path
       $script:sScriptFileName = "Chrome_74.0_ScriptedInstall.ps1" # Variable for VBScript file name
    
    
    $script:sRegAuditPath = "$sCompanyName\Applications"
    $script:sLogPath = "$sSystemDrive\$sCompanyName\Logs"
	$script:quotes= '"'
    $script:bFileOverwrite = $true  # Set to $true to overwrite existing files if they exist
	#if((test-path "$sLogPath\$sPkgName-(Script).log")){Remove-Item "$sLogPath\$sPkgName-(Script).log" -Force}
}catch{LogItem "Error occurred at SetCustomizations $($_.exception.message)" $True $False}
}

function GetTaskSeqInfo{
try{
	$script:sTSKeyPath = "SOFTWARE\FRB\OS Deployment"
	$script:sTSName = "Task Sequence Name"
	$script:sTSNameData	= (Get-ItemProperty -Path "HKLM:\$($sTSKeyPath)").$sTSName
	$script:sTSValueName = "Task Sequence Version"
	$script:sTSValue = (Get-ItemProperty -Path "HKLM:\$($sTSKeyPath)").$sTSValueName
	$script:sBIOSKeyPath = "HARDWARE\DESCRIPTION\System\BIOS"
	$script:sBIOReleaseDate = "BIOSReleaseDate"
	$script:sBIOSValue	= (Get-ItemProperty -Path "HKLM:\$($sBIOSKeyPath)").$sBIOReleaseDate
	$script:sBaseImageKeyPath = "SOFTWARE\FRB\Base Image"
	$script:sBaseImageName = "Base Image Creation Task Sequence Name"
	$script:sBaseImageNameData = (Get-ItemProperty -Path "HKLM:\$($sBaseImageKeyPath)").$sBaseImageName
	$script:sBaseImageVersion = "Base Image Creation Version"
	$script:sBaseImageVersionData = (Get-ItemProperty -Path "HKLM:\$($sBaseImageKeyPath)").$sBaseImageVersion
}catch{LogItem "Error occurred at GetTaskSeqInfo $($_.exception.message)" $True $False}
}


#========================================================================
# Main Script Logic
#========================================================================

function MainScript{
try{
    Init
    SetCustomizations
    GetTaskSeqInfo
    SystemLockStatus
    BeginLog
    If ($Uninstall){
           If (-not $(IsProductInstalled)){
                  LogItem "Deployment script will now exit!" $True $False
                  QuitScript(0)
				}
			Else{
                  # Uninstall Logic
				  GetRegPath
				  TaskKill "chrome.exe"
				  TaskKill "GoogleCrashHandler.exe"
				  TaskKill "GoogleCrashHandler64.exe"
				  LogItem "Deployment script will now proceed with $sVerb." $True $False

				  $sChromeUninst_CMD, $sUninstExitCode
                  $sChromeUninst_CMD = "$Quotes$sWinDir\System32\msiexec.exe$Quotes /x {934334B7-4870-3796-B00B-BB14FA736926} /qb-! ALLUSERS=1 REINSTALLMODE=omus REBOOT=ReallySuppress /l*xv $Quotes$sLogPath\$($sPkgName)_(MSI)_Uninstall.log$Quotes"
                  LogItem ("Uninstall Command for - $sPkgName - is: $sChromeUninst_CMD") $True $False
                  LogItem ("Uninstallation has started for: $sPkgName") $True $False
                  $sUninstExitCode = $oShell.Run($sChromeUninst_CMD,0,$True)
                  LogItem "Uninstalled - $sPkgName - with an Exit code: $sUninstExitCode" $True $False
                  If ($sUninstExitCode -eq "1603") {
                         LogItem "Repair of currently installed version is required before installing new version of Google Chrome." $True $False
						 $sChromeRepair_CMD, $sRepairExitCode
                		$sChromeRepair_CMD = "$Quotes$sWinDir\System32\msiexec.exe$Quotes /fvmous {934334B7-4870-3796-B00B-BB14FA736926} /qb-! ALLUSERS=1 REINSTALLMODE=omus REBOOT=ReallySuppress$Quotes"
                		LogItem ("Repair Command for previously installed Google_Chrome_74.0.3729.157 - is: $sChromeRepair_CMD") $True $False
                		$sRepairExitCode = $oShell.Run($sChromeRepair_CMD,0,$True)
               			LogItem "Repaired - Google_Chrome_74.0.3729.157 - with an Exit code: $sRepairExitCode" $True $False
						LogItem "Uninstalling - Google_Chrome_74.0.3729.157" $True $False
						UninstallMSI $sProductCode "/qb-! REBOOT=ReallySuppress /l*xv $Quotes$sLogPath\$($sPkgName)_(MSI)_Uninstall.log$Quotes"
                  }
				  Start-Sleep -Milliseconds 15000 
				  ValidateUninstall_File
                  #Deleting leftover shortcut(s)
                  DeleteFile "$sAllUsersStartPrograms\Google Chrome.lnk"
                  DeleteFile "$sProgFiles32\FRB Programs\Google Chrome.lnk"
                  DeleteFile "$sSystemDrive\Users\Public\Desktop\Google Chrome.lnk"             
                  DeleteFolder $sInstallFolderPath
                  DeleteFolder "$sProgFiles32\Google\CrashReports" 
                  DeleteFolder "$sProgFiles32\Google\Update"
				  DeleteFolderIfEmpty "$sProgFiles32\Google"              
                  $oShell.Run("cmd /c RD /S /Q $Quotes$sProgFiles32\Google\Update$Quotes")
                  Uninstall "$sSysDir64\schtasks.exe" "/delete /f /tn `"GoogleUpdateTaskMachineCore`""
                  Uninstall "$sSysDir64\schtasks.exe" "/delete /f /tn `"GoogleUpdateTaskMachineUA`""
              
                  # Deleting Google Update Services
                  Delservice "gupdate"
                  Delservice "gupdatem"
                  # Cleanup per-user folder
                  DeleteFolderFromProfile "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Google Chrome"
				  DeleteFileFromProfile "AppData\Roaming\Microsoft\Windows\Internet Explorer\Quick Launch\Google Chrome.lnk"
                  RemoveTattoo
				  DeleteKey "HKLM" "SOFTWARE\Classes\Installer\Products\1281FAC89A051793882C73A1CA0E301E"
				  DeleteKey "HKLM" "$sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode"
				  DeleteKey "HKLM" "$sRegKeyRoot32\Microsoft\Active Setup\Installed Components\$sProductCode"
				  DeleteFromUserHives "Software\Wow6432Node\Microsoft\Active Setup\Installed Components\$sProductCode"
				}
		}
	Else{ 
		If (IsProductInstalled) {
                  LogItem "Deployment script will now exit!" $True $False
                  QuitScript(0)
           }
		 Else{
            # Install Logic
            TaskKill "chrome.exe"
            TaskKill "GoogleCrashHandler.exe"
            TaskKill "GoogleCrashHandler64.exe"
             
			 
           
		   $sChrome62Productcode = "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($sPreviousVersionProductCode)" 
			If (($sOSName -eq "Microsoft Windows 10 Enterprise") -and (Test-Path $sChrome62Productcode)){
			   If(Test-Path $sProductFilePath){
				$sChromeRepair_CMD, $sRepairExitCode
                  $sChromeRepair_CMD = "$Quotes$sWinDir\System32\msiexec.exe$Quotes /fmous {6D86A294-19C9-3B64-A92E-E1BD29883C66} /qb-! ALLUSERS=1 REINSTALLMODE=omus REBOOT=ReallySuppress$Quotes"
                  LogItem ("Repair Command for previously installed Google_Chrome_62.0.3202.75 - is: $sChromeRepair_CMD") $True $False
                  $sRepairExitCode = $oShell.Run($sChromeRepair_CMD,0,$True)
                  LogItem "Repaired - Google_Chrome_62.0.3202.75 - with an Exit code: $sRepairExitCode" $True $False
                  LogItem "Uninstalling - Google_Chrome_62.0.3202.75" $True $False
                  UninstallMSI $sPreviousVersionProductCode "/qb-! REBOOT=ReallySuppress /l*xv $Quotes$sLogPath\Google_Chrome_62.0.3202.75_(MSI)_Uninstall.log$Quotes"
					
				   Start-Sleep -Milliseconds 30000 
					
				   ValidateUninstall_File
                       
                  #Deleting leftover shortcut(s)
                  DeleteFile "$sAllUsersStartPrograms\Google Chrome.lnk"
                  DeleteFile "$sProgFiles32\FRB Programs\Google Chrome.lnk"
                  DeleteFile "$sSystemDrive\Users\Public\Desktop\Google Chrome.lnk"             
                  DeleteFolder $sInstallFolderPath
                  DeleteFolder "$sProgFiles32\Google\CrashReports" 
                  DeleteFolder "$sProgFiles32\Google\Update"
				  
				  DeleteFolderIfEmpty "$sProgFiles32\Google"              
                  $oShell.Run("cmd /c RD /S /Q $Quotes$sProgFiles32\Google\Update$Quotes")
                  Uninstall "$sSysDir64\schtasks.exe" "/delete /f /tn `"GoogleUpdateTaskMachineCore`""
                  Uninstall "$sSysDir64\schtasks.exe" "/delete /f /tn `"GoogleUpdateTaskMachineUA`""
              
                  # Deleting Google Update Services
                  Delservice "gupdate"
                  Delservice "gupdatem"
              
                  # Cleanup per-user folder
                  DeleteFolderFromProfile "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Google Chrome"
				  DeleteFileFromProfile "AppData\Roaming\Microsoft\Windows\Internet Explorer\Quick Launch\Google Chrome.lnk"
				  DeleteKey "HKLM" "SOFTWARE\Wow6432Node\FRB\Applications\Google_Chrome_62.0.3202.75"
				  DeleteKey "HKLM" "SOFTWARE\Policies\Google\Chrome"
				  DeleteKey "HKLM" "SOFTWARE\Policies\Google\Update"        
				  DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Policies\Google\Chrome"
                  DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Policies\Google\Update"
                  DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Google\Update"
				  DeleteKey "HKLM" "SOFTWARE\Classes\Installer\Products\492A68D69C9146B39AE21EDB9288C366"
				  DeleteKey "HKLM" "$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sPreviousVersionProductCode"
                  }
			}
				  TaskKill "chrome.exe"
				  LogItem "Deleting Audit and Update registry key for previous package" $True $False
				  Cleanup_Chrome_AuditRegKeys
				  DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Google\Update"
				  LogItem "Deployment script will now proceed with $sVerb." $True $False
                  $sChromeInst_CMD, $sInstExitCode
                  $sChromeInst_CMD = "$Quotes$sScriptDir\$sInstallFile$Quotes /qb-! ALLUSERS=1 REINSTALLMODE=omus REBOOT=ReallySuppress /l*xv $Quotes$sLogPath\$($sPkgName)_(MSI)_Install.log$Quotes"
                  LogItem ("Install Command for - $sPkgName - is: $sChromeInst_CMD") $True $False
                  LogItem ("Installation has started for: $sPkgName") $True $False
                  $sInstExitCode = $oShell.Run($sChromeInst_CMD,0,$True)
                  LogItem "Installed - $sPkgName - with an Exit code: $sInstExitCode" $True $False
                  If ($sInstExitCode -eq "1603") {
                         LogItem "Reboot required before installing new version of Google Chrome." $True $False
						 QuitScript 1641
                  }
				  Start-Sleep -Milliseconds 15000
                  ValidateInstall
				  GetRegPath
				  #To copy shortcut
				  LogItem "To copy shortcut to StartMenu and FRB Programs" $True $False
                  CopyFile "$sAllUsersStartPrograms\Google Chrome.lnk" "$sProgFiles32\FRB Programs\Google Chrome.lnk"
				  
				  #To copy preferences file
				  LogItem "To copy preferences file to INSTALLDIR location" $True $False
                  CopyFile "$sScriptDir\Chrome_74.0\master_preferences" "$sInstallFolderPath\Application\master_preferences"
				  
				  LogItem "To copy Del_UserProfileChromeShortcut file to INSTALLDIR location" $True $False
				  CopyFile "$sScriptDir\Chrome_74.0\Del_UserProfileChromeShortcut.vbs" "$sInstallFolderPath\Application\Del_UserProfileChromeShortcut.vbs"
				   
                  LogItem "Copying First Run file to [Userprofile] to avoid second StartMenu Shortcut Creation" $True $False
                  CopyFolderToProfile "$sScriptDir\Chrome_74.0\User Data" "AppData\Local\Google\Chrome\User Data"
				 
				  CopyFolder "$sScriptDir\Chrome_74.0\User Data" "$sSystemDrive\Users\Default\AppData\Local\Google\Chrome\User Data"
					
				  DeleteFileFromProfile "Desktop\Google Chrome.lnk"
					
                  #To delete unwanted shortcut
                  DeleteFile "$sAllUsersDesktop\Google Chrome.lnk"
                  DeleteFolder "$sAllUsersStartPrograms\Google Chrome"
                  DeleteFile "$sSystemDrive\Users\Public\Desktop\Google Chrome.lnk"
				  
				  #LogItem "Create the URL values under PluginsAllowedForUrls location" $True $False
				  #Write-RegKeyURLs
              
                  #Applying Registry Keys settings, to customize Chrome
                  Install "$sWinDir\regedit.exe" "/s `"$sScriptDir\Chrome_74.0\Chrome_CustomSettings.reg`""
              
                  #Executing GPUpdate
                  ExecuteCMD "$sWinDir\SysWOW64\gpupdate.exe" "/force /wait:0"
              
                  #To disable Google Update Schedule Tasks
                  Install "$sSysDir64\schtasks.exe" "/change /disable /tn `"GoogleUpdateTaskMachineCore`""
                  Install "$sSysDir64\schtasks.exe" "/change /disable /tn `"GoogleUpdateTaskMachineUA`""
				  
				  DeleteFile "$sProgFiles32\Google\Update\GoogleUpdate.exe"
              
                  #To disable Google Update Services
                  ConfigService "gupdate" "disabled"
                  ConfigService "gupdatem" "disabled"
				
                  #Applying Registry Keys settings, to disable Google Update
                  SetRegVal "HKLM" "$sRegKeyRoot64\Policies\Google\Update" "AutoUpdateCheckPeriodMinutes" "REG_DWORD" "0"
                  SetRegVal "HKLM" "$sRegKeyRoot64\Policies\Google\Update" "UpdateDefault" "REG_DWORD" "0"
              
                  DeleteFolder "$sProgFiles32\Google\Update"
                  $oShell.Run("cmd /c RD /S /Q $Quotes$sProgFiles32\Google\Update$Quotes")
				  LogItem "Delete user desktop shortcut" $True $False
				  DeleteFileFromProfile "Desktop\Google Chrome.lnk"
				  # Uninstall previous version uninstall strings
				  Cleanup_Chrome_PreviousUninstallStrings		

				  LogItem "Creating active setup to delete user desktop shortcut" $True $False
				  SetRegVal "HKLM" "$sRegKeyRoot32\Microsoft\Active Setup\Installed Components\$sProductCode" "StubPath" "REG_SZ" "wscript.exe \`"$sInstallFolderPath\Application\Del_UserProfileChromeShortcut.vbs\`""
				  SetRegVal "HKLM" "$sRegKeyRoot32\Microsoft\Active Setup\Installed Components\$sProductCode" "Version" "REG_SZ" "1,0,0"
				  
                  #Applying Registry Keys settings, to customize ARP Settings
                  ARPCustomization
                  SetRegVal "HKLM" "SOFTWARE\Classes\Installer\Products\1281FAC89A051793882C73A1CA0E301E" "ProductName" "REG_SZ" $sPkgName         
                  SetRegVal "HKLM" "$sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode" "DisplayIcon" "REG_SZ" $sProductFilePath
                  SetRegVal "HKLM" "$sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode" "DisplayVersion" "REG_SZ" $sProductVersion
				  #DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Google\Update"
                  SetRegVal "HKLM" "$sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode" "UninstallString" "REG_SZ" "powershell.exe -executionpolicy bypass -file \`"$sScriptDir\$sScriptFileName\`" -Uninstall"
				  SetRegVal "HKLM" "$sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode" "Comments" "REG_SZ" "Script Package; SNOW# REQ0468681"
                  TattooRegistry   
           }
    }

    #Success (No Reboot)
    #QuitScript 0

    # Success (No Reboot) 
    #QuitScript(1707)

    #Soft Reboot 
    #QuitScript(3010)

    #Hard Reboot
    QuitScript 1641

    #Force Reboot
    #ForceReboot(1641)
}
catch
{
    LogItem "Error occurred at main script $($_.exception.message)" $True $False
}
}

function Init{
try{
    $script:oShell  = New-Object -ComObject "Wscript.Shell"
    $script:oSysInfo = New-Object -ComObject "ADSystemInfo"
    $script:oShellApp = New-Object -ComObject "Shell.Application"
    $script:oFSO = New-Object -ComObject "Scripting.FileSystemObject"
    $script:oNetwork = New-Object -ComObject "WScript.Network"
    $script:oSMSClient = New-Object -ComObject "Microsoft.SMS.Client"
	
	# Determine OS Architecture
	$sProcessArchitectureX86 = [System.Environment]::ExpandEnvironmentVariables("%PROCESSOR_ARCHITECTURE%")
	If ([System.Environment]::ExpandEnvironmentVariables("%PROCESSOR_ARCHITEW6432%") -eq "%PROCESSOR_ARCHITEW6432%") {
		$sProcessArchitectureW6432 = "Not Defined"
	}

	If (($sProcessArchitectureX86 -eq "x86") -and ($sProcessArchitectureW6432 -eq "Not Defined")){
		# Windows 32-bit
		$script:iOSArch = 32
		$script:iScriptArch = 32
		$script:sSysDir32 = $oShellApp.NameSpace(37).Self.Path
		$script:sProgFiles32 = $oShellApp.NameSpace(38).Self.Path
		$script:sRegKeyRoot64 = "SOFTWARE"
		$script:sRegKeyRoot32 = "SOFTWARE"
    }else{
		# Windows 64-bit
		$script:iOSArch = 64
		$script:iScriptArch = 64
		$script:sSysDir64 = $oShellApp.NameSpace(37).Self.Path
		$script:sProgFiles64 = $oShellApp.NameSpace(38).Self.Path
		$script:sSysDir32 = $oShellApp.NameSpace(41).Self.Path
		$script:sProgFiles32 = $oShellApp.NameSpace(42).Self.Path
		$script:sRegKeyRoot32 = "SOFTWARE\Wow6432Node"
		$script:sRegKeyRoot64 = "SOFTWARE"
	}

	# %ProgramData%\Microsoft\Windows\Start Menu
	$script:sAllUsersStartMenu = $oShellApp.NameSpace(22).Self.Path
		
	# %ProgramData%\Microsoft\Windows\Start Menu\Programs
	$script:sAllUsersStartPrograms = $oShellApp.NameSpace(23).Self.Path
	
	# %SystemDrive%\Users\Public\Desktop
	$script:sAllUsersDesktop = $oShellApp.NameSpace(25).Self.Path
	
	# %SYSTEMDRIVE%\ProgramData
	$script:sAllUsersAppData = $oShellApp.NameSpace(35).Self.Path
	
	# %WINDIR%
	$script:sWinDir = $oShellApp.NameSpace(36).Self.Path
	
	# %SYSTEMDRIVE%
	$script:sSystemDrive = [System.Environment]::ExpandEnvironmentVariables("%SystemDrive%")
			
	# %WINDIR%\Temp - System Account
	$script:sTempDir = [System.Environment]::ExpandEnvironmentVariables("%TEMP%")
	
	#Roaming appdata
	$script:sAppData = [System.Environment]::ExpandEnvironmentVariables("%appdata%")
		
	# Get script directory without trailing slash
	$script:sScriptDir = $PSScriptRoot
	
	#If root of drive, strip trailing backslash
	If (([string]$sScriptDir).Length -eq 3){ $sScriptDir = $sScriptDir.functionstring(0, 2) }
	
	# %SYSTEMDRIVE%\Users\%USERNAME%
	$script:sUserProfile = [System.Environment]::ExpandEnvironmentVariables("%USERPROFILE%")
	
	# Check if /uninstall was passed to the script
	$script:bUninstall = $false

	if($Uninstall){$script:bUninstall = $true }
		
	# used to encapsulate paths
	$script:Quotes = '"'
		
	# Convert Now() to String
	$script:sInstallDate = (Get-Date).ToString()  
	
	# Generate GUID
	$script:sTempGuid = CreateGuid
}catch{LogItem "Error occurred at Init $($_.exception.message)" $True $False}
}

#========================================================================
# Write Flash Reg URL's Values under plugin key
#========================================================================

function Write-RegKeyURLs{
	param(
		[ValidateSet('HKCR','HKLM','HKCU','HKU')]$RegHive="HKLM",
		[string]$RegPath="SOFTWARE\WOW6432Node\Policies\Google\Chrome\PluginsAllowedForUrls",
		[string[]]$RegValues=@("https://global.hgncloud.com/firstrepublic/welcome.jsp","https://frb.service-now.com","https://collaborate.corp.firstrepublic.com","https://eagleinvest.firstrepublic.com","https://app.wdesk.com/auth/login/","https://10.7.77.53/","https://webapps.firstrepublic.com","https://websso.firstrepublic.com/idp/startSSO.ping?PartnerSpId=WageWorks","https://websso.firstrepublic.com/idp/startSSO.ping?PartnerSpId=https%3A%2F%2Fsso.solium.com","https://firstrepublic.icims.com","https://websso.firstrepublic.com/idp/startSSO.ping?PartnerSpId=https://www.docusign.net","https://frb.ultipro.com","https://websso.firstrepublic.com/idp/startSSO.ping?PartnerSpId=Concur","https://flightpath.corp.firstrepublic.com/suite","https://firstrepublic.webex.com/","https://app.proofhq.com/","https://cloud.scorm.com","https://finance.google.com","https://apps/EagleSweep/EagleSweepForms/","https://collaborate.firstrepublic.com/","https://firstrepublic.pagerduty.com/","https://team.corp.firstrepublic.com/SitePages/Home.aspx","https://collaborate.firstrepublic.com/docs/DOC-9052","https://collaborate.corp.firstrepublic.com/docs/","https://[*.]hgncloud.com/firstrepublic/")
	)
	try{
	$FullKeyName = $($RegHive+":\$RegPath")
	[int]$LastHighestValue = 1
	[bool]$newKeyCreated = $false
	if(-not (Test-Path -Path "$FullKeyName")){
		#"$FullKeyName - not found, creating a new key"
		$null = New-Item -Path "$FullKeyName" -Force -Confirm:$false -ErrorAction Stop
		$newKeyCreated = $true
	}
	#"HighestValueWhenStartOfTheScript: $LastHighestValue"
	foreach($RegValue in $RegValues){
		if($newKeyCreated){
			$null = New-ItemProperty -Path "$FullKeyName" -Name $LastHighestValue -PropertyType "String" -Value $RegValue -ErrorAction Stop
			#"Created new element '$LastHighestValue' with the value '$RegValue'"
			$newKeyCreated = $false
			continue
		}
		$OReg = Get-Item -Path "$FullKeyName"
		[array]$String_SZ_WithNos = $OReg.Property | Where-Object { (-not [string]::IsNullOrEmpty($_)) -and ($_ -inotmatch "default") }
		[bool]$found = $false
		[hashtable]$KeyPairs = @{}
		$OReg.Property | `
		ForEach-Object { 
			$curValue = $OReg.GetValue($_)
			if(-not [string]::IsNullOrEmpty($curValue)){
				if($curValue -ieq $RegValue){ 
					$found = $true 
				}
			}
			$KeyPairs.Add([int]$_, $curValue)
		}
		if($found -ne $true){
			$HighestValue = $String_SZ_WithNos | ForEach-Object { [int]$_ } | Sort-Object -Descending | Select-Object -First 1
			$LastHighestValue = [int]$HighestValue
			#"LastHighestValue: $LastHighestValue"
			[array]$MissingNos = @()
			#Check number sequence and report if any number is missing
			for($i = 1; $i -le $LastHighestValue ; $i++){
				if(-not $KeyPairs.ContainsKey($i)){
					$MissingNos += $i
				}else{
					if([string]::IsNullOrEmpty($KeyPairs.$i)){
						$MissingNos += $i
					}
				}
			}
			if($MissingNos.Count -ge 1){
				#"Missing or empty keys: ($($MissingNos -join ',')) found, using this instead"    
				$LastHighestValue = $MissingNos | Sort-Object | Select-Object -First 1
				#"Lowest value from the Missing or Empty keys is : $LastHighestValue"
			}else{
				$LastHighestValue += 1
				#"Incrementing the Last Highest value from the List : $LastHighestValue"
			}
			if($KeyPairs.ContainsKey($LastHighestValue)){
				$null = Set-ItemProperty -Path "$FullKeyName" -Name $LastHighestValue -Value $RegValue -Force -Confirm:$false -ErrorAction Stop
				LogItem "Edited the empty valued element '$LastHighestValue' with the value '$RegValue'" $True $False
			}else{
				$null = New-ItemProperty -Path "$FullKeyName" -Name $LastHighestValue -PropertyType "String" -Value $RegValue -ErrorAction Stop
				LogItem "Created new element '$LastHighestValue' with the value '$RegValue'" $True $False
			}
		}else{ 
			LogItem "$RegValue already found"  $True $False
		}
	}
	}catch{
		LogItem "Error occurred at Init $($_.exception.message)" $True $False
	}
}

#========================================================================
# Check Pending Reboot Routines
#========================================================================
function InstallRebootExe{
param($sFilePath)
try{
	$subroutine = "Reboot Pending status checking: " 
	LogItem "$subroutine Started" $True $False	
	$sCmd =  "$Quotes$sFilePath$Quotes"
   	LogItem "$subroutine Running: $sCmd" $True $False
	$iRetVal = $oShell.Run($sCmd,0,$True)
	If ($iRetVal -eq "0"){
			LogItem "$subroutine No Reboot Required. (Returned $iRetVal)" $True $True
    }Elseif($iRetVal -eq "1") {
		
			LogItem "$subroutine Reboot Required (Returned 3010 -- Soft Reboot Required)" $True $True
			#QuitScript 3010
            ForceReboot(1641)
			exit $iRetVal
	}
}catch{LogItem "Error occurred at InstallRebootExe $($_.exception.message)" $True $False}
}
Function RestartExplorer
{
	Get-Process -ProcessName explorer -ErrorAction SilentlyContinue | Stop-Process -Force
}
#===============================================================================
# For deleting empty directory
#===============================================================================
Function DeleteFolderIfEmpty{
param($sSrc)
try{
        if((Test-Path -Path "$sSrc") -ne $true){
            LogItem "Remove-FolderIfEmpty - Folder doesn't exist $sSrc" $True $False
        }
        else{
            $FolderStatus = Get-ChildItem -Path "$sSrc" -Recurse | Where-Object { !$_.PSisContainer } | Select-Object Length | Measure-Object -Property Length -Sum | Select-Object Sum
			#Get-ChildItem -Path "$sSrc" -Recurse | Select-Object Length | Measure-Object -Property Length -Sum
            if(($FolderStatus -eq $null) -or ($FolderStatus.Sum -eq 0)){
                Remove-Item -Path "$sSrc" -Recurse -Force -ErrorAction Stop | Out-Null
				LogItem "Deleted the empty folder $sSrc" $True $False
            }
            else{
                LogItem "Folder not empty $sSrc" $True $False
            }
        }
}
catch{LogItem "Error occurred at DeleteFolderIfEmpty $($_.exception.message)" $True $False}
}
#========================================================================
# Logging
#========================================================================

function BeginLog{
try{
	If($bUninstall){$Script:sVerb = "Uninstall"}
    Else{ $Script:sVerb = "Install" }
	
	# Create log folder path if needed
	If (-Not $oFSO.FolderExists($sLogPath)){
		CreateFolderIfNeeded $sLogPath
	}

	# Determine OS Properties
	$oOS = Get-WmiObject -Class Win32_OperatingSystem
    $sOSVersion=$sOSManufacturer=$sOSOrganization=$null
    $script:sOSName = $oOS.Caption
	$sOSVersion = "$($oOS.Version) $($oOS.CSDVersion) Build $($oOS.BuildNumber)"
	$sOSManufacturer = $oOS.Manufacturer
	$sOSOrganization = $oOS.Organization			
	
	# Determine ComputerSystem Properties
	$oComputer = Get-WmiObject -Class Win32_ComputerSystem
	$sSystemManufacturer=$sSystemName=$sSystemModel=$sSystemType=$sSystemRAM=$null
	$sSystemManufacturer = $oComputer.Manufacturer
	$sSystemName = $oComputer.Name
	$sSystemModel = $oComputer.Model
	$sSystemType = $oComputer.SystemType		
	$sSystemRAM = "{0} GB" -f ([System.Math]::Round(($oComputer.TotalPhysicalMemory/1073741824),2))

	# Determine LogicalDisk Properties
	$colItems = Get-WmiObject -Class Win32_LogicalDisk
	$sVolumeName=$sTotalHardDriveSize=$sAvailableHardDriveFreeSpace=$null
	ForEach($oItem in $colItems){
		If($oItem.Description -eq "Local Fixed Disk"){				
			$sVolumeName = $oItem.VolumeName
			$sTotalHardDriveSize = "{0} GB" -f ([System.Math]::Round($oItem.Size /1073741824))
			$sAvailableHardDriveFreeSpace = "{0} GB" -f ([System.Math]::Round($oItem.FreeSpace /1073741824))		
		}
	}

	# Determine TimeZone Properties
	$colTimeZone = Get-WmiObject -Class Win32_TimeZone
	
	ForEach($oTimeZone in $colTimeZone){
		$script:sSystemTimeZone = $oTimeZone.StandardName
	}

	# Determine Processor Properties
	$colSettings = Get-WmiObject -Class Win32_Processor
	$sProcessorDescription=$sProcessorName=$null
	ForEach($oProcessor in $colSettings){
		$sProcessorDescription = $oProcessor.Description
		$sProcessorName = $oProcessor.Name
	}

	# Determine BIOS Properties
	$colSettings = Get-WmiObject -Class Win32_BIOS
	$sBIOSVersion = $null
	ForEach($oBIOS In $colSettings){
		$sBIOSVersion = $oBIOS.Manufacturer + " " + $oBIOS.SMBIOSBIOSVersion + ", " + $sBIOSValue	
	}	
	
	$Type = $oSysInfo.GetType()


	LogItem ("******************** Begin " + $sVerb + " ********************")  $true  $false
	LogItem ("Deployment Script Version: " + $sScriptVer) $true $false
	LogItem ("OS Architecture: " + $iOSArch + " / Script Architecture: " + $iScriptArch) $true  $false
	LogItem ("OS Name: " + $sOSName)  $true  $false
	LogItem ("OS Version: " + $sOSVersion)  $true  $false		
	LogItem ("OS Manufacturer: " + $sOSManufacturer)  $true  $false
	LogItem ("Organization: " + $sOSOrganization)  $true  $false
	LogItem ("System Name: " + $sSystemName)  $true  $false
	LogItem ("System Manufacturer: " + $sSystemManufacturer)  $true  $false
	LogItem ("System Model: " + $sSystemModel)  $true  $false
	LogItem ("System Type: " + $sSystemType)  $true  $false	
	LogItem ("Processor: " + $sProcessorName ) $true  $false
	LogItem ("BIOS Version/Date: " + $sBIOSVersion)  $true  $false	
	LogItem ("System RAM: " + $sSystemRAM ) $true  $false
	LogItem ("System Site Name: " + $Type.InvokeMember("sitename","GetProperty",$null,$oSysInfo,$null))  $true  $false
	LogItem ("System Time Zone: " + $sSystemTimeZone)  $true  $false
	LogItem ("System Lock Status: " + $sSystemLockStatus)  $true  $false
	LogItem ("Volume Name: " + $sVolumeName)  $true  $false
	LogItem ("Total Hard Drive Size: " + $sTotalHardDriveSize)  $true  $false
	LogItem ("Available Hard Drive Free Space: " + $sAvailableHardDriveFreeSpace)  $true  $false	
	LogItem ("Domain Information: " + $oNetwork.UserDomain)  $true  $false
	LogItem ("DNS Name: " + $Type.InvokeMember("DomainDnsName","GetProperty",$null,$oSysInfo,$null))  $true  $false
	LogItem ("Assigned Management Point: " + $oSMSClient.GetCurrentManagementPoint())  $true  $false
	LogItem ("Base Image Name: " + $sBaseImageNameData ) $true  $false
	LogItem ("Base Image Version: " + $sBaseImageVersionData )  $true  $false	
	LogItem ("SCCM Task Sequence Name: " + $sTSNameData ) $true  $false
	LogItem ("SCCM Task Sequence Version: " + $sTSValue)  $true  $false
	LogItem ("User Executing 	" + $oNetwork.UserName ) $true  $false		
	LogItem ("Product Vendor: " + $sProductVendor)  $true  $false
	LogItem ("Product Name: " + $sProductName ) $true  $false
	LogItem ("Product Version: " + $sProductVersion ) $true  $false
	LogItem ("Product Code: " + $sProductCode)  $true  $false
	LogItem ("Date: " + (Get-Date))  $true  $false
}catch{LogItem "Error occurred at BeginLog $($_.exception.message)" $True $False}
}

function QuitScript{
param($iRetVal=0)
try{
	LogItem "Exiting Script $iRetVal"  $True $True
	LogItem "******************** End $sVerb ********************" $True $False
	#$sLogText = $sLogText.Trim()
	$oShell = New-Object -ComObject "WScript.Shell"
	switch($iRetVal){
	{($_ -eq 0) -or ($_ -eq 3010)}{
		$oShell.Logevent(4, $iRetVal) #$sLogText)
        }
	default{
		$oShell.LogEvent(1, $iRetVal) #$sLogText)
        }
	}

    exit $iRetVal
}catch{LogItem "Error occurred at QuitScript $($_.exception.message)" $True $False}
}

function ForceReboot{
param($iRetVal)
try{
	LogItem "Exiting Script $iRetVal " $True $True
	LogItem "System gets Force Rebooted." $True $True
	LogItem "******************** End $sVerb ********************" $True $False
	#$sLogText = $sLogText.Trim()
	switch($iRetVal){
	1641{
		$oShell.Logevent(1, $iRetVal) #$sLogText)
		Restart-Computer -Force
        }
	}
	exit $iRetVal
}catch{LogItem "Error occurred at ForceReboot $($_.exception.message)" $True $False}
}

function LogItem{
param($sMessage, $bLogFile, $bEventLog)
try{
	If ($bEventLog){
		$sLogText += "$sMessage`n"
	}

	If($bLogFile){
		#$tsLog = $null ; $tsLog = $oFSO.OpenTextFile("$sLogPath\$sPkgName-Script).log", 8, $true )
		#$tsLog.WriteLine("($((get-date).ToString())) - $sMessage")
        Out-File -FilePath "$sLogPath\$($sPkgName)_(Script).log" -InputObject "$((Get-Date).ToString()) - $sMessage" -Append -Force 
	}
}catch{Write-Output "Error occurred at LogItem $($_.exception.message)"}
}

#========================================================================
# Product Discovery Workers
#========================================================================

Function IsProductInstalled
{
try{
    [boolean]$ProductInstall = $false
        $path64bit = "HKLM:\SOFTWARE\Wow6432Node\FRB\Applications\$($sPkgName)" 
    	$path32bit = "HKLM:\SOFTWARE\FRB\Applications\$($sPkgName)" 
        $path64bitproductcode = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$($sProductCode)" 
        $path32bitproductcode = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($sProductCode)" 
 
    if((Test-Path $path64bit) -or (Test-Path $path32bit)){
        LogItem "AuditKey - Found" $True $False
        if((Test-Path $path64bitproductcode) -or (Test-Path $path32bitproductcode)){
            LogItem "ProductCode - Found" $True $False
            if(Test-Path $sProductFilePath){
                LogItem "ProductFile - Found" $True $False
				$oFSO = New-Object -ComObject "Scripting.FileSystemObject"
				$versionInfo = $oFSO.GetFileVersion($sProductFilePath)
				#Bug found in get-item file version in powershell, switching to good old VBs
                #$versionInfo = (Get-Item $sProductFilePath).VersionInfo
                #LogItem "ProductFileVersion - $($versionInfo.FileVersion)" $True $False
				LogItem "ProductFileVersion - $versionInfo" $True $false
				if($versionInfo -eq $sProductFileVersion){
                #if($versionInfo.FileVersion -eq $sProductFileVersion){
                    $ProductInstall = $true
                    LogItem "$sPkgName - is installed." $True $False
                    If (!$Uninstall){
                        LogItem "Quiting the script" $True $False
                        QuitScript 0
                    }
                }else{
                    LogItem "$sPkgName - is NOT Installed" $True $False
                   $ProductInstall = $false
                }
            }else{LogItem "ProductFile - Not Found" $True $False }
        }else{LogItem "ProductCode - Not Found" $True $False }
    }else{LogItem "AuditKey - Not Found" $True $False }
    return $ProductInstall
}catch{LogItem "Error occurred at IsProductInstalled $($_.exception.message)" $True $False}
}


#========================================================================
# Install + Uninstall Routines
#========================================================================

function InstallMSI{
param($sFileName,$sParameters)
try{
	$subroutine = "Install MSI: "

	LogItem "$subroutine Started" $True $False

	$sFilePath = "$sScriptDir\$sFileName"
	
	$sCmd = "$Quotes$sWinDir\System32\MsiExec.exe$Quotes /I $Quotes$sFilePath$Quotes $sParameters"
	
	LogItem "$subroutine About to execute $sCmd" $True $False
	
	$iRetVal = $oShell.Run($sCmd,0,$True)
	
	switch($iRetVal){
	{($_ -eq 0) -or ($_ -eq 3010)}{
		#Success!
		LogItem "$subroutine Install was successful.  (Returned $iRetVal)" $True $True}
	1618{
		LogItem "$subroutine Install was unsuccessful (Returned 1618 -- Another installation is in progress)" $True $true
		QuitScript 1618}
	default{
		#Failure
		LogItem "$subroutineInstall was unsuccessful.  (Returned $iRetVal)" $True $True
		QuitScript $iRetVal }
	}
	LogItem "$subroutine Process finished. (Returned $iRetVal)" $true $true
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occurred at InstallMSI $($_.exception.message)" $True $False}
}

function InstallMSP{
param($sFileName,$sParameters)
try{
	$subroutine = "Install MSP: " 

	LogItem "$subroutine Started" $true $False

	$sFilePath = "$sScriptDir\$sFileName"

	$sCmd = "$Quotes$sWinDir\System32\MsiExec.exe$Quotes /p $Quotes$sFilePath$Quotes $sParameters"

   	LogItem "$subroutine About to execute $sCmd" $true $False
   	
   	$iRetVal = $oShell.Run($sCmd,0, $true)
	switch($iRetVal){
	{($_ -eq 0) -or ($_ -eq 3010)}{
		#Success!
		LogItem "$subroutine Install was successful.  (Returned $iRetVal)" $true $true}
	1618{
		LogItem "$subroutine Install was unsuccessful (Returned 1618 -- Another installation is in progress)" $True $true
		QuitScript 1618 }
	default{
		#Failure
		LogItem "$subroutine Install was unsuccessful.  (Returned $iRetVal)" $true $true
		QuitScript $iRetVal }
	}
	LogItem "$subroutine Process finished. (Returned $iRetVal)" $true $true
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occurred at InstallMSP $($_.exception.message)" $True $False}
}

function UninstallMSI{
param($sGUID,$sParameters)
try{
	$subroutine = "Uninstall MSI: "
	
	LogItem "$subroutine Started" $true $False
	
	$sCmd = "$Quotes$sWinDir\System32\MsiExec.exe$Quotes /X $sGUID $sParameters"
   	
   	LogItem "$subroutine About to execute $sCmd" $True $False

	$iRetVal = $oShell.Run($sCmd,0, $true)
	switch($iRetVal){
	{($_ -eq 0) -or ($_ -eq 3010)}{
			#Success!
			LogItem "$subroutine Uninstall was successful.  (Returned $iRetVal)" $true $true}
    1618{
			LogItem "$subroutine Uninstall was unsuccessful (Returned 1618 -- Another installation is in progress)" $True $true
			QuitScript 1618 }
	default{
	
		#Failure
		LogItem "$subroutine Uninstall was unsuccessful.  (Returned $iRetVal )" $true $true
		QuitScript $iRetVal}
	}
	
	LogItem "$subroutine Process finished. (Returned $iRetVal )" $True $False
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occurred at UninstallMSI $($_.exception.message)" $True $False}
}

function Install{
param($sFilePath,$sParameters)
try{
	$sFilePath  | out-null
	$sParameters  | out-null
	$subroutine = "Install Setup: " 
	
	LogItem "$subroutine Started" $true $False
		
	$sCmd =  "$Quotes$sFilePath$Quotes $sParameters"
   	#$sCmd
   	LogItem "$subroutine Running: $sCmd"  $true $False
	
	$iRetVal = $oShell.Run($sCmd,0, $true)
	LogItem "$subroutine Process finished. (Returned $iRetVal )" $true $true
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occurred at Install $($_.exception.message)" $True $False}
}

function Uninstall{
param($sFilePath,$sParameters)
try{
	$sFilePath  | out-null
	$sParameters  | out-null
	$subroutine = "Uninstall Setup: " 

	LogItem "$subroutine Started" $true $False

	$sCmd = "$Quotes$sFilePath$Quotes $sParameters"
	#$sCmd
   	LogItem "$subroutine Running: $sCmd" $true $False
   	
	$iRetVal = $oShell.Run($sCmd,0, $true)
	
	LogItem "$subroutine Process finished. (Returned $iRetVal )" $true $False
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occurred at Uninstall $($_.exception.message)" $True $False}
}

function ExecuteCMD{
param($sFilePath,$sParameters)
try{
	$subroutine = "Execute Command: " 

	LogItem "$subroutine Started" $true $False

	$sCmd = "$Quotes$sFilePath$Quotes $sParameters"
	
   	LogItem "$subroutine Running: $sCmd" $true $False
   	
	$iRetVal = $oShell.Run($sCmd,0, $False)
	
	LogItem "$subroutine Process finished. (Returned $iRetVal)" $true $False
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occurred at ExecuteCMD $($_.exception.message)" $True $False}
}

#========================================================================
# Registry Routines
#========================================================================

Function IsRegKeyExist{
param($sRootKey,$subKey)
try{
	$sKeyName = "$sRootKey\$subKey"
    $iRetVal = $oShell.Run("REG QUERY $Quotes$sKeyName$Quotes",0, $true)
    If ($iRetVal -ne 0) {
        return $false
    }Else{
        return $true
    }
}catch{LogItem "Error occurred at IsRegKeyExist $($_.exception.message)" $True $False ; return $false}
}

Function IsRegValNameExist{
param($sRootKey,$subKey,$sValueName)
try{
	$sKeyName = "$sRootKey\$subKey"
    $iRetVal = $oShell.Run("REG QUERY `"$sKeyName`" /v `"$sValueName`"",0, $true)
	#$iRetVal
    If ($iRetVal -ne 0) {
        return $false
    }Else{
        return $true
    }
}catch{LogItem "Error occurred at IsRegValNameExist $($_.exception.message)" $True $False; return $false}
}


<#Function IsRegValNameExist{
param($sRootKey,$subKey,$sValueName)
try{
	$sKeyName = "$sRootKey\$subKey"
    $iRetVal = Start-Process -FilePath REG.exe -ArgumentList "QUERY `"$sKeyName`" /v `"$sValueName`"" -Wait -PassThru #$oShell.Run("REG QUERY `"$sKeyName`" /v `"$sValueName`"",0, $true)
	$iRetVal = $iRetVal.ExitCode
    If ($iRetVal -ne 0) {
        return $false
    }Else{
        return $true
    }
}catch{LogItem "Error occurred at IsRegValNameExist $($_.exception.message)" $True $False; return $false}
}#>

function CreateKey{
param($sRootKey,$subKey)
try{
	$sKeyName = "$sRootKey\$subKey"
    If (-Not (IsRegKeyExist $sRootKey $subKey) ) {
        LogItem "About to create: $Quotes$sKeyName$Quotes" $true $False
        $iRetVal = $oShell.Run("REG ADD $Quotes$sKeyName$Quotes /f",0, $true)
        If ($iRetVal -ne 0) {
            LogItem "$Quotes$sKeyName$Quotes was not created" $true $False
        }Else{
            LogItem "$Quotes$sKeyName$Quotes has been created" $true $False
        }
    }<#Else{
        LogItem "$Quotes$sKeyName$Quotes already exists" $true $False
    }#>
}catch{LogItem "Error occurred at CreateKey $($_.exception.message)" $True $False}
}

function DeleteKey{
param($sRootKey,$subKey)
try{
	$sKeyName,$iRetVal
	$sKeyName = $sRootKey + "\" + $subKey
	LogItem ("About to delete: " + $Quotes + $sKeyName + $Quotes) $true $False
    If ((IsRegKeyExist $sRootKey $subKey)) {
        $iRetVal = $oShell.Run(("REG DELETE $Quotes$sKeyName$Quotes /f"),0, $true)
        If ($iRetVal -ne 0) {
            LogItem ("$Quotes$sKeyName$Quotes was not deleted") $true $False
        }Else{
            LogItem ("$Quotes$sKeyName$Quotes has been deleted") $true $False
        }
    }Else{
        LogItem ("$Quotes$sKeyName$Quotes does not exist.") $true $False
    }
}catch{LogItem "Error occurred at DeleteKey $($_.exception.message)" $True $False}
}

function SetRegVal{
param($sRootKey,$subKey,$sValueName,$sDataType,$sValue)
try{
    $sKeyName = $sRootKey + "\" + $subKey
    
    If (-Not (IsRegKeyExist $sRootKey $subKey)) {
        CreateKey $sRootKey $subKey
    }
    
    if($svalue.Length -ge 2){If ($sValue.SubString($sValue.Length-1,1) -eq "\") {
    	$sValue = $sValue + "\"
    }}

    $iRetVal = $oShell.Run("REG ADD $Quotes$sKeyName$Quotes /v $Quotes$sValueName$Quotes /t $sDataType /d $Quotes$sValue$Quotes /f",0, $true)
    If ($iRetVal -ne 0) {
        LogItem "The value of $Quotes$sValueName$Quotes under $sKeyName was not set to $Quotes$sValue$Quotes as $sDataType" $true $False
        LogItem "The process returned: $iRetVal" $true $False
    }Else{
        LogItem "The value of $Quotes$sValueName$Quotes under $sKeyName was set to $Quotes$sValue$Quotes as $sDataType" $true $False
    }
}catch{LogItem "Error occurred at SetRegVal $($_.exception.message)" $True $False}
}

function DelRegValName{
param($sRootKey,$subKey,$sValueName)
try{
	$iRetVal=$sKeyName=$null
	$sKeyName = $sRootKey + "\" + $subKey
    If ((IsRegValNameExist $sRootKey $subKey $sValueName)) {
        LogItem "About to delete $Quotes$sValueName$Quotes under $Quotes$sRootKey\$subKey$Quotes" $true $False
        $iRetVal = $oShell.Run("REG DELETE `"$sKeyName`" /v $sValueName /f",0, $true)
        If ($iRetVal -ne 0) {
            LogItem "$Quotes$sKeyName $sValueName$Quotes was not deleted/found" $true $False
        }Else{
            LogItem "$Quotes$sKeyName $sValueName$Quotes has been deleted" $true $False
        }
    }Else{
        LogItem "DelRegValname - $Quotes$sKeyName $sValueName$Quotes does not exist." $true $False
    }
}catch{LogItem "Error occurred at DelRegValName $($_.exception.message)" $True $False}
}

#========================================================================
# Check Pending Reboot Routines
#========================================================================
Function PendingRebootCheck_Reg{
try{
	If (IsRegKeyExist "HKLM" "SYSTEM\CurrentControlSet\services\SNAC") {
		LogItem "Previous Package - Symantec Endpoint Protection - has uninstalled." $true $False
		LogItem "System REBOOT is required, before installing new version." $true $False
		LogItem "Script will EXIT now without attempting to install new version." $true $False
		QuitScript 3010
	}Else{
		LogItem "Script will proceed with installing new version." $true $False
	}	 
}catch{LogItem "Error occurred at PendingRebootCheck_Reg $($_.exception.message)" $True $False}
}

function PendingRebootCheck_Process{
param($sProcess)
try{
	LogItem "Process Status Check: Started" $true $False
	$oProcesses = Get-WmiObject -Class Win32_Process
	
	ForEach($oProcess in $oProcesses){
		If ($sProcess -ieq $oProcess.Name) {
			LogItem "Process Status Check: Found $sProcessis running." $true $False
			LogItem " System Reboot required " $true $False
			LogItem "Process Status Check: Finished" $true $False
			QuitScript 3010
		}
	}
	LogItem "Process Status Check: Finished" $true $False
}catch{LogItem "Error occurred at PendingRebootCheck_Process $($_.exception.message)" $True $False}
}

#========================================================================
# Process and Services Manipulation Workers
#========================================================================
function TaskKill{
param($sProcess)
try{
       LogItem "Task Kill: Started" $true $False
       $oProcesses = Get-WmiObject -Class Win32_Process
       
       $bRunning = $False   

       ForEach ($oProcess In $oProcesses){
              If ($sProcess -ieq $oProcess.Name) {
            try{
                         $bRunning = $True
                Get-Process -Id $($oProcess.ProcessId) | Stop-Process -Force -Confirm:$false
                LogItem "Task Kill: Found $sProcess ($($oProcess.ProcessId)) (ParentId: $($oProcess.ParentProcessId)) and terminated" $true $False
            } 
            catch{LogItem "Task Kill: Unable to terminate the $sProcess ($($oProcess.ProcessId)) (ParentId: $($oProcess.ParentProcessId)) $($_.exception.message)" $true $False}
              }
    }

       If ($bRunning) {
              # Wait and make sure the process is terminated.
              LogItem "Task Kill: Validating if process is terminated" $true $False
              while(-Not $bRunning){
                     $oProcesses = Get-WmiObject -Query "Select * from Win32_Process where Name = '$sProcess'" -ErrorAction SilentlyContinue #Get-WmiObject -Class Win32_Process -Filter "Name -eq $sProcess" -ErrorAction SilentlyContinue
                     #Wait for 100 MilliSeconds
                     Start-Sleep -Milliseconds 100
                     #If no more processes are running, exit Loop
                     If ($oProcesses.Count -eq 0) { 
                           $bRunning = $False
                     }
              }
       }Else{
              LogItem "Task Kill: $sProcess Is not running" $true $False
       }
       LogItem "Task Kill: Finished" $true $False
}catch{LogItem "Error occurred at TaskKill $($_.exception.message)" $True $False}
} 


function WaitForProcess{
param($sProcessName)
try{
    [array]$oProcess = @()
	Do{
		$oProcess = Get-WmiObject -Query "Select * from Win32_Process where Name = '$sProcessName'" -ErrorAction SilentlyContinue
		#-Class Win32_Process -Filter "Name -eq $sProcessName" -ErrorAction SilentlyContinue
		Start-Sleep -Milliseconds 30000
    }while($oProcess.Count -ne 0)
}catch{LogItem "Error occurred at WaitForProcess $($_.exception.message)" $True $False}
}

function WaitAndKill{
param($sProcessName)
try{
	[array]$oProcess = @()

	Do{
		$oProcess = Get-WmiObject -Query "Select * from Win32_Process where Name = '$sProcessName'" -ErrorAction SilentlyContinue
		#Get-WmiObject -Class Win32_Process -Filter "Name -eq $sProcessName" -ErrorAction SilentlyContinue
		If ($oProcess.Count -ne 0) {
			TaskKill $sProcessName
		}
		Start-Sleep -Milliseconds 10000
	}while($oProcess.Count -ne 0)
}catch{LogItem "Error occurred at WaitAndKill $($_.exception.message)" $True $False}
}

function WaitForProcessToFinish{
param($sProcess)
try{
	LogItem "Wait For Process To Finish: $sProcess - Started" $true $False
	$oProcesses = Get-WmiObject -Class Win32_Process
	
	$bRunning = $False	
	
	ForEach ($oProcess in $oProcesses){
		If ($sProcess -ieq $oProcess.Name) {
			$bRunning = $True						
		}
	}

	If ($bRunning) {
		# Wait and make sure the process is terminated.
		LogItem "Wait For Process To Finish: $sProcess - Validating if process has finished." $true $False
		while(-not $bRunning){
			$oProcesses = Get-WmiObject -Query "Select * from Win32_Process where Name = '$sProcess'" -ErrorAction SilentlyContinue 
			#Get-WmiObject -Class Win32_Process -Filter "Name -eq $sProcess" -ErrorAction SilentlyContinue
			#Wait for 100 MilliSeconds
			Start-Sleep -Milliseconds 10000
			#If no more processes are running, exit Loop
			If ($oProcesses.Count -eq 0) { 
				$bRunning = $False
			}
		}
	}Else{
		LogItem "Wait For Process To Finish: $sProcess - Is not running" $true $False
	}
	LogItem "Wait For Process To Finish: $sProcess - Finished" $true $False
}catch{LogItem "Error occurred at WaitForProcessToFinish $($_.exception.message)" $True $False}
}

function SystemLockStatus{
try{
	$isLocked = $null ; $isLocked = Get-Process -ProcessName LogonUI -ErrorAction SilentlyContinue
	if($isLocked -ne $null){$script:sSystemLockStatus = "Locked"}
    Else{$script:sSystemLockStatus = "Unlocked" }
}catch{LogItem "Error occurred at SystemLockStatus $($_.exception.message)" $True $False}
}

#========================================================================
# Service Manipulation Workers
#========================================================================

function StopService{
param($sService)
try{
  LogItem "Stop Service: Started" $true $False
    $sCmd = Get-Service -Name $sService -ErrorAction SilentlyContinue
    LogItem "Stop Service: About to run: $sService" $true $False
    
    if($sCmd -ne $null){
        $sCmd.Stop()
        LogItem "Stop Service: $sService - Stopped" $true $False
    }else{LogItem "Stop Service: $sService not found" $true $False}
    LogItem "Stop Service: Finished" $true $False
}catch{LogItem "Error occurred at StopService $($_.exception.message)" $True $False}
}

function StartService{
param($sService)
try{
 	LogItem "Start Service: Started" $true $False
    $sCmd = Get-Service -Name $sService -ErrorAction SilentlyContinue
    LogItem "Start Service: About to run: $sService" $true $False
    
    if($sCmd -ne $null){
        $sCmd.Start()
        LogItem "Start Service: $sService - Started" $true $False
    }else{LogItem "Start Service: $sService not found" $true $False}
    LogItem "Start Service: Finished" $true $False
}catch{LogItem "Error occurred at StartService $($_.exception.message)" $True $False}
}

function RestartService{
param($sService)
try{
	LogItem "Restart Service: Started" $true $False
	StopService $sService
	StartService $sService
	LogItem "Restart Service: Finished" $true $False
}catch{LogItem "Error occurred at RestartService $($_.exception.message)" $True $False}
} 

function ConfigService{
param($sService,$sState)
try{
 	LogItem "Configure Service: Started" $true $False
    $sCmd = "sc config $sService start= $sState"
    LogItem "Configure Service: About to configure $sService to start as $sState" $true $False

        $iRetVal = $oShell.Run($sCmd,0, $true)
        If ($iRetVal -ne 0) {
            LogItem "Configure Service: $sService Exit Code: $iRetVal Error Number: Err.number" $true $False
        }Else{
            LogItem "Configure Service: $sService Exit Code: $iRetVal Return Number: $iRetVal" $true $False
        }

    LogItem "Configure Service: Finished" $true $False
}catch{LogItem "Error occurred at ConfigService $($_.exception.message)" $True $False}
}

function Delservice{
param($sService)
try{
  LogItem "Delete Service: Started" $true $False
    $sCmd = "sc delete $sService"

        LogItem "Delete Service: About to run: $sCmd" $true $False
        $iRetVal = $oShell.Run($sCmd,0, $true)
        If ($iRetVal -ne 0) {
            LogItem "Delete Service: $sService Exit Code: $iRetVal Error Number: $error[0]" $true $False
        }Else{
            LogItem "Delete Service: $sService Exit Code: $iRetVal Return Number: $error[0]" $true $False
        }

    LogItem "Delete Service: Finished" $true $False
}catch{LogItem "Error occurred at DelService $($_.exception.message)" $True $False}
}

#========================================================================
# File and Folder Workers
#========================================================================
Function FileExist{
param($sFilePath)
try{
            $oFSO = New-Object -ComObject "Scripting.FileSystemObject"
            If ($oFSO.FileExists($sFilePath)) {
                        return $true
            }Else{
                        return $False
            }
}catch{LogItem "Error occurred at FileExist $($_.exception.message)" $True $False ; return $false}
}

function CreateFile{
param($sFile)
try{
            $oFSO = New-Object -ComObject "Scripting.FileSystemObject"
            If (-NOT $oFSO.FileExists($sFile)) {
    $oFile = $oFSO.CreateTextFile($sFile)
            LogItem "$sFile created successfully" $true $False
    }Else{
                        LogItem "$sFile Already exist." $true $False
            }
}catch{LogItem "Error occurred at CreateFile $($_.exception.message)" $True $False}
}

function CopyFile{
param($sSourceFile, $sDest)
try{
            $oFSO = New-Object -ComObject "Scripting.FileSystemObject"
    If ($sDest.SubString($sDest.Length-1,1) -eq "\") {
                        CreateFolderIfNeeded (Split-Path -Path $sDest)
                        $sFile = Split-Path -Path $sSourceFile -Leaf
                        If ($oFSO.fileExists("$sDest$sFile")) {
                                    If ($bFileOverwrite) {
                                                $oFile = $oFSO.GetFile("$sDest$sFile")
                                                $oFile.Attributes = 0
                                                $oFile = $null
                                    }
                        }
            }Else{
                        CreateFolderIfNeeded (Split-Path -Path $sDest)
                        If ($oFSO.FileExists($sDest)) {
                                    If ($bFileOverwrite) {
                                                $oFile = $oFSO.GetFile($sDest)
                                                $oFile.Attributes = 0
                                                $oFile = $null
                                    }
                        }
            }
                        #$sSourceFile
            #$oFSO.CopyFile( "$quotes$sSourceFile$quotes", "$quotes$sDest$quotes", $bFileOverwrite)
			copy-item  "$sSourceFile" "$sDest" -force
}catch{LogItem "Error occurred at CopyFile $($_.exception.message)" $True $False}
}

function CopyFolder{
param($sSourceFolder, $sDestFolder)
try{
            $oFSO = New-Object -ComObject "Scripting.FileSystemObject"
            CreateFolderIfNeeded $sDestFolder
            If ($sDestFolder.SubString($sDestFolder.Length-1,1) -ne "\") {
                        $sDestFolder = "$sDestFolder\"
            }
            LogItem "CopyFolder $sDestFolder" $True $False
            $oFolder = $oFSO.GetFolder($sSourceFolder)
            ForEach ($oFile In $oFolder.Files){
                        #Write-Host $($oFile.Path) -back Green
                        LogItem "CopyFolder: Copying File $($oFile.Path)" $True $False
                        CopyFile $oFile.Path $sDestFolder
            }

            ForEach ($ofunctionFolder In $oFolder.SubFolders){
                        CopyFolder $ofunctionFolder.Path ("$sDestFolder{0}" -f $ofunctionFolder.Name)
            }
}catch{LogItem "Error occurred at CopyFolder $($_.exception.message)" $True $False}
}

function CreateFolderIfNeeded{
param($sTarget)
try{
            $oFSO = New-Object -ComObject "Scripting.FileSystemObject"
            $sFolder = $aFolders = $sWorkingFolder = $null ; $i = 1
            $sFolder = $sTarget
            #LogItem "CreateFolderIfNeeded: Folder requested to be created $sFolder" $True $False
            If ($sFolder.Substring($sFolder.Length-1,1) -eq "\"){
                        $sFolder = $sFolder.Substring(0,$sFolder.Length-1)
            }
            [array]$aFolders = $sFolder.Split("\")
            #write-host "Split Folders" -back yellow
            [bool]$created=$false
            $sWorkingFolder = $aFolders[0]
            foreach($i in $aFolders){ #.GetUpperBound(0)){
                        if($sWorkingFolder -ne $i){
                        $sWorkingFolder = "$sWorkingFolder\$i" #$($aFolders[$i])"
                        #LogItem "CreateFolderIfNeeded: Checking if folder is present $sWorkingFolder" $True $False
                        If ($oFSO.FolderExists($sWorkingFolder) -eq $false){
                                    #LogItem "Create Folder if Needed: Creating Folder $Quotes $sWorkingFolder $Quotes" $true  $false
                                    $oFSO.CreateFolder($sWorkingFolder) | Out-Null
                                    $created=$true
                        } 
						}
            }
            if($created){LogItem "CreateFolderIfNeeded: Folder created $sFolder" $True $False}
}catch{LogItem "Error occurred at CreateFolderIfNeeded $($_.exception.message)" $True $False}
}

function DeleteFile{
param($sFile)
try{
	LogItem "About to delete file $sFile." $true $False
	if((Test-Path -Path "$sFile") -ne $true){
		LogItem "$sFile does not exist." $true $False
	}
	else{
		remove-item "$sFile" -force
	}
}catch{LogItem "Error occurred at DeleteFile $($_.exception.message)" $True $False}
}

function MoveFile{
param($sSource,$sDest)
try{
		if((Test-path "$sSource") -ne $true){
			CreateFolderIfNeeded $sDest
			Move-Item -Path "$sSource" -Destination "$sDest"
		}
		else
		{
			LogItem "$sSource does not exist." $true $False
		}

}catch{LogItem "Error occurred at MoveFile $($_.exception.message)" $True $False}
}

function DeleteFolder{
param($sFolder)
try{
	LogItem "About to delete folder $sFolder." $true $False
            	if((Test-Path -Path "$sFolder") -ne $true){
						LogItem "$sFolder does not exist." $true $False
            }Else{
					remove-item "$sFolder" -recurse -force -confirm:$False
                        LogItem "$sFolder deleted." $true $False
            }
}catch{LogItem "Error occurred at DeleteFolder $($_.exception.message)" $True $False}
}

 function DeleteFileFromProfile{
 param($sFilePath)
 try{
	$sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	$afunctionKeys = Get-ChildItem "HKLM:\$sKeyPath"
	$oFSO = New-Object -ComObject "Scripting.FileSystemObject"
	LogItem "About to delete $sFilePath." $True $False
	$sProfile=$sProfileName=$null
	ForEach ($ofunctionKey In $afunctionKeys){
	
		$sValueName=$sValue=$sfunctionPath=$null
		$sValueName = "ProfileImagePath"
		$sfunctionPath = "$sKeyPath\$(Split-Path -Path $ofunctionKey.Name -Leaf)"
		$sValue = (Get-ItemProperty "HKLM:\$sfunctionPath" -Name $sValueName).$sValueName

		$sProfile = $sValue.split("\")
        $sProfileName = $sProfile[2]
		
		# filter out unnecessary profiles 
	    If (($sProfileName -ne "config") -And ($sProfileName -ne "system32") -And ($sProfileName -ne "ServiceProfiles") -And ($sProfileName -ne "UpdatusUser") -And ($sProfileName -ne "Administrator") -And ($sProfileName -ne ("Administrator."  + $oNetwork.ComputerName)) -And ($sProfileName -ne "z_cseappinstall") -And ($sProfileName -ne "ctx_cpsvcuser") -And ($sProfileName -ne "MsDtsServer110") -And ($sProfileName -ne "ReportServer") -And ($sProfileName -ne "MSSQLFDLauncher") -And ($sProfileName -ne "SQLSERVERAGENT") -And ($sProfileName -ne "MSSQLSERVER") -And ($sProfileName -ne "QBDataServiceUser26") -And ($sProfileName -ne "CitrixTelemetryService")) {
	    	#$oFolder = $oFSO.GetFolder($sValue)
			$sPath =  "$sValue\$sFilePath"
			LogItem "About to delete $sPath." $True $False
			If ($oFSO.FileExists($sPath)) {
				DeleteFile $sPath
			}
		}
	}
}catch{LogItem "Error occurred at DeleteFileFromProfile $($_.exception.message)" $True $False}
}

function DeleteFolderFromProfile{
param($sFolderPath)
try{
	$sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	$afunctionKeys = Get-ChildItem "HKLM:\$sKeyPath"
	$oFSO = New-Object -ComObject "Scripting.FileSystemObject"
	$sProfile=$sProfileName=$null
	ForEach ($ofunctionKey In $afunctionKeys){
		$sValueName=$sValue=$sfunctionPath=$null
		$sValueName = "ProfileImagePath"
		$sfunctionPath = "$sKeyPath\$(Split-Path -Path $ofunctionKey.Name -Leaf)"
		$sValue = (Get-ItemProperty "HKLM:\$sfunctionPath" -Name $sValueName).$sValueName
		
		$sProfile = $sValue.split("\")
        $sProfileName = $sProfile[2]
		
	    # filter out unnecessary profiles 
	    If (($sProfileName -ne "config") -And ($sProfileName -ne "system32") -And ($sProfileName -ne "ServiceProfiles") -And ($sProfileName -ne "UpdatusUser") -And ($sProfileName -ne "Administrator") -And ($sProfileName -ne ("Administrator."  + $oNetwork.ComputerName)) -And ($sProfileName -ne "z_cseappinstall") -And ($sProfileName -ne "ctx_cpsvcuser") -And ($sProfileName -ne "MsDtsServer110") -And ($sProfileName -ne "ReportServer") -And ($sProfileName -ne "MSSQLFDLauncher") -And ($sProfileName -ne "SQLSERVERAGENT") -And ($sProfileName -ne "MSSQLSERVER") -And ($sProfileName -ne "QBDataServiceUser26") -And ($sProfileName -ne "CitrixTelemetryService")) {
	    	#$oFolder = $oFSO.GetFolder($sValue)
			$sPath =  "$sValue\$sFolderPath"
			If ($oFSO.FolderExists($sPath)) {
				DeleteFolder $sPath
			}
		}
	}
}catch{LogItem "Error occurred at DeleteFolderFromProfile $($_.exception.message)" $True $False}
}

function CopyFolderToProfile{
param($sSourceFolder,$sTargetFolder)
try{
	$sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	$afunctionKeys = Get-ChildItem "HKLM:\$sKeyPath"
	$oFSO = New-Object -ComObject "Scripting.FileSystemObject"
	$sProfile = $sProfileName = $null
	ForEach ($ofunctionKey In $afunctionKeys){
		$sValueName=$sValue=$sfunctionPath=$null
		$sValueName = "ProfileImagePath"
		$sfunctionPath = "$sKeyPath\$(Split-Path -Path $ofunctionKey.Name -Leaf)"
		$sValue = (Get-ItemProperty "HKLM:\$sfunctionPath" -Name $sValueName).$sValueName
		
		$sProfile = $sValue.split("\")
        $sProfileName = $sProfile[2]
		
	    # filter out unnecessary profiles 
	    If (($sProfileName -ne "config") -And ($sProfileName -ne "system32") -And ($sProfileName -ne "ServiceProfiles") -And ($sProfileName -ne "UpdatusUser") -And ($sProfileName -ne "Administrator") -And ($sProfileName -ne ("Administrator."  + $oNetwork.ComputerName)) -And ($sProfileName -ne "z_cseappinstall") -And ($sProfileName -ne "ctx_cpsvcuser") -And ($sProfileName -ne "MsDtsServer110") -And ($sProfileName -ne "ReportServer") -And ($sProfileName -ne "MSSQLFDLauncher") -And ($sProfileName -ne "SQLSERVERAGENT") -And ($sProfileName -ne "MSSQLSERVER") -And ($sProfileName -ne "QBDataServiceUser26") -And ($sProfileName -ne "CitrixTelemetryService")) {
			If ($oFSO.FolderExists($sSourceFolder)) {
				$sRobocopyCmd = "Robocopy.exe $Quotes$sSourceFolder$Quotes $Quotes$sValue\$sTargetFolder$Quotes /E /Z"
				LogItem "About to Execute: $sRobocopyCmd" $true $False
				$iRetVal = $oShell.Run($sRobocopyCmd,0, $true)
				switch($iRetVal){
					0{
						LogItem "No files were copied. No failure was encountered. No files were mismatched. The files already exist in the destination directory; therefore, the copy operation was skipped." $true $False
					}1{
						LogItem "All files were copied successfully." $true $False
					}2{
						LogItem "There are some additional files in the destination directory that are not present in the source directory. No files were copied." $true $False
					}3{
						LogItem "Some files were copied. Additional files were present. No failure was encountered." $true $False 
					}5{
						LogItem "Some files were copied. Some files were mismatched. No failure was encountered." $true $False
					}6{
						LogItem "Additional files and mismatched files exist. No files were copied and no failures were encountered. This means that the files already exist in the destination directory." $true $False
					}7{
						LogItem "Files were copied, a file mismatch was present, and additional files were present." $true $true
						QuitScript 1603
                    }8{
						LogItem "Several files did not copy." $true $true
						QuitScript 1603
                    }default{
						LogItem "Unknown exception. Error number was Err.Number and the description was: $($error[0])" $true $true
						QuitScript 1603
                    }
				}
			}Else{
				LogItem "Source folder does not exist. Please validate $sSourceFolder path exists." $true $False
				QuitScript 1603
			}
		}
	}
}catch{LogItem "Error occurred at CopyFolderToProfile $($_.exception.message)" $True $False}
}

function CopyFileToProfile{
param($sSourceFolder,$sTargetFolder,$sSourceFileName)
try{
	$sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	$afunctionKeys = Get-ChildItem "HKLM:\$sKeyPath"
	$oFSO = New-Object -ComObject "Scripting.FileSystemObject"
	$sProfile=$sProfileName=$null
	ForEach ($ofunctionKey In $afunctionKeys){
		$sValueName=$sValue=$sfunctionPath=$null
		$sValueName = "ProfileImagePath"
		$sfunctionPath = "$sKeyPath\$(Split-Path -Path $ofunctionKey.Name -Leaf)"
		$sValue = (Get-ItemProperty "HKLM:\$sfunctionPath" -Name $sValueName).$sValueName
		
		$sProfile = $sValue.split("\")
        $sProfileName = $sProfile[2]
		
	    # filter out unnecessary profiles 
	    If (($sProfileName -ne "config") -And ($sProfileName -ne "system32") -And ($sProfileName -ne "ServiceProfiles") -And ($sProfileName -ne "UpdatusUser") -And ($sProfileName -ne "Administrator") -And ($sProfileName -ne ("Administrator."  + $oNetwork.ComputerName)) -And ($sProfileName -ne "z_cseappinstall") -And ($sProfileName -ne "ctx_cpsvcuser") -And ($sProfileName -ne "MsDtsServer110") -And ($sProfileName -ne "ReportServer") -And ($sProfileName -ne "MSSQLFDLauncher") -And ($sProfileName -ne "SQLSERVERAGENT") -And ($sProfileName -ne "MSSQLSERVER") -And ($sProfileName -ne "QBDataServiceUser26") -And ($sProfileName -ne "CitrixTelemetryService")) {
			If ($oFSO.FolderExists($sSourceFolder)) {
				$sRobocopyCmd = "Robocopy.exe $Quotes$sSourceFolder$Quotes $Quotes$sValue\$sTargetFolder$Quotes $Quotes$sSourceFileName$Quotes /E /Z"
				LogItem "About to Execute: $sRobocopyCmd" $true $False
				$iRetVal = $oShell.Run($sRobocopyCmd,0, $true)
				switch ($iRetVal){
					0{
						LogItem "No files were copied. No failure was encountered. No files were mismatched. The files already exist in the destination directory; therefore, the copy operation was skipped." $true $False
					} 1{
						LogItem "All files were copied successfully." $true $False
					} 2{
						LogItem "There are some additional files in the destination directory that are not present in the source directory. No files were copied." $true $False
					} 3{
						LogItem "Some files were copied. Additional files were present. No failure was encountered." $true $False 
					} 5{
						LogItem "Some files were copied. Some files were mismatched. No failure was encountered." $true $False
					} 6{
						LogItem "Additional files and mismatched files exist. No files were copied and no failures were encountered. This means that the files already exist in the destination directory." $true $False
					}7{
						LogItem "Files were copied, a file mismatch was present, and additional files were present." $true $true
						QuitScript 1603
					}8{
						LogItem "Several files did not copy." $true $true
						QuitScript 1603 
					}default{
						LogItem "Unknown exception. Error number was Err.Number and the description was: $($error[0])" $true $true
						QuitScript 1603}
				}
			}Else{
				LogItem "Source folder does not exist. Please validate $sSourceFolder path exists." $true $False
				QuitScript 1603
			}
		}
	}
}catch{LogItem "Error occurred at CopyFileToProfile $($_.exception.message)" $True $False}
}

#========================================================================
# Install Validation and Tattoo Workers
#========================================================================

function ValidateInstall{
try{	
	$sProductFilePath_update = ($sProductFilePath).Replace("\","\\") 
    $sInstalledFileVer =(Get-WmiObject -Class CIM_DataFile -Filter "Name='$sProductFilePath_update'" | Select-Object Version).Version
    LogItem "Validate Install: Product File Version: $sProductFileVersion"
	
	If ($sInstalledFileVer -eq $sProductFileVersion){
		LogItem "Validate Install: Installed Target File: $sInstalledFileVer" $true $False 
    	LogItem "Validate Install: Able to validate Installed Target File Version with the given Product File Version." $true $False 
	}Else{
		LogItem "Validate Install: Unable to validate Target File Version." $true $False 
		QuitScript 1603
	}
}catch{LogItem "Error occurred at ValidateInstall $($_.exception.message)" $True $False ; QuitScript 1603}
}

function ValidateInstall_DateModified_BaseApp{
try{
	$oFile = $oFSO.GetFile($sProductFilePath)

	$sInstalledFile_DateModified_BaseApp = $oFSO.DateLastModified($oFile)
	
	If ($sInstalledFile_DateModified_BaseApp -eq "10/31/2012 3:09 PM") { 
		LogItem "Base Application has installed: Control-M/Enterprise Manager 8.0.00" $true $False
	}Else{
		LogItem "Base Application has NOT installed: Control-M/Enterprise Manager 8.0.00" $true $False
		QuitScript 1603
	}
}catch{LogItem "Error occurred at Validate_DateModified_BaseApp $($_.exception.message)" $True $False ; QuitScript 1603}
}

function ValidateInstall_DateModified_Patch{
try{
	$oFile = $oFSO.GetFile($sProductFilePath)

	$sInstalledFile_DateModified_Patch = $oFSO.DateLastModified($oFile)
	
	If ($sInstalledFile_DateModified_Patch -eq "5/14/2015 1:10 AM") { 
		LogItem "Patch Application has installed: Control-M/Enterprise Manager 8.0.00 Fix Pack 7 (Default)" $true $False
	}Else{
		LogItem "Patch Application has NOT installed: Control-M/Enterprise Manager 8.0.00 Fix Pack 7 (Default)" $true $False
		QuitScript 1603
	}
}catch{LogItem "Error occurred at ValidateInstall_DateModified_Patch $($_.exception.message)" $True $False ; QuitScript 1603}
}

function ValidateInstall_File{
try{
	LogItem "Validate Install: Started" $true $False
	#Validate installation
	If ($oFSO.FileExists($sProductFilePath)) { 
		LogItem "Validate Install: Able to validate Target File existence. Target File path is: $sProductFilePath" $true $False 
	}Else{
		LogItem "Validate Install: Unable to validate Target File." $true $False
		QuitScript 1603
	}
	LogItem "Validate Install: Finished" $true $False
}catch{LogItem "Error occurred at ValidateInstall_File $($_.exception.message)" $True $False ; QuitScript 1603}
}

function ValidateUninstall_File{
try{
	LogItem "Validate Uninstall: Started" $true $False
	#Validate Uninstallation
	If (-Not ($oFSO.FileExists($sProductFilePath))) { 
		LogItem "Validate Uninstall: Able to validate non-existence of Target File." $true $False 
	}Else{
		LogItem "Validate Uninstall: Target File still exits at: $sProductFilePath" $true $False
		QuitScript 1603
	}
	LogItem "Validate Uninstall: Finished" $true $False
}catch{LogItem "Error occurred at ValidateUninstall_File $($_.exception.message)" $True $False ; QuitScript 1603}
}

Function GetRegPath{
try{
	LogItem "GetRegPath: checking for HKLM\32or64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode" $true $false
	If ((IsRegKeyExist "HKLM" ("$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode"))) {
		#LogItem "GetRegPath: found HKLM\$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode - setting sRegPath" $true $false
		$script:sRegPath = "$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode"
		$script:sTattooPath = "$sRegKeyRoot64\$sRegAuditPath"
	}ElseIf ((IsRegKeyExist "HKLM" ("$sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode"))) {	
	#LogItem "GetRegPath: found $sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode - setting sRegPath" $true $false
		$script:sRegPath = "$sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode"
		$script:sTattooPath = "$sRegKeyRoot32\$sRegAuditPath"
	}else{LogItem "GetRegPath: uninstall\$sProductCode -keynot found - sregpath not set" $true $false}
	
}catch{LogItem "Error occurred at GetRegPath $($_.exception.message)" $True $False}
}

function TattooRegistry{
try{
	LogItem "TattooRegistry: process Started" $true $false
    SetRegVal "HKLM" "$sTattooPath\$sPkgName" "Install Date" "REG_SZ" $sInstallDate
    SetRegVal "HKLM" "$sTattooPath\$sPkgName" "Vendor Name" "REG_SZ" $sProductVendor
    SetRegVal "HKLM" "$sTattooPath\$sPkgName" "Product Version" "REG_SZ" $sProductVersion
    SetRegVal "HKLM" "$sTattooPath\$sPkgName" "Product Code" "REG_SZ" $sProductCode
    SetRegVal "HKLM" "$sTattooPath\$sPkgName" "Target Directory" "REG_SZ" $sInstallFolderPath
    SetRegVal "HKLM" "$sTattooPath\$sPkgName" "Install Source" "REG_SZ" $sScriptDir
}catch{LogItem "Error occurred at TatooRegistry $($_.exception.message)" $True $False}
}

function RemoveTattoo{
try{
	LogItem "Remove TattooRegistry: process Started" $true $false
	DeleteKey "HKLM" ("$sTattooPath\$sPkgName")
}catch{LogItem "Error occurred at Remove Tattoo $($_.exception.message)" $True $False}
}

function ARPCustomization{
try{
	LogItem "$sRegPath" $true $false
	if($sRegPath -eq ""){exit}
   # SetRegVal "HKLM" $sRegPath "DisplayName" "REG_SZ" $sPkgName
    SetRegVal "HKLM" $sRegPath "NoModify" "REG_DWORD" "1"
    SetRegVal "HKLM" $sRegPath "NoRepair" "REG_DWORD" "0"
    SetRegVal "HKLM" $sRegPath "Comments" "REG_SZ" "Script Package"
    SetRegVal "HKLM" $sRegPath "Contact" "REG_SZ" ""
    SetRegVal "HKLM" $sRegPath "HelpLink" "REG_SZ" ""
    SetRegVal "HKLM" $sRegPath "Readme" "REG_EXPAND_SZ" ""
    SetRegVal "HKLM" $sRegPath "URLUpdateInfo" "REG_SZ" ""
    SetRegVal "HKLM" $sRegPath "URLInfoAbout" "REG_SZ" ""
}catch{LogItem "Error occurred at ARPCustomization $($_.exception.message)" $True $False}
}

Function TattooInstallFolder{
param($sInstallDir)
try{
	$oFSO.CreateTextFile($sInstallDir)
}catch{LogItem "Error occurred at TattooInstallFolder $($_.exception.message)" $True $False}
}

#=======================================================================
# Per-User Registry Workers
#========================================================================

function AddToUserHives{
param([parameter(mandatory=$true,position=1)][string]$sKeyPath,
    [parameter(mandatory=$true,position=2)][string]$sValueName,
    [parameter(mandatory=$true,position=3)]
    [ValidateSet('REG_SZ','REG_EXPAND_SZ','REG_BINARY','REG_DWORD','REG_QWORD','REG_MULTI_SZ')]
    [string]$sDataType,
    [parameter(mandatory=$true,position=4)][AllowEmptyString()][string]$sValue)
    try{
		$tempfile = $([System.IO.Path]::GetTempFileName()) 
		$tempEfile = $([System.IO.Path]::GetTempFileName()) 
		try{ 
$error.clear()		
		Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG LOAD HKEY_USERS\CUSTOM $($env:SystemDrive)\Users\Default\NTUSER.DAT" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
        #LogItem ("AddToUserHives - DefaultHive LOAD Command completed with result {0}" -f $($tmp = "" ; gc $tempfile | % { if([string]::IsNullOrEmpty($_) -ne $true){$tmp = "$tmp $_"} } ; $tmp) ) $true $false
		$null = Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG ADD `"HKU\CUSTOM\$sKeyPath`" /v `"$sValueName`" /t `"$sDataType`" /d `"$sValue`" /f" -Wait -NoNewWindow -RedirectStandardOutput $tempfile -RedirectStandardError $tempEfile -passthru
        
		[string]$tmp = "" ; gc $tempfile, $tempEfile -ea 0 | % { if([string]::IsNullOrEmpty($_) -ne $true){$tmp = "$tmp $_"} }
		LogItem ("AddToUserHives - DefaultHive ADD Command REG ADD `"HKU\CUSTOM\$sKeyPath`" /v `"$sValueName`" /t `"$sDataType`" /d `"$sValue`" /f completed with result $tmp") $true $false
		Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG UNLOAD HKEY_USERS\CUSTOM" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
        #LogItem ("AddToUserHives - DefaultHive UNLOAD Command completed with result {0}" -f $($tmp = "" ; gc $tempfile | % { if([string]::IsNullOrEmpty($_) -ne $true){$tmp = "$tmp $_"} } ; $tmp) ) $true $false
		}catch{LogItem "AddToUserHives - Unable to add Reg Value to the default hive (new user) Error occurred $($_.exception.message)" $true $false}

        $RegValueType = ''
        switch ($sDataType.ToUpper()) { 
            'REG_SZ' { 
                $RegValueType = 'String' 
            } 
            'REG_DWORD' { 
                $RegValueType = 'Dword'
            } 
            'REG_BINARY' { 
                $RegValueType = 'Binary' 
            } 
            'REG_EXPAND_SZ' { 
                $RegValueType = 'ExpandString'
            } 
            'REG_MULTI_SZ' { 
                $RegValueType = 'MultiString'
            } 
            default { 
                throw "Registry type '$sDataType' not recognized" 
                LogItem "AddToUserHives - Registry type '$sDataType' not recognized" $true $false
            } 
        } 

        New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null 
        
		# Get each user profile SID and Path to the profile
        [Array]$UserProfiles = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" | `
         Where-Object {($_.PSChildName -match "S-1-5-21-(\d+-?){4}$") -and ($excludeUserProfile -inotcontains (Split-Path $_.ProfileImagePath -Leaf))} | ` # -and ($AllUsersList.SID -inotmatch $_.PSChildName)
         Select-Object @{Name="SID"; Expression={$_.PSChildName}}, @{Name="UserHive";Expression={"$($_.ProfileImagePath)\NTuser.dat"}}
		
        # Loop through each profile on the machine
        Foreach ($UserProfile in $UserProfiles) {
            # Load User ntuser.dat if it's not already loaded
            If (($ProfileWasLoaded = Test-Path Registry::HKEY_USERS\$($UserProfile.SID)) -eq $false) {
                Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG.EXE LOAD HKU\$($UserProfile.SID) $($UserProfile.UserHive)" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
                LogItem ("AddToUserHives - UserHive {0} Load command completed with result {1}" -f $($UserProfile.SID), $($tmp = "" ; gc $tempfile | % { if([string]::IsNullOrEmpty($_) -ne $true){$tmp = "$tmp $_"} } ; $tmp) ) $true $false
            }
            # Manipulate the registry
            $key = "Registry::HKEY_USERS\$($UserProfile.SID)\$sKeyPath"
            
            if((Test-path -Path "$key") -eq $false){
    		    ## Create the key path if it doesn't exist
                New-Item -Path "Registry::HKEY_USERS\$($UserProfile.SID)\$($sKeyPath | Split-Path -Parent)" -Name $($sKeyPath | Split-Path -Leaf) -Force -ErrorAction Stop | Out-Null
            }
            ## Create (or modify) the value specified in the param 
            New-ItemProperty -Path "Registry::HKEY_USERS\$($UserProfile.SID)\$sKeyPath" -Name $sValueName -Value $sValue -PropertyType $RegValueType -Force -ErrorAction Stop | Out-Null
            LogItem ("AddToUserHives - REG $sKeyPath -Name $sValueName -Value $sValue -PropertyType $sDataType added for the user {0}" -f $(($UserProfile.UserHive).Split("\")[2])) $true $false

            # Unload NTuser.dat        
            If ($ProfileWasLoaded -eq $false) {
                [gc]::Collect()
                Start-Sleep 1
                Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG.EXE UNLOAD HKU\$($UserProfile.SID)" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
                LogItem ("AddToUserHives - UserHive {0} UnLoad command completed with result {1}" -f $($UserProfile.SID), $($tmp = "" ; gc $tempfile | % { if([string]::IsNullOrEmpty($_) -ne $true){$tmp = "$tmp $_"} } ; $tmp) ) $true $false
            }
        }
        Remove-Item $tempfile -ErrorAction SilentlyContinue -Force
    }
    catch{LogItem "AddToUserHives - Error occurred $($_.exception.message)" $true $false}
}

function DeleteFromUserHives{
    param([parameter(mandatory=$true,position=1)][string]$sKeyPath)
    try{
		$tempfile = $([System.IO.Path]::GetTempFileName()) 
		$tempEfile = $([System.IO.Path]::GetTempFileName()) 
		try{  
		Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG LOAD HKEY_USERS\CUSTOM $($env:SystemDrive)\Users\Default\NTUSER.DAT" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
		Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG DELETE `"HKU\CUSTOM\$sKeyPath`" /f" -Wait -NoNewWindow -RedirectStandardOutput $tempfile -RedirectStandardError $tempEfile
        [string]$tmp = "" ; gc $tempfile, $tempEfile | % { if([string]::IsNullOrEmpty($_) -ne $true){$tmp = "$tmp $_"} }
		if(($tmp | Out-String) -imatch "ERROR: The system was unable to find the specified registry key or value."){ 
			LogItem "DeleteFromUserHives - DefaultHive `"HKU\CUSTOM\$sKeyPath`" - Key doesn't exist" $true $false
		}
		else{
			LogItem "DeleteFromUserHives - DefaultHive Delete Command REG DELETE `"HKU\CUSTOM\$sKeyPath`" /f completed with result $tmp" $true $false
		}		
		Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG UNLOAD HKEY_USERS\CUSTOM" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
		}catch{LogItem "DeleteFromUserHives - Unable to delete Reg Value from the default hive (new user) Error occurred $($_.exception.message)" $true 	$false}	
        New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null 
        
                # Get each user profile SID and Path to the profile
        [array]$UserProfiles = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" | `
         Where-Object {($_.PSChildName -match "S-1-5-21-(\d+-?){4}$") -and ($excludeUserProfile -inotcontains (Split-Path $_.ProfileImagePath -Leaf))} | ` #-and ($AllUsersList.SID -inotmatch $_.PSChildName) 
         Select-Object @{Name="SID"; Expression={$_.PSChildName}}, @{Name="UserHive";Expression={"$($_.ProfileImagePath)\NTuser.dat"}}

        # Loop through each profile on the machine
        Foreach ($UserProfile in $UserProfiles) {
            # Load User ntuser.dat if it's not already loaded
            If (($ProfileWasLoaded = Test-Path Registry::HKEY_USERS\$($UserProfile.SID)) -eq $false) {
                Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG.EXE LOAD HKU\$($UserProfile.SID) $($UserProfile.UserHive)" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
            }
            # Manipulate the registry
            $key = "Registry::HKEY_USERS\$($UserProfile.SID)\$sKeyPath"
            
            if(Test-path -Path "$key"){
                Remove-Item -Path $key -Force -Recurse -Confirm:$false -erroraction stop
				[string]$tmp = ($UserProfile.UserHive).Split("\")[2]
				LogItem "DeleteFromUserHives - REG $sKeyPath Path deleted for the user $tmp" $true $false
            }
			
            # Unload NTuser.dat        
            If ($ProfileWasLoaded -eq $false) {
                [gc]::Collect()
                Start-Sleep 1
                Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG.EXE UNLOAD HKU\$($UserProfile.SID)" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
            }
        }
		Remove-Item $tempfile, $tempEfile -ErrorAction SilentlyContinue -Force
    }
    catch{LogItem "DeleteFromUserHives - Error occurred $($_.exception.message)" $true $false}
}

function DelFromUserHivesRegKeyValue{
    param([parameter(mandatory=$true,position=1)][string]$sKeyPath,
			[parameter(mandatory=$true,position=2)][string]$Value)
    try{
		$tempfile = $([System.IO.Path]::GetTempFileName()) 
		$tempEfile = $([System.IO.Path]::GetTempFileName()) 
		try{  
		Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG LOAD HKEY_USERS\CUSTOM $($env:SystemDrive)\Users\Default\NTUSER.DAT" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
		Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG DELETE `"HKU\CUSTOM\$sKeyPath`" /f" -Wait -NoNewWindow -RedirectStandardOutput $tempfile -RedirectStandardError $tempEfile
        [string]$tmp = "" ; gc $tempfile, $tempEfile | % { if([string]::IsNullOrEmpty($_) -ne $true){$tmp = "$tmp $_"} }
		if(($tmp | Out-String) -imatch "ERROR: The system was unable to find the specified registry key or value."){ 
			LogItem "DelFromUserHivesRegKeyValue - DefaultHive `"HKU\CUSTOM\$sKeyPath`" - Key doesn't exist" $true $false
		}
		else{
			LogItem "DelFromUserHivesRegKeyValue - DefaultHive Delete Command REG DELETE `"HKU\CUSTOM\$sKeyPath`" /f completed with result $tmp" $true $false
		}	
		Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG UNLOAD HKEY_USERS\CUSTOM" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
		}catch{LogItem "DelFromUserHivesRegKeyValue - Unable to delete Reg Value from the default hive (new user) Error occurred $($_.exception.message)" $true 	$false}	
        New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null 
        
                # Get each user profile SID and Path to the profile
        [array]$UserProfiles = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" | `
         Where-Object {($_.PSChildName -match "S-1-5-21-(\d+-?){4}$") -and ($excludeUserProfile -inotcontains (Split-Path $_.ProfileImagePath -Leaf))} | ` #-and ($AllUsersList.SID -inotmatch $_.PSChildName) 
         Select-Object @{Name="SID"; Expression={$_.PSChildName}}, @{Name="UserHive";Expression={"$($_.ProfileImagePath)\NTuser.dat"}}

        # Loop through each profile on the machine
        Foreach ($UserProfile in $UserProfiles) {
            # Load User ntuser.dat if it's not already loaded
            If (($ProfileWasLoaded = Test-Path Registry::HKEY_USERS\$($UserProfile.SID)) -eq $false) {
                Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG.EXE LOAD HKU\$($UserProfile.SID) $($UserProfile.UserHive)" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
            }
            # Manipulate the registry
            $key = "Registry::HKEY_USERS\$($UserProfile.SID)\$sKeyPath"
            [string]$tmp = ($UserProfile.UserHive).Split("\")[2]
            if(Test-path -Path "$key"){
				if((Get-Item -Path "$key" | Where-Object { $_.Property -ieq $Value }).Count -ge 1){
					Remove-ItemProperty -Path $key -Name $Value -Force -Confirm:$false -erroraction stop
					
					LogItem "DelFromUserHivesRegKeyValue - REG Value $Value in $sKeyPath Path deleted for the user $tmp" $true $false
				}else{
					LogItem "DelFromUserHivesRegKeyValue - REG Value $Value in $sKeyPath Path not found for the user $tmp" $true $false
				}
            }
			
            # Unload NTuser.dat        
            If ($ProfileWasLoaded -eq $false) {
                [gc]::Collect()
                Start-Sleep 1
                Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG.EXE UNLOAD HKU\$($UserProfile.SID)" -Wait -NoNewWindow -RedirectStandardOutput $tempfile
            }
        }
		Remove-Item $tempfile, $tempEfile -ErrorAction SilentlyContinue -Force
    }
    catch{LogItem "DelFromUserHivesRegKeyValue - Error occurred $($_.exception.message)" $true $false}
}

Function MountHive{
param($sHivePath)
try{
	$sCmd = "REG.EXE LOAD HKEY_USERS\CUSTOM $Quotes$sHivePath$Quotes"
	LogItem "About to run: $sCmd" $true $False
	$iRetVal = $oShell.Run($sCmd,0, $true)
    If ($iRetVal -ne 0) {
        LogItem "$Quotes$sHivePath$Quotes is currently in use." $true $False
        return $false
    }Else{
        LogItem "$Quotes$sHivePath$Quotes is now mounted." $true $False
        return $false
    }
}catch{LogItem "Error occurred at MountHive $($_.exception.message)" $True $False}
}

function UnmountHive{
try{
	$sCmd = "REG UNLOAD HKEY_USERS\CUSTOM"
	$iRetVal = $oShell.Run($sCmd,0, $true)
    If ($iRetVal -ne 0) {
        LogItem "Unable to unmount user hive. Exit code: $iRetVal" $true $False
        QuitScript 1603
    }Else{
        LogItem "Unmounted user hive. Exit code: $iRetVal" $true $False
    }
}catch{LogItem "Error occurred at UnmountHive $($_.exception.message)" $True $False}
}

function MountDefaultHive{
try{
	$sCmd = "REG LOAD HKEY_USERS\CUSTOM $Quotes%SYSTEMDRIVE%\Users\Default\NTUSER.DAT$Quotes"
	$iRetVal = $oShell.Run($sCmd,0, $true)
    If ($iRetVal -ne 0) {
        LogItem "Unable to mount default hive. Exit code: $iRetVal" $true $False
        QuitScript 1603
    }Else{
        LogItem "Mounted default hive. Exit code: $iRetVal" $true $False
    }
}catch{LogItem "Error occurred at MountDefaultHive $($_.exception.message)" $True $False}
}

function CreateGuid
{
try{
	return [guid]::NewGuid()
}catch{LogItem "Error occurred at CreateGuid $($_.exception.message)" $True $False}
}

Function CreateShortcuts{
try{
	$oShell = New-Object -com "Wscript.Shell"
	#To create Short at StartMenuPrograms
	$olnk = $oShell.CreateShortcut("$sAllUsersStartPrograms\Control-M Configuration Manager.lnk")   
   	$olnk.TargetPath = "$sProgFiles64\BMC Software\Control-M EM 8.0.00\Default\bin\emccm.exe"
   	$olnk.Arguments = ""
   	$olnk.Description = "Control-M Configuration Manager "
   	$olnk.IconLocation = "$sProgFiles64\BMC Software\Control-M EM 8.0.00\Default\bin\emccm.ico"
   	$olnk.WindowStyle = "1"
   	$olnk.WorkingDirectory = "$sProgFiles64\BMC Software\Control-M EM 8.0.00\Default\bin"
   	$olnk.Save()
      	
   	#To create Short at C:\Program Files\FRB Programs
   	$olnk = $oShell.CreateShortcut("$sProgFiles64\FRB Programs\Control-M Configuration Manager.lnk")   
   	$olnk.TargetPath = "$sProgFiles64\BMC Software\Control-M EM 8.0.00\Default\bin\emccm.exe"
   	$olnk.Arguments = ""
   	$olnk.Description = "Control-M Configuration Manager "
   	$olnk.IconLocation = "$sProgFiles64\BMC Software\Control-M EM 8.0.00\Default\bin\emccm.ico"
   	$olnk.WindowStyle = "1"
   	$olnk.WorkingDirectory = "$sProgFiles64\BMC Software\Control-M EM 8.0.00\Default\bin"
   	$olnk.Save()
}catch{LogItem "Error occurred at CreateShortcuts $($_.exception.message)" $True $False}	
}

function AppendSysPATH{
param($sPathToAdd)
try{
	$oShell = New-Object -ComObject "WScript.Shell"
	$sSysEnv = $oShell.Environment("SYSTEM")	
	$sOldSysPath = $sSysEnv.Item("Path")
	LogItem "Current System PATH Variable value is: $sOldSysPath" $true $False
	LogItem "About to Append this value from System PATH Variable: $sPathToAdd" $true $False										
	$sNewSysPath = "$sPathToAdd;$sOldSysPath"
	$sSysEnv.Item("Path") = $sNewSysPath
	LogItem "After Appending System PATH Variable value is: $($sSysEnv.Item('Path'))" $true $False	
}catch{LogItem "Error occurred at AppenSysPath $($_.exception.message)" $True $False}	
}

function DelSysPATH{
param($sPathToRemove)
try{
	$oShell = New-Object -ComObject "WScript.Shell"
	$sSysEnv = $oShell.Environment("SYSTEM")	
	$sOldSysPath = $sSysEnv.Item("Path")
	LogItem "Current System PATH Variable value is: $sOldSysPath" $true $False
	LogItem "About to Remove this value from System PATH Variable: $sPathToRemove" $true $False
	$sNewSysPath = $sOldSysPath.ToLower().Replace($sPathToRemove.ToLower(), "")
	$sSysEnv.Item("Path") = $sNewSysPath
	LogItem "After Removing System PATH Variable value is: $($sSysEnv.Item('Path'))" $true $False	
}catch{LogItem "Error occurred at DelSysPath $($_.exception.message)" $True $False}	
}

function UnZip{
		param($sZipFilePath,$sExtractDirPath)
		try{
		$subroutine=$oShellApp=$oFSO=$sFilesInZip=$null
		$subroutine = "UnZip Process: "
		LogItem "$subroutine Started" $true $False
		#ZipFile Path
		LogItem "$subroutine Current location of Zip File: $sZipFilePath" $true $False
		#The folder the contents should be extracted to.
		LogItem "$subroutine Will get Extracted to this directory: $sExtractDirPath" $true $False
		#Extract the contants of the zip file.
		if((Test-path -path "$sZipFilePath") -eq $false){
			LogItem "Zip file not found $sZipFilePath, exiting script" $True $False ; QuitScript 0 
		}
		
		Add-Type -assembly "System.IO.Compression.Filesystem"

		[io.compression.zipfile]::ExtractToDirectory("$sZipFilePath","$sExtractDirPath") 		

		LogItem "$subroutine Finished" $true $False
		
		}catch{LogItem "Error occurred at Unzip $($_.exception.message), exiting script" $True $False ; QuitScript 0 }
} 


 Function Uninstall_Pre_Chrome{
try{
            $oReg=$sKeyPath=$sValueName=$sValue=$sUnInst_CMD=$sUnInstExitCode=$sChromeVer=$sDisplayVer=$null
            #$oReg = GetObject("winmgmts://./root/default:StdRegProv")    
            $oShell = New-Object -ComObject "WScript.Shell"   
            $sKeyPath = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
            $sValueName = "UninstallString"
            #$oReg.GetStringValue($HKEY_LOCAL_MACHINE,$sKeyPath, $sValueName, $sValue)
            $sValue = (Get-ItemProperty -Path "HKLM:\$sKeyPath" -Name $sValueName).$sValueName
            $sUnInst_CMD = "$sValue --force-uninstall"
            $sChromeVer = "DisplayVersion"
            #$oReg.GetStringValue($HKEY_LOCAL_MACHINE, $sKeyPath, $sChromeVer, $sDisplayVer)
            $sDisplayVer = (Get-ItemProperty -Path "HKLM:\$sKeyPath" -Name $sChromeVer).$sChromeVer
            LogItem "Found previous version of Chrome with version: $sDisplayVer" $true $False    
			LogItem "Uninstall Command: $sUnInst_CMD" $true $False
			$sUnInstExitCode = $oShell.Run($sUnInst_CMD,0, $true)
		LogItem "Manually Installed Chrome has uninstalled with an Exitcode: $sUnInstExitCode" $true $False
}catch{LogItem "Error occurred at Uninstall_Pre_Chrome $($_.exception.message)" $True $False}
}   

 Function Cleanup_Chrome_AuditRegKeys{
try{
	[array]$arrAuditRegKeys = @("Google_Chrome_51.0","Google_Chrome_52.0.2743","Google_Chrome_55.0.2883.75","Google_Chrome_57.0.2987.98","Google_Chrome_58.0.3029.96","Google_Chrome_59.0.3071.86","Google_Chrome_60.0.3112.78","Google_Chrome_61.0.3163.100","Google_Chrome_62.0.3202.75","Google_Chrome_62.0.3202.94","Google_Chrome_63.0.3239.132","Google_Chrome_64.0.3282.140","Google_Chrome_64.0.3282.186")
	
	ForEach($arrAuditRegKey In $arrAuditRegKeys){
		If ((IsRegKeyExist "HKLM" "SOFTWARE\Wow6432Node\FRB\Applications\$arrAuditRegKey")){
			LogItem "Found: HKLM\SOFTWARE\Wow6432Node\FRB\Applications\$arrAuditRegKey" $True $False
			DeleteKey "HKLM" "SOFTWARE\Wow6432Node\FRB\Applications\$arrAuditRegKey"	
		}
	}	
}catch{LogItem "Error occured at Cleanup_Chrome_AuditRegKeys $($_.exception.message)" $True $False}
}  

 Function Cleanup_Chrome_PreviousUninstallStrings{
try{
	[array]$arrUninstallRegKeys = @("{328C5246-6FF0-3191-A48F-9F2652D06154}","{6D86A294-19C9-3B64-A92E-E1BD29883C66}","{E551C6E5-FB51-38DA-8400-BF6C8D82394D}","{21E2140F-3DD6-3291-8E0A-DCC94E1D41C4}","{47501057-0AA8-3BFF-A90A-854951421508}","{B969EF2E-1A11-3A73-9FF7-942EBFBBCC05}","{C6CF5A81-401A-3BBD-9475-275DF0401F4D}","{BDC829D3-2606-3EC4-9A6C-94C34AABF2F4}","{DD550B1D-C176-3E11-AB0C-85C8077C86B8}","{A58EE139-F99A-3991-B9D2-EBB6A6E2F9AE}","{125B436B-3F17-317F-8D2F-9C470DC68905}","{076D9EC4-5DF0-3179-AB3E-33D96C705980}","{9BCCDBF4-DB5F-3792-98B0-74100A584B07}")
	
	ForEach($arrUninstallRegKey In $arrUninstallRegKeys){
		If ((IsRegKeyExist "HKLM" "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$arrUninstallRegKey")){
			LogItem "Found: HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$arrUninstallRegKey" $True $False
			DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$arrUninstallRegKey"	
		}
	}	
}catch{LogItem "Error occured at Cleanup_Chrome_PreviousUninstallStrings $($_.exception.message)" $True $False}
}  

function RemoveExcelAddIn{
    param([parameter(mandatory=$true,position=1)][string]$sAddInPath,
	[parameter(mandatory=$true,position=2)][string]$sExcelKeyPath)
	<#AuthorName Thanuj, .NOTES Complete powershell, Date 2:00 PM 9/1/2017#>
    try{
		$sKeyPath = $sExcelKeyPath #"SOFTWARE\Microsoft\Office\14.0\Excel\Options"
				
        $AllUsersList = Get-WmiObject -Query "SELECT * FROM Win32_UserProfile WHERE SPECIAL='false'" | `
        Where-Object { $excludeUserProfile -inotcontains (Split-Path $_.LocalPath -Leaf) }

        New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS | Out-Null 
        
        ## Change the registry values for the currently logged on user. Each logged on user SID is under HKEY_USERS 
        [Array]$LoggedOnSids = (Get-ChildItem HKU: | Where-Object { ($_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$') -and ($AllUsersList.SID -imatch $_.SID) }).PSChildName 
        
        foreach($sid in $LoggedOnSids){
            try{
                foreach($item in (Get-Item -Path "HKU:\$sid\$sKeyPath" -ErrorAction SilentlyContinue | Select-Object Property).Property){
                    try{
						[string]$value = (Get-ItemProperty -Path "HKU:\$sid\$sKeyPath" -Name $item).$item
						if($value.tolower().contains($sAddInPath.tolower()) -eq $true){Remove-ItemProperty -Path "HKU:\$sid\$sKeyPath" -Name $item -Force}
						}
                    catch{}
                }
            }
            catch{LogItem "RemoveExcelAddIn - Unable to add Reg Value for the user $sid Error occurred $($_.exception.message)"  $True $False}
        }     

        # Get each user profile SID and Path to the profile
        $UserProfiles = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" | `
         Where-Object {($_.PSChildName -match "S-1-5-21-(\d+-?){4}$") -and ($AllUsersList.SID -inotmatch $_.PSChildName) -and ($excludeUserProfile -inotcontains (Split-Path $_.ProfileImagePath -Leaf))} | `
         Select-Object @{Name="SID"; Expression={$_.PSChildName}}, @{Name="UserHive";Expression={"$($_.ProfileImagePath)\NTuser.dat"}}

        # Add in the .DEFAULT User Profile
        $DefaultProfile = "" | Select-Object SID, UserHive
        $DefaultProfile.SID = ".DEFAULT"
        $DefaultProfile.Userhive = "C:\Users\Public\NTuser.dat"
        $UserProfiles += $DefaultProfile

        # Loop through each profile on the machine
        Foreach ($UserProfile in $UserProfiles) {
            # Load User ntuser.dat if it's not already loaded
            If (($ProfileWasLoaded = Test-Path Registry::HKEY_USERS\$($UserProfile.SID)) -eq $false) {
                Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG.EXE LOAD HKU\$($UserProfile.SID) $($UserProfile.UserHive)" -Wait -NoNewWindow
            }
            # Manipulate the registry
            $key = "Registry::HKEY_USERS\$($UserProfile.SID)\$sKeyPath"
            
            foreach($item in (Get-Item -Path $key -ErrorAction SilentlyContinue | Select-Object Property).Property){
                [string]$value = (Get-ItemProperty -Path $key -Name $item).$item
				if($value.tolower().contains($sAddInPath.tolower()) -eq $true){Remove-ItemProperty -Path $key -Name $item -Force}
            }

            # Unload NTuser.dat        
            If ($ProfileWasLoaded -eq $false) {
                [gc]::Collect()
                Start-Sleep 1
                Start-Process -FilePath "CMD.EXE" -ArgumentList "/C REG.EXE UNLOAD HKU\$($UserProfile.SID)" -Wait -NoNewWindow| Out-Null
            }
        }
    }
    catch{LogItem "Error occurred at RemoveExcelAddIn $($_.exception.message)" $True $False}
}

    #region UserProfile Exclusion List
    [Array]$script:excludeUserProfile = @("config","system32","ServiceProfiles","UpdatusUser", `
    "Administrator","z_cseappinstall","ctx_cpsvcuser", `
    "MsDtsServer110","ReportServer","MSSQLFDLauncher", `
    "SQLSERVERAGENT","MSSQLSERVER","QBDataServiceUser26","CitrixTelemetryService",
    "CtxAppVCOMAdmin","zz.")
    #endregion UserProfile Exclusion List
 
#Main script starts here
MainScript
