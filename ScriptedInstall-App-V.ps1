param([switch]$Uninstall)

$ErrorActionPreference="Stop"
Trap {"Error: $_"; Break;}
#Set-StrictMode -Version Latest
#Set-StrictMode -Version 2.0

New-Variable -Name sScriptVer -Value "2.0" -Scope Global

New-Variable -Name HKEY_CLASSES_ROOT -Value "&H80000000" -Option Constant -Scope Global
New-Variable -Name HKEY_CURRENT_USER -Value "&H80000001" -Option Constant -Scope Global
New-Variable -Name HKEY_LOCAL_MACHINE -Value "&H80000002" -Option Constant -Scope Global
New-Variable -Name HKEY_USERS -Value "&H80000003" -Option Constant -Scope Global


# Set Registry Constants
# SetCustomizations
$Global:sPkgName = $Global:sProductVendor = $Global:sProductName = $Global:sProductVersion = $Global:sProductCode = $Global:sProductFilePath_PreReq = $null
$Global:sInstallFile = $Global:sInstallFolderPath = $Global:sProductFilePath = $Global:sProductFileVersion = $Global:sScriptFileName = $Global:sSystemLockStatus = $null
$Global:sCompanyName = $Global:sRegAuditPath = $Global:sLogPath = $Global:bFileOverwrite = $null

# Init
$Global:oShell = $Global:oShellApp = $Global:oFSO = $Global:oNetwork = $Global:oSMSClient = $Global:oSysInfo = $null
$Global:oClass = $Global:colSettings = $Global:oOS = $Global:colComputer = $Global:oComputer = $Global:colItems = $Global:oItem = $null
$Global:colTimeZone = $Global:oTimeZone = $Global:oProcessor = $Global:oBIOS = $Global:sOSName = $null
$Global:sProcessArchitectureX86 = $Global:sProcessArchitectureW6432 = $null
$Global:iOSArch = $Global:iScriptArch = $null
$Global:sSysDir32 = $Global:sSysDir64 = $null
$Global:sProgFiles32 = $Global:sProgFiles64 = $null
$Global:sRegKeyRoot32 = $Global:sRegKeyRoot64 = $null
$Global:sAllUsersStartMenu = $Global:sAllUsersStartPrograms = $null
$Global:sAllUsersDesktop = $Global:sAllUsersAppData = $null
$Global:sWinDir = $Global:sSystemDrive = $Global:sTempDir = $Global:sScriptDir = $null
$Global:sUserProfile = $Global:bUninstall = $Global:Quotes = $Global:sInstallDate = $null
$Global:sTempGuid = $Global:sAppData = $null

# Logging
$Global:sVerb = $Global:sLogText = $null

# GetRegPath
$Global:sRegPath = $Global:sTattooPath = $null

# Set Task Sequence Variables
$Global:sTSKeyPath = $Global:sTSValueName = $Global:sTSValue = $Global:sTSName = $Global:sTSNameData = $Global:sBIOSKeyPath = $Global:sBIOReleaseDate = $Global:sBIOSValue = $null

#=============================================
 # App-V Deploy modifications
#=============================================
# Package friendly name
$PackageName = "SystemToolsSoftware_Hyena_12.5_AppV" #Edit Sample Data
# Product Vendor
$ProductVendor = "SystemToolsSoftware" #Edit Sample Data
# Product Name
$ProductName = "Hyena" #Edit Sample Data
# Product Version
$ProductVersion = "12.5" #Edit Sample Data
# PackageID
$PackageID = "7dccfdb7-a61f-48b6-96b7-f93f08540ad5" #Edit Sample Data
 # VersionID
$VersionID = "297107cc-f978-4f3d-a1a1-b1c000a5ba12" #Edit Sample Data
# Icon Location
$Icon = "Root\HYENA.exe.0.ico" #Edit Sample Data
# Target File Name
$TargetFileName = "HYENA.exe"
# Target File Path
$PrgData = $env:ProgramData
$TargetFilePath = "$PrgData\App-V\$PackageID\$VersionID\Root\$TargetFileName"
$TargetDir = "$PrgData\App-V\$PackageID\$VersionID\Root"
#In case of Supersedence, details of previous Package
$PreviousPackageName = "SystemToolsSoftware_Hyena_12.0.1_AppV" #Edit Sample Data
$PreviousPackageID = "E2984E2C-7D1A-4DBD-B717-9F8C169B4656" #Edit Sample Data
$PreviousVersionID = "e6de0d53-e9fd-4760-985c-5023d2c71a6b" #Edit Sample Data
$PreviousTargetFileName = "HYENA.exe" #Edit Sample Data
$PreviousTargetFilePath = "$PrgData\App-V\$PreviousPackageID\$PreviousVersionID\Root"

#=============================================
 # App-V End Deploy modifications
#=============================================

function SetCustomizations{
try{
	   $Global:sPkgName = "Google_Chrome_59.0.3071.86" # Application Friendly Name
       $Global:sProductVendor = "Google, Inc." # Product vendor
       $Global:sProductName = "Chrome"  # Product name
       $Global:sProductVersion = "59.0.3071.86" # Product version
       $Global:sProductCode = "{B969EF2E-1A11-3A73-9FF7-942EBFBBCC05}" # Product code  
       $Global:sInstallFile = "Chrome_59.0\googlechromestandaloneenterprise.msi" # Install file name
       $Global:sInstallFolderPath = "$Global:sProgFiles32\Google\Chrome" # Install folder path
       $Global:sProductFilePath = "$Global:sInstallFolderPath\Application\chrome.exe" # Main application executable
       $Global:sProductFileVersion = "59.0.3071.86" # Main application executable product 
       $Global:sCompanyName = "FRB" # Variable used to create log folder + registry key path
       $Global:sScriptFileName = "Chrome_59.0_ScriptedInstall.ps1" # Variable for VBScript file name
    
    $Global:sRegAuditPath = "$Global:sCompanyName\Applications"
    $Global:sLogPath = "$Global:sSystemDrive\$Global:sCompanyName\Logs"
	$Global:quotes= '"'
    $Global:bFileOverwrite = $true  # Set to $true to overwrite existing files if they exist
	#if((test-path "$Global:sLogPath\$Global:sPkgName-(Script).log")){Remove-Item "$Global:sLogPath\$Global:sPkgName-(Script).log" -Force}
}catch{LogItem "Error occured at SetCustomizations $($_.exception.message)" $True $False}
}

function GetTaskSeqInfo{
try{
	$Global:sTSKeyPath = "SOFTWARE\FRB\OS Deployment"
	$Global:sTSName = "Task Sequence Name"
	$Global:sTSNameData	= (Get-ItemProperty -Path "HKLM:\$($Global:sTSKeyPath)").$Global:sTSName
	$Global:sTSValueName = "Task Sequence Version"
	$Global:sTSValue = (Get-ItemProperty -Path "HKLM:\$($Global:sTSKeyPath)").$Global:sTSValueName
	$Global:sBIOSKeyPath = "HARDWARE\DESCRIPTION\System\BIOS"
	$Global:sBIOReleaseDate = "BIOSReleaseDate"
	$Global:sBIOSValue	= (Get-ItemProperty -Path "HKLM:\$($Global:sBIOSKeyPath)").$Global:sBIOReleaseDate
}catch{LogItem "Error occured at GetTaskSeqInfo $($_.exception.message)" $True $False}
}

#========================================================================
# Main Script Logic FOR MSI/EXE
#========================================================================

function MainScript{
	#Comment APPV or MSI/EXE accordingly using symbol "#" in the front
	#MainRoutine_MSI_EXE
	MainRoutine_APPV
}

function MainRoutine_MSI_EXE{
    Init
    SetCustomizations
    GetTaskSeqInfo
    SystemLockStatus
    BeginLog
    If ($bUninstall){
           If (-not $(IsProductInstalled)){
                  LogItem "Deployment script will now exit!" $True $False
                  QuitScript 0
           }
			Else{
                  # Uninstall Logic
            GetRegPath
            TaskKill "chrome.exe"
            LogItem "Deployment script will now proceed with $sVerb." $True $False
            UninstallMSI $sProductCode "/qb-! REBOOT=ReallySuppress /l*xv $Quotes$sLogPath\$($sPkgName)_(MSI)_Uninstall.log$Quotes"
                  ValidateUninstall_File
                  Start-Sleep 30000
                  #To Delete file Chromex64
                  #DeleteFile "$sProgFiles32\Google\Chrome\Application\Chromex64.txt"     
                  #Deleting leftover shortcut(s)
                  DeleteFile "$sAllUsersStartPrograms\Google Chrome.lnk"
                  DeleteFile "$sProgFiles32\FRB Programs\Google Chrome.lnk"
                  DeleteFile "$sSystemDrive\Users\Public\Desktop\Google Chrome.lnk"             
                  DeleteFolder $sInstallFolderPath
                  DeleteFolder "$sProgFiles32\Google\CrashReports" 
                  DeleteFolderIfEmpty "$sProgFiles32\Google"
              
                  DeleteFolder "$sProgFiles32\Google\Update"
                  $oShell.Run("cmd /c RD /S /Q $Quotes$sProgFiles32\Google\Update$Quotes")
                  Uninstall "$sSysDir64\schtasks.exe", "/delete /f /tn $QuotesGoogleUpdateTaskMachineCore$Quotes"
                  Uninstall "$sSysDir64\schtasks.exe", "/delete /f /tn $QuotesGoogleUpdateTaskMachineUA$Quotes"
              
                  # Deleting Google Update Services
                  Delservice "gupdate"
                  Delservice "gupdatem"
              
                  # Cleanup per-user folder
                  DelUserProfileDir "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Google Chrome"
                  #DeleteFolderFromProfile("AppData\Local\Google\Chrome")
                  #DeleteFolder sSystemDrive\Users\Default\AppData\Local\Google\Chrome"
                  RemoveTattoo
				  DeleteKey "HKLM" "SOFTWARE\Policies\Google\Chrome"
				  DeleteKey "HKLM" "SOFTWARE\Policies\Google\Update"        
				  DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Policies\Google\Chrome"
                  DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Policies\Google\Update"
                  DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Google\Update"
				  DeleteKey "HKLM" "SOFTWARE\Classes\Installer\Products\931EE85AA99F19939B2DBE6B6A2E9FEA"
				  DeleteKey "HKLM" "$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode"
           }
    }
	Else{ 
		If (IsProductInstalled) {
                  LogItem "Deployment script will now exit!" $True $False
                  QuitScript 0
           }
		 else{
			
                 TaskKill "chrome.exe"
                  If ((IsRegKeyExist "HKLM" "$sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome")){
                        LogItem "Found manually installed Google Chrome on this system. Attempting to uninstall it." $True $False              
                        Uninstall_Pre_Chrome
                        DelUserProfileDir "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Google Chrome"
                        DeleteFolder "$sAllUsersStartPrograms\Google Chrome"
                        DeleteFolder $sInstallFolderPath
                        DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"
                        TaskKill "iexplore.exe"
						DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Google\Chrome"
						DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Google\Update"
						DeleteKey "HKLM" "SOFTWARE\Policies\Google\Chrome"
						DeleteKey "HKLM" "SOFTWARE\Policies\Google\Update"
						DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Policies\Google\Chrome"
						DeleteKey "HKLM" "SOFTWARE\Wow6432Node\Policies\Google\Update"
						DeleteKey "HKLM" "SOFTWARE\Classes\Installer\Products\104B7C566883E1C30AE7319D24995B77"
						DeleteKey "HKLM" "SOFTWARE\Classes\Installer\Products\B634B52171F3F713D8F2C974D06C9850"
						DeleteKey "HKLM" "SOFTWARE\FRB\Applications\Google_Chrome_55.0.2883.87"
						DeleteKey "HKLM" "SOFTWARE\Wow6432Node\FRB\Applications\Google_Chrome_55.0.2883.87"
						DelUserProfileDir "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Google Chrome"
						DeleteFile "$sSystemDrive\Users\Public\Desktop\Google Chrome.lnk"
						DeleteFolder $sInstallFolderPath
						DeleteFolder "$sProgFiles32\Google\CrashReports"
						DeleteFromUserHives "Software\Google\Chrome"
						DeleteFromUserHives "Software\Google\Software Removal Tool"
						DeleteFromUserHives "Software\Google\Update"
                         QuitScript 1641
                  }
				  Cleanup_Chrome_AuditRegKeys
                  $sChromeInst_CMD, $sInstExitCode
                  $sChromeInst_CMD = "$Quotes$sScriptDir\$sInstallFile$Quotes /qb-! ALLUSERS=1 REBOOT=ReallySuppress /l*xv $Quotes$sLogPath\$($sPkgName)_(MSI)_Install.log$Quotes"
                  LogItem ("Install Command for - $sPkgName - is: $sChromeInst_CMD") $True $False
                  LogItem ("Installation has started for: $sPkgName") $True $False
                  $sInstExitCode = $oShell.Run($sChromeInst_CMD,0,$True)
                  LogItem "Installed - $sPkgName - with an Exitcode: $sInstExitCode" $True $False
                  If ($sInstExitCode -eq "1603") {
                         LogItem "Reboot required before installing new version of Google Chrome." $True $False
                         QuitScript 1641
                  }
              
                  ValidateInstall
                  Start-Sleep -Milliseconds 30000          
				 
				  GetRegPath
              
                  #To create file Chromex64
                  #$CreateFile "$sProgFiles32\Google\Chrome\Application\Chromex64.txt"
              
				  #To copy shortcut
                  CopyFile "$sAllUsersStartPrograms\Google Chrome\Google Chrome.lnk" "$sAllUsersStartPrograms\Google Chrome.lnk"
                  CopyFile "$sAllUsersStartPrograms\Google Chrome.lnk" "$sProgFiles32\FRB Programs\Google Chrome.lnk"
              
                  #To copy preferences file
                  CopyFile "$sScriptDir\Chrome_59.0\master_preferences" "$sInstallFolderPath\Application\master_preferences"
                  LogItem "Copying First Run file to [Userprofile] to avoid second StartMenu Shortcut Creation" $True $False
                  #CopyFileToProfile sScriptDir\Chrome_57.0","AppData\Local\Google\Chrome\User Data","First Run"
                  CopyFolderToProfile "$sScriptDir\Chrome_59.0\User Data" "AppData\Local\Google\Chrome\User Data"
                  CopyFolder "$sScriptDir\Chrome_59.0\User Data" "$sSystemDrive\Users\Default\AppData\Local\Google\Chrome\User Data"
                           
                  #To delete unwanted shortcut
                  DeleteFile "$sAllUsersDesktop\Google Chrome.lnk"
                  DeleteFolder "$sAllUsersStartPrograms\Google Chrome"
                  DeleteFile "$sSystemDrive\Users\Public\Desktop\Google Chrome.lnk"
              
                  #Applying Registry Keys settings, to customize Chrome
                  Install "$sWinDir\regedit.exe" "/s $Quotes$sScriptDir\Chrome_59.0\Chrome_CustomSettings.reg$Quotes"
              
                  #Executing GPUpdate
                  ExecuteCMD "$sWinDir\SysWOW64\gpupdate.exe" "/force /wait:0"
              
                  #To disable Google Update Schedule Tasks
                  Install "$sSysDir64\schtasks.exe" "/change /disable /tn $QuotesGoogleUpdateTaskMachineCore$Quotes"
                  Install "$sSysDir64\schtasks.exe" "/change /disable /tn $QuotesGoogleUpdateTaskMachineUA$Quotes"
              
                  #To disable Google Update Services
                  ConfigService "gupdate" "disabled"
                  ConfigService "gupdatem" "disabled"
				
                  #Applying Registry Keys settings, to disable Google Update
                  SetRegVal "HKLM" "$sRegKeyRoot64\Policies\Google\Update" "AutoUpdateCheckPeriodMinutes" "REG_DWORD" "0"
                  SetRegVal "HKLM" "$sRegKeyRoot64\Policies\Google\Update" "UpdateDefault" "REG_DWORD" "0"
              
                  DeleteFolder "$sProgFiles32\Google\Update"
                  $oShell.Run("cmd /c RD /S /Q $Quotes$sProgFiles32\Google\Update$Quotes")
              
                  #Applying Registry Keys settings, to customize ARP Settings
                  ARPCustomization
                  SetRegVal "HKLM" "SOFTWARE\Classes\Installer\Products\E2FE969B11A137A3F97F49E2FBBBCC50" "ProductName" "REG_SZ" $sPkgName         
                  SetRegVal "HKLM" "$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode" "DisplayIcon" "REG_SZ" $sProductFilePath
                  SetRegVal "HKLM" "$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode" "DisplayVersion" "REG_SZ" $sProductVersion
                  SetRegVal "HKLM" "$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode" "UninstallString" "REG_SZ" "cscript.exe \$Quotes$sScriptDir\$sScriptFileName\$Quotes /Uninstall"
              
                  TattooRegistry   
           }
    #}

    #Success (No Reboot)
    #QuitScript(0)

    # Success (No Reboot) 
    #QuitScript(1707)

    #Soft Reboot 
    #QuitScript(3010)

    #Hard Reboot
    QuitScript 1641

    #Force Reboot
    ##ForceReboot(1641)
}
catch
{
    LogItem "Error occured at main script $($_.exception.message)" $True $False
}
}

#========================================================================
# Main Script Logic FOR App-V
#========================================================================

Function MainRoutine_APPV
{
    #  Beginning Main Logic
    BeginLog

    If (!$uninstall)
    {
        # Install Logic
        TaskKill "HYENA.exe"
        TaskKill "stexport.exe" 
        TaskKill "stuc.exe"
        RemovePreviousPackage
        DeleteFile "$ProgramData\Microsoft\Windows\Start Menu\Google Earth.lnk"
        DeployPackage
        VerifyInstall_Appv "$TargetFilePath"
        Tatoo_Appv
        ARPCustomization_APPV
        #AddRunVirtual("EXCEL.EXE") 
        #AddActiveSetup 
    }
    Else
    {
        # Uninstall Logic
        TaskKill "HYENA.exe"
        TaskKill "stexport.exe" 
        TaskKill "stuc.exe"      
        RemovePackage
        VerifyUninstall_Appv "$TargetFilePath"
        RemoveTatoo_Appv
        RemoveARPCustomization_AppV
        #RemoveRunVirtual("EXCEL.EXE")  
        #RemoveActiveSetup 
    }
}

<#
################################################################################
App-V Functions
################################################################################
#>

Function DeployPackage
{
    # Import App-V Module
    LogItem "Executing: Import-Module $($ProgFiles64)\Microsoft Application Virtualization\Client\AppvClient\AppvClient.psd1"
    $null = Import-Module "$($ProgFiles64)\Microsoft Application Virtualization\Client\AppvClient\AppvClient.psd1"
    
    # Publish App-V package to current user
    # Remove App-V publishing server

    LogItem "Executing: Get-AppVPublishingServer | Remove-AppVPublishingServer"
    $null = Get-AppVPublishingServer | Remove-AppVPublishingServer

    # Import App-V package into store
    LogItem "Executing: Add-AppvClientPackage -Path $($ScriptPath)\$($ProductName)\$($PackageName).appv" -DynamicDeploymentConfiguration "$($ScriptPath)\$($ProductName)\$($PackageName)_DeploymentConfig.xml"
    $null = Add-AppvClientPackage -Path "$($ScriptPath)\$($ProductName)\$($PackageName).appv" -DynamicDeploymentConfiguration "$($ScriptPath)\$($ProductName)\$($PackageName)_DeploymentConfig.xml"
        
    # Publish App-V package to user
    LogItem "Executing:  Publish-AppvClientPackage -PackageID $($PackageID) -VersionID $($VersionID)  -DynamicUserConfigurationPath $($ScriptPath)\$($ProductName)\$($PackageName)_UserConfig.xml -ev err"
    $null = Publish-AppvClientPackage -PackageID $($PackageID) -VersionID $($VersionID) -DynamicUserConfigurationPath "$($ScriptPath)\$($ProductName)\$($PackageName)_UserConfig.xml" -ev err

    If (($err -ne $null) -and ($error[0].Exception.AppvWarningCode -eq 8589935887)) 
    {
        $host.SetShouldExit(4736)
    }
    
    # Publish App-V package to system
    # Remove App-V publishing server
    LogItem "Executing: Get-AppVPublishingServer | Remove-AppVPublishingServer"
    $null = Get-AppVPublishingServer | Remove-AppVPublishingServer

    # Load app-v package into app-v client cache
    LogItem "Executing: Mount-AppvClientPackage -PackageID $($PackageID) -VersionID $($VersionID)"
    $null = Mount-AppvClientPackage -PackageID $($PackageID) -VersionID $($VersionID)

    # Publish extension points
    LogItem "Executing: Publish-AppvClientPackage -PackageID $PackageID -VersionID $VersionID -Global -ev err"
    $null = Publish-AppvClientPackage -PackageID $PackageID -VersionID $VersionID -Global -ev err
    if (($err -ne $null) -and ($error[0].Exception.AppvWarningCode -eq 8589935887))
    {
        $host.SetShouldExit(4736)
    }
}

Function RemovePackage
{
    # Import App-V Module
    LogItem "Executing: Import-Module $($ProgFiles64)\Microsoft Application Virtualization\Client\AppvClient\AppvClient.psd1"
    $null = Import-Module "$($ProgFiles64)\Microsoft Application Virtualization\Client\AppvClient\AppvClient.psd1"

    # Get App-V package
    LogItem "Executing: Get-AppvClientApplication -Name $($PackageName)"
    $null = Get-AppvClientApplication -Name $($PackageName)

    # Stop App-V package
    LogItem "Executing: Stop-AppvClientPackage -PackageId $($PackageID) -VersionId $($VersionID) -Global"
    $null = Stop-AppvClientPackage -PackageID $($PackageID) -VersionId $($VersionID) -Global

    # Unpublish App-V package globally
    LogItem "Executing: Unpublish-AppvClientPackage -PackageID $($PackageID) -VersionID $($VersionID) -Global"
    $null = Unpublish-AppvClientPackage -PackageID $($PackageID) -VersionID $($VersionID) -Global
    
    # Remove App-V package globally
    LogItem "Executing: Remove-AppvClientPackage -PackageID $($PackageID) -VersionID $($VersionID) -ev err "
    $null = Remove-AppvClientPackage -PackageID $($PackageID) -VersionID $($VersionID) -ev err 
    If (($err -ne $null) -and ($error[0].Exception.AppvErrorCode -eq 0x0C8007012f))
    {
        $host.SetShouldExit(0)
    }
}

Function RemovePreviousPackage
{
    LogItem "Checking if previous package $PreviousPackageName exists"
    If (Test-Path $PreviousTargetFilePath)
    {  
        LogItem "Previous package $PreviousPackageName found, removing now"
        # Import App-V Module
        LogItem "Executing: Import-Module $($ProgFiles64)\Microsoft Application Virtualization\Client\AppvClient\AppvClient.psd1"
        $null = Import-Module "$($ProgFiles64)\Microsoft Application Virtualization\Client\AppvClient\AppvClient.psd1"

        # Get App-V package
        LogItem "Executing: Get-AppvClientApplication -Name $($PreviousPackageName)"
        $null = Get-AppvClientApplication -Name $($PreviousPackageName)

        # Stop App-V package
        LogItem "Executing: Stop-AppvClientPackage -PackageId $($PreviousPackageID) -VersionId $($PreviousVersionID) -Global"
        $null = Stop-AppvClientPackage -PackageID $($PreviousPackageID) -VersionId $($PreviousVersionID) -Global

        # Unpublish App-V package globally
        LogItem "Executing: Unpublish-AppvClientPackage -PackageID $($PreviousPackageID) -VersionID $($PreviousVersionID) -Global"
        $null = Unpublish-AppvClientPackage -PackageID $($PreviousPackageID) -VersionID $($PreviousVersionID) -Global
    
        # Remove App-V package globally
        LogItem "Executing: Remove-AppvClientPackage -PackageID $($PreviousPackageID) -VersionID $($PreviousVersionID) -ev err "
        $null = Remove-AppvClientPackage -PackageID $($PreviousPackageID) -VersionID $($PreviousVersionID) -ev err 
        If (($err -ne $null) -and ($error[0].Exception.AppvErrorCode -eq 0x0C8007012f))
        {
            $host.SetShouldExit(0)
        }

        LogItem "Validate $PreviousPackageName Uninstall: Started"
        If (-NOT (Test-Path -Path $PreviousTargetFilePath))
        {
    	    LogItem "Validate $PreviousPackageName Uninstall: Able to validate Non Existence of Target File."
        }
        Else
        {
    	    LogItem "Validate $PreviousPackageName Uninstall: Target File still exits. Target File path: $PreviousTargetFilePath"
    	    QuitScript 1603
        }
        LogItem "Validate $PreviousPackageName Uninstall: Finished" 
    }
    Else
    {
        LogItem "Previous package $PreviousPackageName not found, Installing $PackageName now"
    }
}

<#
################################################################################
Additional Functions - APPV
################################################################################
#>
Function CreateKey($hive,$keyPath)
{
    $null = New-Item -Path "$($hive):\$($keyPath)" -Force -ErrorAction SilentlyContinue
}

Function SetReg($hive,$keyPath,$valueName,$value)
{
    $null = New-ItemProperty -Path "$($hive):\$($keyPath)" -Name "$($valueName)" -Value "$($value)" -Force -ErrorAction SilentlyContinue
}

Function SetDefaultReg($hive,$keyPath,$value)
{
    $null = Set-Item -Path "$($hive):\$($keyPath)" -Value "$($value)" -Type String -Force -ErrorAction SilentlyContinue
}

Function RemoveKey($hive,$keypath)
{
    $null = Remove-Item "$($hive):\$($keypath)" -Force -Recurse -ErrorAction SilentlyContinue
}

Function SetRegDWORD($hive,$keyPath,$valueName,$value)
{
    $null = New-ItemProperty -Path "$($hive):\$($keyPath)" -Name "$($valueName)" -Value ($value) -PropertyType DWord -Force -ErrorAction SilentlyContinue
}


Function AddRunVirtual($exeName)
{
    LogItem "Adding RunVirtual to registry: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppV\Client\RunVirtual\$($exeName)"
    CreateKey "HKLM" "SOFTWARE\Microsoft\AppV\Client\RunVirtual\$($exeName)"
    SetDefaultReg "HKLM" "SOFTWARE\Microsoft\AppV\Client\RunVirtual\$($exeName)" "$($PackageID)_$($VersionID)"
}

Function RemoveRunVirtual($exeName)
{
    LogItem "Removing RunVirtual registry: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppV\Client\RunVirtual\$($exeName)"
    RemoveKey "HKLM" "SOFTWARE\Microsoft\AppV\Client\RunVirtual\$($exeName)"
}

Function AddActiveSetup()
{
    LogItem "Adding ActiveSetup to registry: HKLM\SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components\$($PackageName)"
    CreateKey "HKLM" "SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components\$($PackageName)"
    SetReg "HKLM" "SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components\$($PackageName)" "StubPath" "SAMPLE STUBD PATH" #Edit Sample Data
    SetReg "HKLM" "SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components\$($PackageName)" "Version" "$($ProductVersion.Replace(".", ","))"
    SetReg "HKLM" "SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components\$($PackageName)" "ProductCode" "$($PackageID)"
}

Function RemoveActiveSetup()
{
    LogItem "Removing ActiveSetup registry: HKLM\SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components\$($PackageName)"
    RemoveKey "HKLM" "SOFTWARE\Wow6432Node\Microsoft\Active Setup\Installed Components\$($PackageName)"    
}

Function ARPCustomization_APPV
{
    LogItem "Adding ARPCustomization to registry: HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)"
    CreateKey "HKLM" "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)"
    SetReg "HKLM" "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)" "DisplayName" "$($PackageName)"
    SetReg "HKLM" "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)" "Publisher" "$($ProductVendor)"
    SetReg "HKLM" "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)" "DisplayVersion" "$($ProductVersion)"
    SetReg "HKLM" "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)" "Comments" "Script Package"
    SetReg "HKLM" "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)" "DisplayIcon" "$($ProgramData)\App-V\$($PackageID)\$($VersionID)\$($Icon)"
    SetReg "HKLM" "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)" "UninstallString" "$SysDir64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File $($ScriptPath)\$($ProductName)_Deploy.ps1 -Uninstall"
	SetRegDWORD "HKLM" "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)" "NoModify" 1

}

Function RemoveARPCustomization
{
    LogItem "Removing ARPCustomization to registry: HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)"
    RemoveKey "HKLM" "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($PackageName)"
}

Function Tatoo_Appv
{
    LogItem "Adding tatoo to registry: HKLM\Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)"
    CreateKey "HKLM" "Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)"
    SetReg "HKLM" "Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)" "Install Date" "$(Get-Date -Format g)"
    SetReg "HKLM" "Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)" "Product Vendor" "$($ProductVendor)"
    SetReg "HKLM" "Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)" "Product Name" "$($ProductName)"
    SetReg "HKLM" "Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)" "Product Version" "$($ProductVersion)"
    SetReg "HKLM" "Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)" "PackageID" "$($PackageID)"
    SetReg "HKLM" "Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)" "VersionID" "$($VersionID)"
    SetReg "HKLM" "Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)" "Install Source" "$($ScriptPath)"
    SetReg "HKLM" "Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)" "Target Directory" "$TargetDir"
}

Function RemoveTatoo_Appv
{
    LogItem "Removing Tatoo from registry: HKLM\Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)"
    RemoveKey "HKLM" "Software\Wow6432Node\$($CustomerName)\Applications\$($PackageName)"
}

Function VerifyInstall_Appv($TargetFilePath)
{
    LogItem "Validate Install: Started"
    If (Test-Path -Path $TargetFilePath)
    {
    	LogItem "Validate Install: Able to validate Target File existence. Target File path: $TargetFilePath"
    }
    Else
    {
    	LogItem "Validate Install: Unable to validate Target File existence."
    	QuitScript 1603
    }
    LogItem "Validate Install: Finished"    
}

Function VerifyUninstall_Appv($TargetFilePath)
{
    LogItem "Validate Uninstall: Started"
    If (-Not (Test-Path -Path $TargetFilePath))
    {
    	LogItem "Validate Uninstall: Able to validate Non Existence of Target File."
    }
    Else
    {
    	LogItem "Validate Uninstall: Target File still exits. Target File path: $TargetFilePath"
    	QuitScript 1603
    }
    LogItem "Validate Uninstall: Finished"   
}

<#
	MSI/EXE Functions
#>

function Init{
try{
    $Global:oShell  = New-Object -ComObject "Wscript.Shell"
    $Global:oSysInfo = New-Object -ComObject "ADSystemInfo"
    $Global:oShellApp = New-Object -ComObject "Shell.Application"
    $Global:oFSO = New-Object -ComObject "Scripting.FileSystemObject"
    $Global:oNetwork = New-Object -ComObject "WScript.Network"
    $Global:oSMSClient = New-Object -ComObject "Microsoft.SMS.Client"
	
	# Determine OS Architecture
	$Global:sProcessArchitectureX86 = [System.Environment]::ExpandEnvironmentVariables("%PROCESSOR_ARCHITECTURE%")
	If ([System.Environment]::ExpandEnvironmentVariables("%PROCESSOR_ARCHITEW6432%") -eq "%PROCESSOR_ARCHITEW6432%") {
		$Global:sProcessArchitectureW6432 = "Not Defined"
	}

	If (($Global:sProcessArchitectureX86 -eq "x86") -and ($Global:sProcessArchitectureW6432 -eq "Not Defined")){
		# Windows 32-bit
		$Global:iOSArch = 32
		$Global:iScriptArch = 32
		$Global:sSysDir32 = $Global:oShellApp.NameSpace(37).Self.Path
		$Global:sProgFiles32 = $Global:oShellApp.NameSpace(38).Self.Path
		$Global:sRegKeyRoot64 = "SOFTWARE"
		$Global:sRegKeyRoot32 = "SOFTWARE"
    }else{
		# Windows 64-bit
		$Global:iOSArch = 64
		$Global:iScriptArch = 64
		$Global:sSysDir64 = $Global:oShellApp.NameSpace(37).Self.Path
		$Global:sProgFiles64 = $Global:oShellApp.NameSpace(38).Self.Path
		$Global:sSysDir32 = $Global:oShellApp.NameSpace(41).Self.Path
		$Global:sProgFiles32 = $Global:oShellApp.NameSpace(42).Self.Path
		$Global:sRegKeyRoot32 = "SOFTWARE\Wow6432Node"
		$Global:sRegKeyRoot64 = "SOFTWARE"
	}

	# %ProgramData%\Microsoft\Windows\Start Menu
	$Global:sAllUsersStartMenu = $Global:oShellApp.NameSpace(22).Self.Path
		
	# %ProgramData%\Microsoft\Windows\Start Menu\Programs
	$Global:sAllUsersStartPrograms = $Global:oShellApp.NameSpace(23).Self.Path
	
	# %SystemDrive%\Users\Public\Desktop
	$Global:sAllUsersDesktop = $Global:oShellApp.NameSpace(25).Self.Path
	
	# %SYSTEMDRIVE%\ProgramData
	$Global:sAllUsersAppData = $Global:oShellApp.NameSpace(35).Self.Path
	
	# %WINDIR%
	$Global:sWinDir = $Global:oShellApp.NameSpace(36).Self.Path
	
	# %SYSTEMDRIVE%
	$Global:sSystemDrive = [System.Environment]::ExpandEnvironmentVariables("%SystemDrive%")
			
	# %WINDIR%\Temp - System Account
	$Global:sTempDir = [System.Environment]::ExpandEnvironmentVariables("%TEMP%")
	
	#Roaming appdata
	$Global:sAppData = [System.Environment]::ExpandEnvironmentVariables("%appdata%")
		
	# Get script directory without trailing slash
	$Global:sScriptDir = $PSScriptRoot
	
	#If root of drive, strip trailing backslash
	If (([string]$Global:sScriptDir).Length -eq 3){ $Global:sScriptDir = $Global:sScriptDir.functionstring(0, 2) }
	
	# %SYSTEMDRIVE%\Users\%USERNAME%
	$Global:sUserProfile = [System.Environment]::ExpandEnvironmentVariables("%USERPROFILE%")
	
	# Check if /uninstall was passed to the script
	$Global:bUninstall = $false

	if($Uninstall){$Global:bUninstall = $true }
		
	# used to encapsute paths
	$Global:Quotes = #"#
		
	# Convert Now() to String
	$Global:sInstallDate = (Get-Date).ToString()  
	
	# Generate GUID
	$Global:sTempGuid = CreateGuid
}catch{LogItem "Error occured at Init $($_.exception.message)" $True $False}
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
			QuitScript 3010
            #ForceReboot(1641)
			exit $iRetVal
	}
}catch{LogItem "Error occured at InstallRebootExe $($_.exception.message)" $True $False}
}

#===============================================================================
# For deleting empty directory
#===============================================================================

function DeleteFolderIfEmpty{
param($sDir)
try{
	$objFSO, $objFolder
 
	If($oFSO.FolderExists("$sDir\")){
		$objFolder = $objFSO.GetFolder("$sDir\")
    
		If (($objFolder.Files.Count -eq 0) -and ($objFolder.functionFolders.Count -eq 0)){
			DeleteFolder $sDir
		}
    }
}catch{LogItem "Error occured at DeleteFolderIfEmpty $($_.exception.message)" $True $False}
}

#========================================================================
# Logging
#========================================================================

function BeginLog{
try{
	If($Global:bUninstall){$Global:sVerb = "Uninstall"}
    Else{ $Global:sVerb = "Install" }
	
	# Create log folder path if needed
	If (-Not $oFSO.FolderExists($Global:sLogPath)){
		CreateFolderIfNeeded $Global:sLogPath
	}

	# Determine OS Properties
	$oOS = Get-WmiObject -Class Win32_OperatingSystem
    $sOSName=$sOSVersion=$sOSManufacturer=$sOSOrganization=$null
    $sOSName = $oOS.Caption
	$sOSVersion = "$($oOS.Version) $($oOS.CSDVersion) Build $($oOS.BuildNumber)"
	$sOSManufacturer = $oOS.Manufacturer
	$sOSOrganization = $oOS.Organization			
	
	# Determine ComputerSystem Properties
	$oComputer = Get-WmiObject -Class Win32_ComputerSystem
	$sSystemManufacturer,$sSystemName,$sSystemModel,$sSystemType,$sSystemRAM
	$sSystemManufacturer = $oComputer.Manufacturer
	$sSystemName = $oComputer.Name
	$sSystemModel = $oComputer.Model
	$sSystemType = $oComputer.SystemType		
	$sSystemRAM = "{0} GB" -f ([System.Math]::Round(($oComputer.TotalPhysicalMemory/1073741824),2))

	# Determine LogicalDisk Properties
	$colItems = Get-WmiObject -Class Win32_LogicalDisk
	$sVolumeName,$sTotalHardDriveSize,$sAvailableHardDriveFreeSpace
	ForEach($oItem in $colItems){
		If($oItem.Description -eq "Local Fixed Disk"){				
			$sVolumeName = $oItem.VolumeName
			$sTotalHardDriveSize = "{0} GB" -f ([System.Math]::Round($oItem.Size /1073741824))
			$sAvailableHardDriveFreeSpace = "{0} GB" -f ([System.Math]::Round($oItem.FreeSpace /1073741824))		
		}
	}

	# Determine TimeZone Properties
	$colTimeZone = Get-WmiObject -Class Win32_TimeZone
	$sSystemTimeZone
	ForEach($oTimeZone in $colTimeZone){
		$sSystemTimeZone = $oTimeZone.StandardName
	}

	# Determine Processor Properties
	$colSettings = Get-WmiObject -Class Win32_Processor
	$sProcessorDescription,$sProcessorName
	ForEach($oProcessor in $colSettings){
		$sProcessorDescription = $oProcessor.Description
		$sProcessorName = $oProcessor.Name
	}

	# Determine BIOS Properties
	$colSettings = Get-WmiObject -Class Win32_BIOS
	$sBIOSVersion
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
	LogItem ("Task Sequence Name: " + $sTSNameData ) $true  $false
	LogItem ("Task Sequence Version: " + $sTSValue)  $true  $false
	LogItem ("User Executing Script:	" + $oNetwork.UserName ) $true  $false		
	LogItem ("Product Vendor: " + $sProductVendor)  $true  $false
	LogItem ("Product Name: " + $sProductName ) $true  $false
	LogItem ("Product Version: " + $sProductVersion ) $true  $false
	LogItem ("Product Code: " + $sProductCode)  $true  $false
	LogItem ("Date: " + (Get-Date))  $true  $false
}catch{LogItem "Error occured at BeginLog $($_.exception.message)" $True $False}
}

function QuitScript{
param($iRetVal=0)
try{
	LogItem "Exiting Script $iRetVal"  $True $True
	LogItem "******************** End $sVerb ********************" $True $False
	#$sLogText = $sLogText.Trim()
	$oShell = New-Object -ComObject "WScript.Shell"
	switch($iRetVal){
	{0 -or 3010}{
		$oShell.Logevent(4, $iRetVal) #$sLogText)
        }
	default{
		$oShell.LogEvent(1, $iRetVal) #$sLogText)
        }
	}

    exit $iRetVal
}catch{LogItem "Error occured at QuitScript $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at ForceReboot $($_.exception.message)" $True $False}
}

function LogItem{
param($sMessage, $bLogFile, $bEventLog)
try{
	If ($bEventLog){
		$sLogText += "$sMessage`n"
	}

	If($bLogFile){
		#$tsLog = $null ; $tsLog = $oFSO.OpenTextFile("$Global:sLogPath\$Global:sPkgName-(Script).log", 8, $true )
		#$tsLog.WriteLine("($((get-date).ToString())) - $sMessage")
        Out-File -FilePath "$Global:sLogPath\$Global:sPkgName-(Script).log" -InputObject "$((Get-Date).ToString()) - $sMessage" -Append -Force 
	}
}catch{Write-Output "Error occured at LogItem $($_.exception.message)"}
}

#========================================================================
# Product Discovery Workers
#========================================================================

Function IsProductInstalled
{
return $false
try{
    [boolean]$ProductInstall = $false
       $path64bit = "HKLM:\SOFTWARE\Wow6432Node\FRB\Applications\$($sPkgName)" 
    $path32bit = "HKLM:\SOFTWARE\FRB\Applications\$($sPkgName)" 
       $path64bitproductcode = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$($sProductCode)" 
    $path32bitproductcode = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$($sProductCode)" 
 
    if((Test-Path $path64bit) -or (Test-Path $path32bit)){
        LogItem "AuditKey - Found"
        if((Test-Path $path64bitproductcode) -or (Test-Path $path32bitproductcode)){
            LogItem "ProductCode - Found"
            if(Test-Path $sProductFilePath){
                LogItem "ProductFile - Found"
                $versionInfo = (Get-Item $sProductFilePath).VersionInfo
                LogItem "ProductFileVersion - $($versionInfo.FileVersion)"
                if($versionInfo.FileVersion -eq $sProductFileVersion){
                    $ProductInstall = $true
                    LogItem "$sPkgName - is installed."
                    If (!$Uninstall){
                        LogItem "Quiting the script"
                        QuitScript 0
                    }
                }else{
                    LogItem "$sPkgName - is NOT Installed"
                   $ProductInstall = $false
                }
            }else{LogItem "ProductFile - Not Found"}
        }else{LogItme "ProductCode - Not Found"}
    }else{LogItem "AuditKey - Not Found"}
    return $ProductInstall
}catch{LogItem "Error occured at IsProductInstalled $($_.exception.message)" $True $False}
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
	{0 -or 3010}{
		#Success!
		LogItem "$subroutine Install was successful.  (Returned $iRetVal)" $True $True}
	1618{
		LogItem "$subroutine Install was uncessessful (Returned 1618 -- Another installation is in progress)" $True $true
		QuitScript 1618}
	default{
		#Failure
		LogItem "$subroutineInstall was unsuccessful.  (Returned $iRetVal)" $True $True
		QuitScript $iRetVal }
	}
	LogItem "$subroutine Process finished. (Returned $iRetVal)" $true $true
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occured at InstallMSI $($_.exception.message)" $True $False}
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
	{0 -or 3010}{
		#Success!
		LogItem "$subroutine Install was successful.  (Returned $iRetVal)" $true $true}
	1618{
		LogItem "$subroutine Install was uncessessful (Returned 1618 -- Another installation is in progress)" $True $true
		QuitScript 1618 }
	default{
		#Failure
		LogItem "$subroutine Install was unsuccessful.  (Returned $iRetVal)" $true $true
		QuitScript $iRetVal }
	}
	LogItem "$subroutine Process finished. (Returned $iRetVal)" $true $true
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occured at InstallMSP $($_.exception.message)" $True $False}
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
	{0 -eq 3010}{
			#Success!
			LogItem "$subroutine Uninstall was successful.  (Returned $iRetVal)" $true $true}
    1618{
			LogItem "$subroutine Uninstall was uncessessful (Returned 1618 -- Another installation is in progress)" $True $true
			QuitScript 1618 }
	default{
	
		#Failure
		LogItem "$subroutine Uninstall was unsuccessful.  (Returned $iRetVal )" $true $true
		QuitScript $iRetVal}
	}
	
	LogItem "$subroutine Process finished. (Returned $iRetVal )" $True $False
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occured at UninstallMSI $($_.exception.message)" $True $False}
}

function Install{
param($sFilePath,$sParameters)
try{
	$sFilePath
	$sParameters
	$subroutine = "Install Setup: " 
	
	LogItem "$subroutine Started" $true $False
		
	$sCmd =  "$Quotes$sFilePath$Quotes $sParameters"
   	$sCmd
   	LogItem "$subroutine Running: $sCmd"  $true $False
	
	$iRetVal = $oShell.Run($sCmd,0, $true)
	LogItem "$subroutine Process finished. (Returned $iRetVal )" $true $true
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occured at Install $($_.exception.message)" $True $False}
}

function Uninstall{
param($sFilePath,$sParameters)
try{
	$subroutine = "Uninstall Setup: " 

	LogItem "$subroutine Started" $true $False

	$sCmd = "$Quotes$sFilePath$Quotes $sParameters"
	
   	LogItem "$subroutine Running: $sCmd" $true $False
   	
	$iRetVal = $oShell.Run($sCmd,0, $true)
	
	LogItem "$subroutine Process finished. (Returned $iRetVal )" $true $False
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occured at Uninstall $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at ExecuteCMD $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at IsRegKeyExist $($_.exception.message)" $True $False ; return $false}
}

Function IsRegValNameExist{
param($sRootKey,$subKey,$sValueName)
try{
	$sKeyName = "$sRootKey\$subKey"
    $iRetVal = $oShell.Run("REG QUERY $Quotes$sKeyName$Quotes /v $Quotes$sValueName$Quotes",0, $true)
    If ($iRetVal -ne 0) {
        return $false
    }Else{
        return $true
    }
}catch{LogItem "Error occured at IsRegValNameExist $($_.exception.message)" $True $False; return $false}
}

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
}catch{LogItem "Error occured at CreateKey $($_.exception.message)" $True $False}
}

function DeleteKey{
param($sRootKey,$subKey)
try{
	$sKeyName,$iRetVal
	$sKeyName = $sRootKey + "\" + $subKey
    If ((IsRegKeyExist $sRootKey $subKey)) {
        LogItem ("About to delete: " + $Quotes + $sKeyName + $Quotes) $true $False
        $iRetVal = $oShell.Run(("REG DELETE $Quotes$sKeyName$Quotes /f"),0, $true)
        If ($iRetVal -ne 0) {
            LogItem ("$Quotes$sKeyName$Quotes was not deleted") $true $False
        }Else{
            LogItem ("$Quotes$sKeyName$Quotes has been deleted") $true $False
        }
    }Else{
        LogItem ("$Quotes$sKeyName$Quotes does not exist.") $true $False
    }
}catch{LogItem "Error occured at DeleteKey $($_.exception.message)" $True $False}
}

function SetRegVal{
param($sRootKey,$subKey,$sValueName,$sDataType,$sValue)
try{
    $sKeyName = $sRootKey + "\" + $subKey
    
    If (-Not (IsRegKeyExist($sRootKey,$subKey))) {
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
}catch{LogItem "Error occured at SetRegVal $($_.exception.message)" $True $False}
}

function DelRegValName{
param($sRootKey,$subKey,$sValueName)
try{
	$iRetVal=$sKeyName=$null
	$sKeyName = $sRootKey + "\" + $subKey
    If (IsRegValNameExist($sRootKey,$subKey,$sValueName)) {
        LogItem "About to delete $Quotes$sValueName$Quotes under $Quotes$sRootKey\$subKey$Quotes" $true $False
        $iRetVal = $oShell.Run("REG DELETE $Quotes$sKeyName$Quotes /v $sValueName /f",0, $true)
        If ($iRetVal -ne 0) {
            LogItem "$Quotes$sKeyName\$sValueName$Quotes was not deleted" $true $False
        }Else{
            LogItem "$Quotes$sKeyName\$sValueName$Quotes has been deleted" $true $False
        }
    }Else{
        LogItem "$Quotes$sKeyName\$sValueName$Quotes does not exist." $true $False
    }
}catch{LogItem "Error occured at DelRegValName $($_.exception.message)" $True $False}
}

#========================================================================
# Check Pending Reboot Routines
#========================================================================
Function PendingRebootCheck_Reg{
try{
	If (IsRegKeyExist("HKLM","SYSTEM\CurrentControlSet\services\SNAC")) {
		LogItem "Previous Package - Symantec Endpoint Protection - has uninstalled." $true $False
		LogItem "System REBOOT is requried, before installing new version." $true $False
		LogItem "Script will EXIT now without attempting to install new version." $true $False
		QuitScript 3010
	}Else{
		LogItem "Script will proceed with installing new version." $true $False
	}	 
}catch{LogItem "Error occured at PendingRebootCheck_Reg $($_.exception.message)" $True $False}
}

function PendingRebootCheck_Process{
param($sProcess)
try{
	LogItem "Process Status Check: Started" $true $False
	$oProcesses = Get-WmiObject -Class Win32_Process
	
	ForEach($oProcess in $oProcesses){
		If ($sProcess -ieq $oProcess.Name) {
			LogItem "Process Status Check: Found $sProcessis running." $true $False
			LogItem "<----> System Reboot required <---->. " $true $False
			LogItem "Process Status Check: Finished" $true $False
			QuitScript 3010
		}
	}
	LogItem "Process Status Check: Finished" $true $False
}catch{LogItem "Error occured at PendingRebootCheck_Process $($_.exception.message)" $True $False}
}

#========================================================================
# Process and Serices Manipulation Workers
#========================================================================
function TaskKill{
param($sProcess)
try{
	LogItem "Task Kill: Started" $true $False
	$oProcesses = Get-WmiObject -Class Win32_Process
	
	$bRunning = $False	

	ForEach ($oProcess In $oProcesses){
		If ($sProcess -ieq $oProcess.Name) {
			$bRunning = $True
			LogItem "Task Kill: Found $sProcess and will now be terminated" $true $False
			$oProcess.Terminate()			
		}
    }

	If ($bRunning) {
		# Wait and make sure the process is terminated.
		LogItem "Task Kill: Validating if process is terminated" $true $False
		while(-Not $bRunning){
			$oProcesses = Get-WmiObject -Class Win32_Process -Filter "Name -eq $sProcess" -ErrorAction SilentlyContinue
			#Wait for 100 MilliSeconds
			Start-Sleep 100
			#If no more processes are running, exit Loop
			If ($oProcesses.Count -eq 0) { 
				$bRunning = $False
			}
		}
	}Else{
		LogItem "Task Kill: $sProcess Is not running" $true $False
	}
	LogItem "Task Kill: Finished" $true $False
}catch{LogItem "Error occured at TaskKill $($_.exception.message)" $True $False}
}

function WaitForProcess{
param($sProcessName)
try{
    [array]$oProcess = @()
	Do{
		$oProcess = Get-WmiObject -Class Win32_Process -Filter "Name -eq $sProcessName" -ErrorAction SilentlyContinue
		Start-Sleep 1000
    }while($oProcess.Count -ne 0)
}catch{LogItem "Error occured at WaitForProcess $($_.exception.message)" $True $False}
}

function WaitAndKill{
param($sProcessName)
try{
	[array]$oProcess = @()

	Do{
		$oProcess = Get-WmiObject -Class Win32_Process -Filter "Name -eq $sProcessName" -ErrorAction SilentlyContinue
		If ($oProcess.Count -ne 0) {
			TaskKill $sProcessName
		}
		Start-Sleep 1000
	}while($oProcess.Count -ne 0)
}catch{LogItem "Error occured at WaitAndKill $($_.exception.message)" $True $False}
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
			$oProcesses = Get-WmiObject -Class Win32_Process -Filter "Name -eq $sProcess" -ErrorAction SilentlyContinue
			#Wait for 100 MilliSeconds
			Start-Sleep 100
			#If no more processes are running, exit Loop
			If ($oProcesses.Count -eq 0) { 
				$bRunning = $False
			}
		}
	}Else{
		LogItem "Wait For Process To Finish: $sProcess - Is not running" $true $False
	}
	LogItem "Wait For Process To Finish: $sProcess - Finished" $true $False
}catch{LogItem "Error occured at WaitForProcessToFinish $($_.exception.message)" $True $False}
}

function SystemLockStatus{
try{
	$isLocked = $null ; $isLocked = Get-Process -ProcessName LogonUI -ErrorAction SilentlyContinue
	if($isLocked -ne $null){$Global:sSystemLockStatus = "Locked"}
    Else{$Global:sSystemLockStatus = "Unlocked" }
}catch{LogItem "Error occured at SystemLockStatus $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at StopService $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at StartService $($_.exception.message)" $True $False}
}

function RestartService{
param($sService)
try{
	LogItem "Restart Service: Started" $true $False
	StopService $sService
	StartService $sService
	LogItem "Restart Service: Finished" $true $False
}catch{LogItem "Error occured at RestartService $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at ConfigService $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at DelService $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at FileExist $($_.exception.message)" $True $False ; return $false}
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
}catch{LogItem "Error occured at CreateFile $($_.exception.message)" $True $False}
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
                        $sSourceFile
            #$oFSO.CopyFile( "$quotes$sSourceFile$quotes", "$quotes$sDest$quotes", $bFileOverwrite)
			copy-item  "$sSourceFile" "$sDest" -force
}catch{LogItem "Error occured at CopyFile $($_.exception.message)" $True $False}
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
                        Write-Host $($oFile.Path) -back Green
                        LogItem "CopyFolder: Copying File $($oFile.Path)" $True $False
                        CopyFile $oFile.Path $sDestFolder
            }

            ForEach ($ofunctionFolder In $oFolder.functionFolders){
                        CopyFolder $ofunctionFolder.Path ("$sDestFolder{0}" -f $ofunctionFolder.Name)
            }
}catch{LogItem "Error occured at CopyFolder $($_.exception.message)" $True $False}
}

function CreateFolderIfNeeded{
param($sTarget)
try{
            $oFSO = New-Object -ComObject "Scripting.FileSystemObject"
            $sFolder = $aFolders = $sWorkingFolder = $null ; $i = 1
            $sFolder = $sTarget
            LogItem "CopyFolderIfNeeded: Folder requested to be created $sFolder" $True $False
            If ($sFolder.Substring($sFolder.Length-1,1) -eq "\"){
                        $sFolder = $sFolder.Substring(0,$sFolder.Length-1)
            }
            [array]$aFolders = $sFolder.Split("\")
            #write-host "Split Folders" -back yellow
            #$aFolders
            $sWorkingFolder = $aFolders[0]
            foreach($i in $aFolders){ #.GetUpperBound(0)){
                        if($sWorkingFolder -ne $i){
                        $sWorkingFolder = "$sWorkingFolder\$i" #$($aFolders[$i])"
                        LogItem "CopyFolderIfNeeded: Checking if folder is present $sWorkingFolder" $True $False
                        If ($oFSO.FolderExists($sWorkingFolder) -eq $false){
                                    LogItem "Create Folder if Needed: Creating Folder $Quotes $sWorkingFolder $Quotes" $true  $false
                                    $oFSO.CreateFolder($sWorkingFolder)
                        } }
            }
}catch{LogItem "Error occured at CopyFolderIfNeeded $($_.exception.message)" $True $False}
}

function DeleteFile{
param($sFile)
try{
            $oFSO = New-Object -ComObject "Scripting.FileSystemObject"
            If ($oFSO.FileExists($sFile)) {
                        $oFile = $oFSO.GetFile($sFile)
                        $iRetVal = $oFile.Delete
            }Else{
                        LogItem "$sFile does not exist." $true $False
            }
}catch{LogItem "Error occured at DeleteFile $($_.exception.message)" $True $False}
}

function MoveFile{
param($sSource,$sDest)
try{
            $oFSO = New-Object -ComObject "Scripting.FileSystemObject"
            If ($oFSO.FileExists($sSource)) {
                        CreateFolderIfNeeded $sDest
                        $iRetVal = $oFSO.MoveFile($sSource,"$sDest\")
            }Else{
                        LogItem "$sSource does not exist." $true $False
            }
}catch{LogItem "Error occured at MoveFile $($_.exception.message)" $True $False}
}

function DeleteFolder{
param($sFolder)
try{
            If ($oFSO.FolderExists($sFolder)) {
                        $oFolder = $oFSO.GetFolder($sFolder)
                        $iRetVal = $oFolder.Delete($True)
            }Else{
                        LogItem "$sFolder does not exist." $true $False
            }
}catch{LogItem "Error occured at DeleteFolder $($_.exception.message)" $True $False}
}

 function DeleteFileFromProfile{
 param($sFilePath)
 try{
	$sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	$afunctionKeys = Get-ChildItem "HKLM:\$sKeyPath"
	
	$sProfile,$sProfileName
	ForEach ($ofunctionKey In $afunctionKeys){
		$sValueName,$sValue,$sfunctionPath
		$sValueName = "ProfileImagePath"
		$sfunctionPath = "$sKeyPath\$(Split-Path -Path $ofunctionKey.Name -Leaf)"
		$sValue = (Get-ItemProperty "HKLM:\$sfunctionPath" -Name $sValueName).$sValueName
		
		$sProfile = $sValue.split("\")
        $sProfileName = $sProfile[2]
		
		# filter out unnecessary profiles 
	    If (($sProfileName -ne "config") -And ($sProfileName -ne "system32") -And ($sProfileName -ne "ServiceProfiles") -And ($sProfileName -ne "UpdatusUser") -And ($sProfileName -ne "Administrator") -And ($sProfileName -ne ("Administrator."  + $oNetwork.ComputerName)) -And ($sProfileName -ne "z_cseappinstall") -And ($sProfileName -ne "ctx_cpsvcuser") -And ($sProfileName -ne "MsDtsServer110") -And ($sProfileName -ne "ReportServer") -And ($sProfileName -ne "MSSQLFDLauncher") -And ($sProfileName -ne "SQLSERVERAGENT") -And ($sProfileName -ne "MSSQLSERVER") -And ($sProfileName -ne "QBDataServiceUser26")) {
	    	$oFolder = $oFSO.GetFolder($sValue)
			$sPath =  "$oFolder\$sFilePath"
			If ($oFSO.FileExists($sPath)) {
				DeleteFile $sPath
			}
		}
	}
}catch{LogItem "Error occured at DeleteFileFromProfile $($_.exception.message)" $True $False}
}

function DeleteFolderFromProfile{
param($sFolderPath)
try{
	$sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	$afunctionKeys = Get-ChildItem "HKLM:\$sKeyPath"
	
	$sProfile,$sProfileName
	ForEach ($ofunctionKey In $afunctionKeys){
		$sValueName,$sValue,$sfunctionPath
		$sValueName = "ProfileImagePath"
		$sfunctionPath = "$sKeyPath\$(Split-Path -Path $ofunctionKey.Name -Leaf)"
		$sValue = (Get-ItemProperty "HKLM:\$sfunctionPath" -Name $sValueName).$sValueName
		
		$sProfile = $sValue.split("\")
        $sProfileName = $sProfile[2]
		
	    # filter out unnecessary profiles 
	    If (($sProfileName -ne "config") -And ($sProfileName -ne "system32") -And ($sProfileName -ne "ServiceProfiles") -And ($sProfileName -ne "UpdatusUser") -And ($sProfileName -ne "Administrator") -And ($sProfileName -ne ("Administrator."  + $oNetwork.ComputerName)) -And ($sProfileName -ne "z_cseappinstall") -And ($sProfileName -ne "ctx_cpsvcuser") -And ($sProfileName -ne "MsDtsServer110") -And ($sProfileName -ne "ReportServer") -And ($sProfileName -ne "MSSQLFDLauncher") -And ($sProfileName -ne "SQLSERVERAGENT") -And ($sProfileName -ne "MSSQLSERVER") -And ($sProfileName -ne "QBDataServiceUser26")) {
	    	$oFolder = $oFSO.GetFolder($sValue)
			$sPath =  "$oFolder\$sFolderPath"
			If ($oFSO.FolderExists($sPath)) {
				DeleteFolder $sPath
			}
		}
	}
}catch{LogItem "Error occured at DeleteFolderFromProfile $($_.exception.message)" $True $False}
}

function CopyFolderToProfile{
param($sSourceFolder,$sTargetFolder)
try{
	$sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	$afunctionKeys = Get-ChildItem "HKLM:\$sKeyPath"
	
	$sProfile,$sProfileName
	ForEach ($ofunctionKey In $afunctionKeys){
		$sValueName,$sValue,$sfunctionPath
		$sValueName = "ProfileImagePath"
		$sfunctionPath = "$sKeyPath\$(Split-Path -Path $ofunctionKey.Name -Leaf)"
		$sValue = (Get-ItemProperty "HKLM:\$sfunctionPath" -Name $sValueName).$sValueName
		
		$sProfile = $sValue.split("\")
        $sProfileName = $sProfile[2]
		
	    # filter out unnecessary profiles 
	    If (($sProfileName -ne "config") -And ($sProfileName -ne "system32") -And ($sProfileName -ne "ServiceProfiles") -And ($sProfileName -ne "UpdatusUser") -And ($sProfileName -ne "Administrator") -And ($sProfileName -ne ("Administrator."  + $oNetwork.ComputerName)) -And ($sProfileName -ne "z_cseappinstall") -And ($sProfileName -ne "ctx_cpsvcuser") -And ($sProfileName -ne "MsDtsServer110") -And ($sProfileName -ne "ReportServer") -And ($sProfileName -ne "MSSQLFDLauncher") -And ($sProfileName -ne "SQLSERVERAGENT") -And ($sProfileName -ne "MSSQLSERVER") -And ($sProfileName -ne "QBDataServiceUser26")) {
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
}catch{LogItem "Error occured at CopyFolderToProfile $($_.exception.message)" $True $False}
}

function CopyFileToProfile{
param($sSourceFolder,$sTargetFolder,$sSourceFileName)
try{
	$sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	$afunctionKeys = Get-ChildItem "HKLM:\$sKeyPath"
	
	$sProfile,$sProfileName
	ForEach ($ofunctionKey In $afunctionKeys){
		$sValueName,$sValue,$sfunctionPath
		$sValueName = "ProfileImagePath"
		$sfunctionPath = "$sKeyPath\$(Split-Path -Path $ofunctionKey.Name -Leaf)"
		$sValue = (Get-ItemProperty "HKLM:\$sfunctionPath" -Name $sValueName).$sValueName
		
		$sProfile = $sValue.split("\")
        $sProfileName = $sProfile[2]
		
	    # filter out unnecessary profiles 
	    If (($sProfileName -ne "config") -And ($sProfileName -ne "system32") -And ($sProfileName -ne "ServiceProfiles") -And ($sProfileName -ne "UpdatusUser") -And ($sProfileName -ne "Administrator") -And ($sProfileName -ne ("Administrator."  + $oNetwork.ComputerName)) -And ($sProfileName -ne "z_cseappinstall") -And ($sProfileName -ne "ctx_cpsvcuser") -And ($sProfileName -ne "MsDtsServer110") -And ($sProfileName -ne "ReportServer") -And ($sProfileName -ne "MSSQLFDLauncher") -And ($sProfileName -ne "SQLSERVERAGENT") -And ($sProfileName -ne "MSSQLSERVER") -And ($sProfileName -ne "QBDataServiceUser26")) {
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
}catch{LogItem "Error occured at CopyFileToProfile $($_.exception.message)" $True $False}
}

#========================================================================
# Install Validation and Tattoo Workers
#========================================================================

function ValidateInstall{
try{
	$sInstalledFileVer = $oFSO.GetFileVersion($sProductFilePath)
	
	If ($sInstalledFileVer -ne $sProductFileVersion){ 
		LogItem "Unable to validate install.$sInstalledFileVer" $true $False 
		QuitScript 1603
	}Else{
		LogItem "Validated installation." $true $False
	}
}catch{LogItem "Error occured at ValidateInstall $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at Validate_DateModified_BaseApp $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at ValidateInstall_DateModified_Patch $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at ValidateInstall_File $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at ValidateUninstall_File $($_.exception.message)" $True $False}
}

Function GetRegPath{
try{
	LogItem "GetRegPath: checking for HKLM\32or64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode" $true $false
	If ((IsRegKeyExist "HKLM" ("$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode"))) {
		LogItem "GetRegPath: found HKLM\$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode - setting sRegPath" $true $false
		$Global:sRegPath = "$sRegKeyRoot64\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode"
		$Global:sTattooPath = "$sRegKeyRoot64\$sRegAuditPath"
	}ElseIf ((IsRegKeyExist "HKLM" ("$sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode"))) {	
	LogItem "GetRegPath: found $sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode - setting sRegPath" $true $false
		$Global:sRegPath = "$sRegKeyRoot32\Microsoft\Windows\CurrentVersion\Uninstall\$sProductCode"
		$Global:sTattooPath = "$sRegKeyRoot32\$sRegAuditPath"
	}else{LogItem "GetRegPath: uninstall\$sProductCode -keynot found - sregpath not set" $true $false}
	
}catch{LogItem "Error occured at GetRegPath $($_.exception.message)" $True $False}
}

function TattooRegistry{
try{
	LogItem "TattooRegistry: process Started" $true $false
    SetRegVal "HKLM" "$Global:sTattooPath\$sPkgName" "Install Date" "REG_SZ" $sInstallDate
    SetRegVal "HKLM" "$Global:sTattooPath\$sPkgName" "Vendor Name" "REG_SZ" $sProductVendor
    SetRegVal "HKLM" "$Global:sTattooPath\$sPkgName" "Product Version" "REG_SZ" $sProductVersion
    SetRegVal "HKLM" "$Global:sTattooPath\$sPkgName" "Product Code" "REG_SZ" $sProductCode
    SetRegVal "HKLM" "$Global:sTattooPath\$sPkgName" "Target Directory" "REG_SZ" $sInstallFolderPath
    SetRegVal "HKLM" "$Global:sTattooPath\$sPkgName" "Install Source" "REG_SZ" $sScriptDir
}catch{LogItem "Error occured at TatooRegistry $($_.exception.message)" $True $False}
}

function RemoveTattoo{
try{
	DeleteKey "HKLM" ("$sTattooPath\$sPkgName")
}catch{LogItem "Error occured at RemoveTattoo $($_.exception.message)" $True $False}
}

function ARPCustomization{
try{
	LogItem "$Global:sRegPath" $true $false
	if($Global:sRegPath -eq ""){exit}
    SetRegVal "HKLM" $Global:sRegPath "DisplayName" "REG_SZ" $sPkgName
    SetRegVal "HKLM" $Global:sRegPath "NoModify" "REG_DWORD" "1"
    SetRegVal "HKLM" $Global:sRegPath "NoRepair" "REG_DWORD" "0"
    SetRegVal "HKLM" $Global:sRegPath "Comments" "REG_SZ" "Script Package"
    SetRegVal "HKLM" $Global:sRegPath "Contact" "REG_SZ" ""
    SetRegVal "HKLM" $Global:sRegPath "HelpLink" "REG_SZ" ""
    SetRegVal "HKLM" $Global:sRegPath "Readme" "REG_EXPAND_SZ" ""
    SetRegVal "HKLM" $Global:sRegPath "URLUpdateInfo" "REG_SZ" ""
    SetRegVal "HKLM" $Global:sRegPath "URLInfoAbout" "REG_SZ" ""
}catch{LogItem "Error occured at ARPCustomization $($_.exception.message)" $True $False}
}

Function TattooInstallFolder{
param($sInstallDir)
try{
	$oFSO.CreateTextFile($sInstallDir)
}catch{LogItem "Error occured at TattooInstallFolder $($_.exception.message)" $True $False}
}

#=======================================================================
# Per-User Registry Workers
#========================================================================

function AddToUserHives{
param($sUserRegPath,$sValueName,$sType,$sValue)
try{
    $sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	$afunctionKeys = Get-ChildItem "HKLM:\$sKeyPath"
	
	$sProfile,$sProfileName
	ForEach ($ofunctionKey In $afunctionKeys){
		$sValueName,$sValue,$sfunctionPath
		$sValueName = "ProfileImagePath"
		$sfunctionPath = "$sKeyPath\$(Split-Path -Path $ofunctionKey.Name -Leaf)"
		$sValue = (Get-ItemProperty "HKLM:\$sfunctionPath" -Name $sValueName).$sValueName
		
		$sProfile = $sValue.split("\")
        $sProfileName = $sProfile[2]
		
	    # filter out unnecessary profiles 
	    If (($sProfileName -ne "config") -And ($sProfileName -ne "system32") -And ($sProfileName -ne "ServiceProfiles") -And ($sProfileName -ne "UpdatusUser") -And ($sProfileName -ne "Administrator") -And ($sProfileName -ne ("Administrator."  + $oNetwork.ComputerName)) -And ($sProfileName -ne "z_cseappinstall") -And ($sProfileName -ne "ctx_cpsvcuser") -And ($sProfileName -ne "MsDtsServer110") -And ($sProfileName -ne "ReportServer") -And ($sProfileName -ne "MSSQLFDLauncher") -And ($sProfileName -ne "SQLSERVERAGENT") -And ($sProfileName -ne "MSSQLSERVER") -And ($sProfileName -ne "QBDataServiceUser26")) {
	    	If (Test-Path "$sProfileDir\NTuser.dat") {
				LogItem "Will try to mount user hive: $sProfileDir\NTuser.dat" $true $False
				If ((MountHive "$sProfileDir\NTuser.dat")) {
					SetRegVal "HKU" "CUSTOM\$sUserRegPath" $sValueName $sType $sValue
					UnmountHive
				}Else{
					$sDomain = $env:USERDNSDOMAIN
					
					$sSID = $oProfile
					SetRegVal "HKU" "$sSID\$sUserRegPath" $sValueName $sType $sValue
				}
			}
		}
	}

    MountDefaultHive
    LogItem "Will now update default user hive for all users." $true $False
	SetRegVal "HKU" "CUSTOM\$sUserRegPath" $sValueName $sType $sValue
    LogItem "Default user hive is now updated for all users." $true $False
	UnmountHive
}catch{LogItem "Error occured at AddToUserHives $($_.exception.message)" $True $False}
}

function DeleteFromUserHives{
param($sUserRegPath)
try{
    $sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	$afunctionKeys = Get-ChildItem "HKLM:\$sKeyPath"
	
	$sProfile,$sProfileName
	ForEach ($ofunctionKey In $afunctionKeys){
		$sValueName,$sValue,$sfunctionPath
		$sValueName = "ProfileImagePath"
		$sfunctionPath = "$sKeyPath\$(Split-Path -Path $ofunctionKey.Name -Leaf)"
		$sValue = (Get-ItemProperty "HKLM:\$sfunctionPath" -Name $sValueName).$sValueName
		
		$sProfile = $sValue.split("\")
        $sProfileName = $sProfile[2]
		
	    # filter out unnecessary profiles 
	    If (($sProfileName -ne "config") -And ($sProfileName -ne "system32") -And ($sProfileName -ne "ServiceProfiles") -And ($sProfileName -ne "UpdatusUser") -And ($sProfileName -ne "Administrator") -And ($sProfileName -ne ("Administrator."  + $oNetwork.ComputerName)) -And ($sProfileName -ne "z_cseappinstall") -And ($sProfileName -ne "ctx_cpsvcuser") -And ($sProfileName -ne "MsDtsServer110") -And ($sProfileName -ne "ReportServer") -And ($sProfileName -ne "MSSQLFDLauncher") -And ($sProfileName -ne "SQLSERVERAGENT") -And ($sProfileName -ne "MSSQLSERVER") -And ($sProfileName -ne "QBDataServiceUser26")) {
	         If (Test-Path "$sProfileDir\NTuser.dat") {
                LogItem "Will try to mount user hive: $sProfileDir\NTuser.dat" $true $False
                If ((MountHive "$sProfileDir\NTuser.dat")) {
                    DeleteKey "HKU" "CUSTOM\$sUserRegPath"
                    UnmountHive
                }Else{
                    $sDomain = $env:USERDNSDOMAIN
					
					$sSID = $oProfile

                    DeleteKey "HKU" "$sSID\$sUserRegPath"
                }
            }
        }
    }

    MountDefaultHive
    LogItem "Will now update default user hive for all users." $true $False
	DeleteKey "HKU" "CUSTOM\$sUserRegPath"
    LogItem "Default user hive is now updated for all users." $true $False
	UnmountHive
}catch{LogItem "Error occured at DeleteFromUserHives $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at MountHive $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at UnmountHive $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at MountDefaultHive $($_.exception.message)" $True $False}
}

function CreateGuid
{
try{
	return [guid]::NewGuid()
}catch{LogItem "Error occured at CreateGuid $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at CreateShortcuts $($_.exception.message)" $True $False}	
}

function DelUserProfileDir {
param($sUserProfileDirPath)
try{
	$oFSO=$sStartFolder=$sfunctionFolder=$sTargetFolder=$oShell=$null
	
	$oShell = New-Object -ComObject "WScript.Shell"
	$oFSO = New-Object -ComObject "Scripting.FileSystemObject"
	$sSystemDrive = [System.Environment]::ExpandEnvironmentVariables("%SystemDrive%")	
	$sStartFolder = $sSystemDrive + "\" + "Users"

	ForEach ($sfunctionFolder In $oFSO.GetFolder($sStartFolder).functionFolders){
		$sTargetFolder =  "$sfunctionFolder\$sUserProfileDirPath"
		If ($oFSO.FolderExists($sTargetFolder)) {
			LogItem "About to delete: $sTargetFolder" $true $False
			If ($oFSO.GetFolder($sTargetFolder).Files.Count -ne 0) {
				$oFSO.DeleteFile("$sTargetFolder\*")
			}
   			$oFSO.DeleteFolder($sTargetFolder)
   		}
   }	
}catch{LogItem "Error occured at DelUserProfileDir $($_.exception.message)" $True $False}	
}

function AppendSysPATH{
param($sPathToAdd)
try{
	$oShell = New-Object -ComObject "WScript.Shell"
	$sSysEnv = $oShell.Environment("SYSTEM")	
	$sOldSysPath = $sSysEnv.Item("Path")
	LogItem "Current System PATH Variable value is: $sOldSysPath" $true $False
	LogItem "About to Append this value from System PATH Variable: $sPathToAdd" $true $False										
	$sNewSysPath = "$sOldSysPath;$sPathToAdd"
	$sSysEnv.Item("Path") = $sNewSysPath
	LogItem "After Appending System PATH Variable value is: $($sSysEnv.Item('Path'))" $true $False	
}catch{LogItem "Error occured at AppenSysPath $($_.exception.message)" $True $False}	
}

function DelSysPATH{
param($sPathToRemove)
try{
	$oShell = New-Object -ComObject "WScript.Shell"
	$sSysEnv = $oShell.Environment("SYSTEM")	
	$sOldSysPath = $sSysEnv.Item("Path")
	LogItem "Current System PATH Variable value is: $sOldSysPath" $true $False
	LogItem "About to Remove this value from System PATH Variable: $sPathToRemove" $true $False
	$sNewSysPath = ($sOldSysPath -Replace $sPathToRemove, "")
	$sSysEnv.Item("Path") = $sNewSysPath
	LogItem "After Removing System PATH Variable value is: $($sSysEnv.Item('Path'))" $true $False	
}catch{LogItem "Error occured at DelSysPath $($_.exception.message)" $True $False}	
}

function UnZip{
param($sZipFilePath,$sExtractDirPath)
try{
	$subroutine=$oShellApp=$oFSO=$sFilesInZip=$null
	$subroutine = "UnZip Process: "
	LogItem "$subroutine Started" $true $False
	#ZipFile Path
	LogItem "$subroutine Current location of Zip File: $sZipFilePath"  $true $False
	
	#The folder the contents should be extracted to.
	LogItem "$subroutine Will get Extracted to this directory: $sExtractDirPath"  $true $False
	
	#If the extraction location does not exist create it.
	Set oFSO = CreateObject("Scripting.FileSystemObject")	
	If (-Not $oFSO.FolderExists($sExtractDirPath)) {
   		$oFSO.CreateFolder($sExtractDirPath)
	}

	#Extract the contants of the zip file.
	$oShellApp = New-Object -ComObject "Shell.Application"
	$sFilesInZip= $oShellApp.NameSpace($sZipFilePath).items
	$oShellApp.NameSpace($sExtractDirPath).CopyHere($sFilesInZip)
	$oFSO = $null
	$oShellApp = $null
	LogItem "$subroutine Finished" $true $False
}catch{LogItem "Error occured at Unzip $($_.exception.message)" $True $False}
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
}catch{LogItem "Error occured at Uninstall_Pre_Chrome $($_.exception.message)" $True $False}
}   

 Function Cleanup_Chrome_AuditRegKeys{
try{
	[array]$arrAuditRegKeys = @("Google_Chrome_51.0","Google_Chrome_52.0.2743","Google_Chrome_55.0.2883.75","Google_Chrome_57.0.2987.98","Google_Chrome_58.0.3029.96")
	
	ForEach($arrAuditRegKey In $arrAuditRegKeys){
		If ((IsRegKeyExist "HKLM" "SOFTWARE\Wow6432Node\FRB\Applications\$arrAuditRegKey")){
			LogItem "Found: HKLM\SOFTWARE\Wow6432Node\FRB\Applications\$arrAuditRegKey" $True $False
			DeleteKey "HKLM" "SOFTWARE\Wow6432Node\FRB\Applications\$arrAuditRegKey"	
		}
	}	
}catch{LogItem "Error occured at Cleanup_Chrome_AuditRegKeys $($_.exception.message)" $True $False}
}  
 
#Main script starts here
MainScript