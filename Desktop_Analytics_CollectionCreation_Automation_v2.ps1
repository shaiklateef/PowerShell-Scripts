#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows PowerShell and will connect to the site.
#
# This script was auto-generated at '10/4/2021 6:34:31 PM'.
 
# Uncomment the line below if running in an environment where script signing is 
# required.
#Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
 
# Site configuration
#$SiteCode = "FRW" # Site code 
#$ProviderMachineName = "DC1SCCMPS01V.corp.firstrepublic.com" # SMS Provider machine name
 
# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors
 
# Do not change anything below this line
 
# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}
 
# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}
 
# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams
$Departments = Get-ADuser -filter "Department -like 'Rltnshp Mgmt Bellevue'" -Properties Department | select -ExpandProperty Department -Unique | Sort
$Schedule1 = New-CMSchedule -DurationInterval Days -DurationCount 0 -RecurInterval Days -RecurCount 1
Foreach ($Department in $Departments){
    $DepartmentTrimmed = $Department.Replace(' ','_')
    Write-host "Old: $Department New: $DepartmentTrimmed" -ForegroundColor Cyan
    if((Get-CMCollection -Name "$DepartmentTrimmed" -CollectionType User) -eq $null)
    {
        Write-Host "$DepartmentTrimmed does not exist. Creating user collection..."
        $NewCollection = New-CMCollection -LimitingCollectionId SMS00002 -Name "$DepartmentTrimmed" -RefreshSchedule $Schedule1 -RefreshType Periodic -CollectionType User
        Write-Host $NewCollection 
        #Create User Collection Default Folder 
        $CollectionFolderUser = @{Name ="Desktop_Analytics_User"; ObjectType =5001; ParentContainerNodeId =0}
        Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $CollectionFolderUser -ComputerName $SiteCode.Root
        $FolderPathUser =($SiteCode.Name +":\UserCollection\" + $CollectionFolderUser.Name)
        Move-CMObject -FolderPath '.\UserCollection\Desktop_Analytics_User' -InputObject $NewCollection
        Add-CMUserCollectionQueryMembershipRule -CollectionName $DepartmentTrimmed -RuleName "Department_$DepartmentTrimmed" -QueryExpression "select SMS_R_USER.ResourceID,SMS_R_USER.ResourceType,SMS_R_USER.Name,SMS_R_USER.UniqueUserName,SMS_R_USER.WindowsNTDomain from SMS_R_User where SMS_R_User.department like `"$Department`""
        Write-Host "$DepartmentTrimmed user collection created successfully"
    } else 
    {
         Write-Host "$DepartmentTrimmed user collection already exists"
    }
 
    If((Get-CMCollection -Name "$DepartmentTrimmed" -CollectionType Device) -eq $null)
    {
        Write-Host "$DepartmentTrimmed does not exist. Creating device collection..."
        $DeviceCollection = "$DepartmentTrimmed" + "_" + "Device"
        $CollectionCreated = New-CMCollection -Name $DeviceCollection -LimitingCollectionName "All Windows Workstation Systems" -CollectionType Device
        #Create Device Collection Default Folder
        $CollectionFolderDevice = @{Name ="Desktop_Analytics_Device"; ObjectType =5000; ParentContainerNodeId =0}
        Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $CollectionFolderDevice -ComputerName $SiteCode.Root
        $FolderPathDevice =($SiteCode.Name +":\DeviceCollection\" + $CollectionFolderDevice.Name)
        Move-CMObject -FolderPath '.\DeviceCollection\Desktop_Analytics_Device' -InputObject  $CollectionCreated
        Write-Host "$DepartmentTrimmed device collection created successfully"
    } else 
    {
         Write-Host "$DepartmentTrimmed device collection already exists"
    }
    Write-Host "Fetching users from $DepartmentTrimmed user collection"
    $Users = Get-CMCollection -Name "$DepartmentTrimmed" | Get-CMCollectionMember | Select-Object SMSID
    Write-Host $Users
    Write-Host "Fetching primary devices for $user from $DepartmentTrimmed user collection"
    Foreach ($user in $Users)
    {
       $username = $user.SMSID
       $DeviceNames = (Get-CMUserDeviceAffinity -UserName "$username").ResourceName
       Write-host "Primary Devices for $user`: $DeviceNames. Please wait while these devices are added"
       $PrimaryDevices = (Get-CMUserDeviceAffinity -UserName "$username").ResourceID
       Foreach ($PrimaryDevice in $PrimaryDevices)
       {
            Add-CMDeviceCollectionDirectMembershipRule -CollectionName $DeviceCollection -ResourceID $PrimaryDevice
       }
    }
 
}