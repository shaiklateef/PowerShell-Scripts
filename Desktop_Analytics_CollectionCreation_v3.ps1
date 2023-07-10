#get dept names from ADUC
Import-Module activedirectory
#$Departments = get-aduser -filter * -property department |select -ExpandProperty department -Unique
$Departments = "Rltnshp Mgmt Bellevue"
$DepartmentsCount = $Departments.count
$Departments
$DepartmentsCount

#############################################################################

#Load Configuration Manager PowerShell Module
Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5)+ '\ConfigurationManager.psd1')

#Get SiteCode
$SiteCode = Get-PSDrive -PSProvider FRW
Set-location $SiteCode":"

#Error Handling and output
#Clear-Host
#$ErrorActionPreference= 'SilentlyContinue'



#Create User Collection Default Folder 
$CollectionFolderUser = @{Name ="Desktop_Analytics_User"; ObjectType =5001; ParentContainerNodeId =0}
Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $CollectionFolderUser -ComputerName $SiteCode.Root
$FolderPathUser =($SiteCode.Name +":\UserCollection\" + $CollectionFolderUser.Name)

#Create Device Collection Default Folder
$CollectionFolderDevice = @{Name ="Desktop_Analytics_Device"; ObjectType =5000; ParentContainerNodeId =0}
Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $CollectionFolderDevice -ComputerName $SiteCode.Root
$FolderPathDevice =($SiteCode.Name +":\DeviceCollection\" + $CollectionFolderDevice.Name)

#Set Default limiting collections
$LimitingCollectionUser ="All Users"
$LimitingCollectionDevice ="All Windows Workstations Systems"

#Refresh Schedule
$Schedule =New-CMSchedule –RecurInterval Days –RecurCount 7


#Find Existing User Collections
$ExistingCollectionsUser = Get-CMUserCollection -Name "* | *" | Select-Object CollectionID, Name
$ExistingCollectionsUserCount = $ExistingCollectionsUser.Count

$ExistingCollectionsDevice = Get-CMDeviceCollection -Name "* | *" | Select-Object CollectionID, Name
$ExistingCollectionsDeviceCount = $ExistingCollectionsDevice.Count

$UserCollections = @()

Foreach ($Department in $Departments)
{
$UserCollectionName = @()
$UserCollectionName = $Department
$UserCollections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Desktop_Analytics_User"}},@{L="Query"
; E={select *  from  SMS_R_User where SMS_R_User.department = "$Department"}},@{L="LimitingCollectionUser"
; E={$LimitingCollectionUser}},@{L="Comment"
; E={"All Users in $Department detected by SCCM"}}
}

#Check Existing Collections
$Overwrite = 1
$ErrorCount = 0
$ErrorHeader = "The script has already been run. The following collections already exist in your environment:`n`r"
$ErrorCollections = @()
$ErrorFooter = "Would you like to delete and recreate the collections above? (Default : No) "
$ExistingCollections | Sort-Object Name | ForEach-Object {If($Collections.Name -Contains $_.Name) {$ErrorCount +=1 ; $ErrorCollections += $_.Name}}

Foreach ($collectionUser in $collectionUsers)
{
$UserCollectionData = @()
$CollectionID = $collectionUser.CollectionID
$UserCollectionData = "SMS_CM_RES_COLL_" + $collectionID

$DeviceCollections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Desktop_Analytics_Device"}},@{L="Query"
; E={select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System    LEFT JOIN SMS_UserMachineRelationship ON       SMS_UserMachineRelationship.ResourceID = SMS_R_System.ResourceId        WHERE SMS_UserMachineRelationship.IsActive= 1 AND SMS_UserMachineRelationship.UniqueUserName IN    (SELECT SMS_R_User.UniqueUserName      FROM SMS_R_User       INNER JOIN $UserCollectionData ON         $UserCollectionData.ResourceID = SMS_R_User.ResourceID)
}},@{L="LimitingCollectionDevice"
; E={$LimitingCollectionDevice}},@{L="Comment"
; E={"All Devices for $Department detected by SCCM"}}

}
