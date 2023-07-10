#
# Press 'F5' to run this script. Running this script will load the ConfigurationManager
# module for Windows PowerShell and will connect to the site.
#
# This script was auto-generated at '10/4/2021 6:34:31 PM'.
# Uncomment the line below if running in an environment where script signing is
# required.
#Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
# Site configuration
$SiteCode = "" # Site code
$ProviderMachineName = "" # SMS Provider machine name
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
$Departments = Get-ADuser -filter "Department -like '*'" -Properties Department | select -ExpandProperty Department -Unique | Sort
$Schedule1 = New-CMSchedule -DurationInterval Days -DurationCount 0 -RecurInterval Days -RecurCount 1
Foreach ($Department in $Departments){
    if((Get-CMCollection -Name $Department) -eq $null)
    {
        Write-Host "$Department does not exist. Creating user collection..."
        $NewCollection = New-CMCollection -LimitingCollectionId SMS00002 -Name "$Department" -RefreshSchedule $Schedule1 -RefreshType Periodic -CollectionType User
        Move-CMObject -FolderPath '.\UserCollection\Department Collections' -InputObject $NewCollection
        Add-CMUserCollectionQueryMembershipRule -CollectionName $Department -RuleName "Department_$Department" -QueryExpression "select SMS_R_USER.ResourceID,SMS_R_USER.ResourceType,https://urldefense.com/v3/__http://SMS_R_USER.Name__;!!HcVu0Bn_hI1r_w!k2lEnfiNQYX2GWmRuwifgzJGgSf6w7EZwnrSXfBovUBD3h7uPL5AO3j9FoTHCaK8AvSO$ ,SMS_R_USER.UniqueUserName,SMS_R_USER.WindowsNTDomain from SMS_R_User where SMS_R_User.department like `"$Department`""
        Write-Host "$Department user collection created successfully"
    } else
    {
         Write-Host "$Department user collection already exists"
    }
    If((Get-CMDeviceCollection  -Name $Department) -eq $null)
    {
        Write-Host "$Department does not exist. Creating device collection..."
        New-CMDeviceCollection -Name "$Department" -LimitingCollectionName "All Workstations"
        Write-Host "$Department device collection created successfully"
    } else
    {
         Write-Host "$Department device collection already exists"
    }
    Write-Host "Fetching users from $Department user collection"
    $Users = Get-CMCollection -Name "$Department" | Get-CMCollectionMember | Select-Object Name
    Write-Host $Users
    Write-Host "Fetching primary devices for $user from $Department user collection"
    Foreach ($user in $Users)
    {
       $DeviceNames = (Get-CMUserDeviceAffinity -UserName "$user").ResourceID
       Write-host "Primary Devices for $user`: $DeviceNames"
       $PrimaryDevices = (Get-CMUserDeviceAffinity -UserName "$user").ResourceID
       Foreach ($PrimaryDevice in $PrimaryDevices)
       {
            Add-CMDeviceCollectionDirectMembershipRule -CollectionName "$Department" -ResourceID $PrimaryDevice
       }
    }
} 