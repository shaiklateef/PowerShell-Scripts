<# 
.SYNOPSIS 
Move-ComputersToOU.ps1 The script to move computer objects to AD OU according to their actual location.
 
.DESCRIPTION
The script moves the computer object to the OU. OU is identified by using below logic.

1.	Set a default OU ("OU=Desktops,OU=Computers,OU=San Francisco Corporate,OU=FRB Sites,DC=corp,DC=firstrepublic,DC=com")
a.	This is where the script will move if it cannot get valid OU details as mentioned in email in-line
2.	Extra the FRB Sites OU info and identify it has the “Desktop\Computers” sub ou.
a.	If the sub ou is not found or doesn’t have any computer, then it sets to default OU name for that OU.
3.	Identify the computers and get the top name/site code (first 4 letters) matches add it to the dictionary
4.	Get the computer names from the list
5.	Go through each and every computer and get the OU details
6.	Move to the respective AD OU.



AuthorName      Mphasis
Develeoper      Thanuj M
CreatedDate     12:40 PM 3/23/2018
ModifiedDate    5:20 PM 3/23/2017
VersionNumber   1.0.1.0
VersionHistory	
    1.0.0.0
                Skeleton (basic) script created
                12:40 PM 3/23/2018
    1.0.1.0
                parameter included to allow admins to choose the default OU to move.
                Error/Warning functionality enabled
                5:20 PM 3/23/2018

    1.1.0.0
                Identifying laptop and desktop using chasis type and then move to appropriate OU
                Checking SCCM to identify the chasis type instead of remoting to computer
                3:19 PM 4/5/2018
                   
ScriptUsage     .\Move-ComputersToOU.ps1

.PARAMETER ComputerList
This parameter is mandatory. Specifiy the full path of computer list file. e.g. "C:\Temp\Computers.txt"

.PARAMETER DefaultOUtoMove
This parameter is not mandatory. Define the FULL DistinguishedName of the OU. e.g. "OU=Desktops,OU=Computers,OU=San Francisco Corporate,OU=FRB Sites,DC=corp,DC=firstrepublic,DC=com"

.EXAMPLE  
".\Move-ComputersToOU.ps1" -ComputerList C:\temp\computers.txt

This will move all the computers specified in the ComputerList parameter

.EXAMPLE  
".\Move-ComputersToOU.ps1" -ComputerList C:\temp\computers.txt -DefaultOUtoMove "OU=Desktops,OU=Computers,OU=San Francisco Corporate,OU=FRB Sites,DC=corp,DC=firstrepublic,DC=com"

This will move all the computers specified in the ComputerList parameter and sets the default OU to user defined values instead of default one.
#>

[CmdletBinding(SupportsShouldProcess=$True)]
param(
    [Parameter( position=0,
    Mandatory=$true,
    HelpMessage='Specify the full path of computer list file. e.g. "C:\Temp\Computers.txt"')]
    [string]$ComputerList,
    [Parameter( position=1,
    Mandatory=$False,
    HelpMessage='Specify the full DistinguishedName of the default OU. e.g. "OU=Computers,OU=San Francisco Corporate,OU=FRB Sites,DC=corp,DC=firstrepublic,DC=com"')]
    [string]$DefaultOUtoMove="OU=Computers,OU=San Francisco Corporate,OU=FRB Sites,DC=corp,DC=firstrepublic,DC=com",
    [Parameter( position=2,
    Mandatory=$false,
    HelpMessage='Specify the OU Name of Desktops. e.g. "OU=Desktops"')]
    [string]$DesktopOU="OU=Desktops",
    [Parameter( position=3,
    Mandatory=$false,
    HelpMessage='Specify the OU Name of Laptops. e.g. "OU=Laptops"')]
    [string]$LaptopOU="OU=Laptops",
    [Parameter( position=4,
    Mandatory=$false,
    HelpMessage='Name of your primary site server for e.g. DC1SCCMPS02v')]
    [string]$SiteServer="DC1SCCMPS01v",
    [Parameter( position=5,
    Mandatory=$false,
    HelpMessage='Primary site code for e.g. FR1')]
    [string]$SiteCode="FRW"
 )

function _LogMessage ([parameter(mandatory=$false)]$LogFullPath, ${tMessage}, ${Severity})
{
    #Author: Jack Cheung
       switch (${Severity})
       {
              1 {
                     $LogPrefix = "INFO"
                     $fgcolor = [ConsoleColor]::Blue
                     $bgcolor = [ConsoleColor]::White
              }
              2 {
                     $LogPrefix = "WARNING"
                     $fgcolor = [ConsoleColor]::Black
                     $bgcolor = [ConsoleColor]::Yellow
              }
              3 {
                     $LogPrefix = "ERROR"
                     $fgcolor = [ConsoleColor]::Yellow
                     $bgcolor = [ConsoleColor]::Red
              }
              default
              {
                     $LogPrefix = "DEFAULT"
                     $fgcolor = [ConsoleColor]::Black
                     $bgcolor = [ConsoleColor]::White
              }
       }
       
       if(-not [string]::IsNullOrEmpty($LogFullPath)){ Add-Content -Path $LogFullPath -Value "$((Get-Date).ToString()) ${LogPrefix}: ${tMessage}" -ErrorAction SilentlyContinue }
       Write-Host -ForegroundColor $fgcolor -BackgroundColor $bgcolor -Object "$((Get-Date).ToString()) ${LogPrefix}: ${tMessage}"
}

function isDesktop([parameter(mandatory=$true)]$ResourceName){
    [bool]$IsDesktop = $true

    try{
        #"$ResourceName : Reading SMS_G_System_SYSTEM_ENCLOSURE..."
        $SMS_G_System_SYSTEM_ENCLOSURE = Get-WmiObject -ComputerName $SiteServer -Namespace root\sms\site_$SiteCode `
                                            -Query "SELECT * FROM SMS_R_SYSTEM LEFT OUTER JOIN SMS_G_System_SYSTEM_ENCLOSURE ON SMS_G_System_SYSTEM_ENCLOSURE.ResourceId = SMS_R_SYSTEM.ResourceID WHERE Name = '$ResourceName'"

        if(-not [string]::IsNullOrEmpty($SMS_G_System_SYSTEM_ENCLOSURE.SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes)){
            switch($SMS_G_System_SYSTEM_ENCLOSURE.SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes){
                1{} #"Virtual"
                {($_ -eq 3) -or ($_ -eq 4) -or ($_ -eq 6) -or ($_ -eq 7) -or ($_ -eq 15)}{} #"Desktop"
                {($_ -eq 8) -or ($_ -eq 9) -or ($_ -eq 10) -or ($_ -eq 21)}{ $IsDesktop = $false } #"Laptop"
                default{} #$SMS_G_System_SYSTEM_ENCLOSURE.SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes
            }
        }
        else{
            #Write-Warning "No value returned from SCCM"
        }
    }
    catch{
        #Write-Warning $("{0} - Unable to reach SMS $SiteServer\sms\site_$SiteCode {1}" - $ResourceName, $_.exception.message)
        $IsDesktop = $true
    } 
    return $IsDesktop
}
 
try{
    [string]$LogPath = "$PSScriptRoot\MoveOU-Output.log"
    [array]$IsADThere = Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue

    if($IsADThere.Count -le 0){
        Import-Module ActiveDirectory -ErrorAction Stop
    }

    [array]$CheckADCommands = Get-Command Get-ADComputer, Get-ADOrganizationalUnit, Move-ADObject, Add-ADGroupMember
    
    if($CheckADCommands.Count -lt 4){
        _LogMessage -LogFullPath $LogPath -tMessage "ActiveDirectory Module did not load properly, please make sure the module is available to proceed further" -Severity 2
    }else{

    [hashtable]$SiteVsOUDN = @{}
    [hashtable]$SiteVsCC = @{}
    [string]$SearchBase = (Get-ADOrganizationalUnit -Filter "Name -eq 'FRB Sites'").DistinguishedName
    [string]$ComputersOU = "OU=Computers"

    [array]$SubSites = Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase -SearchScope OneLevel -ErrorAction SilentlyContinue

    if($SubSites.Count -le 0){
        _LogMessage -LogFullPath $LogPath -tMessage "No OU data retrived!" -Severity 2
    }
    else{
        foreach($SubSite in $SubSites){
            try{
                $CC = $null
                $curSearchBase = $SubSite.DistinguishedName
                $SiteVsOUDN.Add($SubSite.Name,"$ComputersOU,$curSearchBase")
                Get-ADOrganizationalUnit -Filter * -SearchBase "$ComputersOU,$curSearchBase" -ErrorAction Stop | Out-Null

                [array]$namingConvention = @()

                @($DesktopOU,$LaptopOU) | ForEach-Object {
                    $tmpOU = $_
                    try{
                        $namingConvention += Get-ADComputer -Filter * -SearchBase "$tmpOU,$ComputersOU,$curSearchBase" -ErrorAction SilentlyContinue | Select-Object @{Label="SiteName";Expression={$SubSite.Name}}, Name, @{Label="CC";Expression={ ($_.Name).SubString(0,4) }}
                    }
                    catch{
                        _LogMessage -LogFullPath $LogPath -tMessage $("{0} - $tmpOU not found" -f $SubSite.Name) -Severity 2
                    }
                }

                if($namingConvention.Count -eq 0){
                    _LogMessage -LogFullPath $LogPath -tMessage $("{0} - Default OU will be used for this - No computer names found to extract information" -f $SubSite.Name) -Severity 2
                }
                else{
                    $namingConvention = $namingConvention | Group-Object CC, SiteName
                    $CC = $namingConvention | Select-Object Name, Count | Sort-Object Count -Descending | Select-Object Name -First 1
                
                    $OUShort = $CC.Name.Substring(0,4)
                    $OUName = $CC.Name.Substring(6,$CC.Name.Length-6)

                    if($SiteVsCC.ContainsKey($OUShort)){
                        $SiteVsCC.$OUShort += $OUName
                    }
                    else{
                        $SiteVsCC.Add($OUShort, @($OUName))
                    }
                }
            }
            catch{
                $SiteVsOUDN.$($SubSite.Name) = $DefaultOUtoMove
                _LogMessage -LogFullPath $LogPath -tMessage $("{0} - Default OU will be used for this - '$ComputersOU' Sub OU Not Found : {1}" -f $SubSite.Name, $_.exception.message) -Severity 2
            }
        }

        if(-not (Test-Path $ComputerList)){
            _LogMessage -LogFullPath $LogPath -tMessage "$ComputerList file not found"-Severity 2
        }
        else{
            [array]$Computers = Get-Content -Path $ComputerList -ErrorAction Stop

            foreach($computer in $Computers){
                try{
                    $ADObject = $computer.Trim()

                    $objComputer = Get-ADComputer $ADObject -Properties Description -ErrorAction Stop | Select-Object -First 1

                    $curComputerCC = ([string]$objComputer.Name).Substring(0,4)

                    [string]$OUtoMove = $DefaultOUtoMove
                    [bool]$DescMatch = $false
                    if(-not [string]::IsNullOrEmpty($objComputer.Description)){
                        if($objComputer.Description -imatch $($SiteVsOUDN.Keys -join "|")){
                            $OUtoMove = $SiteVsOUDN.$($Matches[0])
                            if(-not [string]::IsNullOrEmpty($OUtoMove)){
                                $DescMatch = $true
                                #"Desc: $($objComputer.Description)"
                            }
                        }
                    }

                    if(-not $DescMatch){
                        if($SiteVsCC.ContainsKey($curComputerCC)){
                            $tmpOUtoMove = ""
                            $tmpOUtoMove = ([array]$SiteVsCC.$curComputerCC)[0]

                            if($SiteVsOUDN.ContainsKey($tmpOUtoMove)){
                                $OUtoMove = $SiteVsOUDN.$($tmpOUtoMove)
                            }
                        }
                    }

                    #region decide which chasis is the computer
                    $OUtoMove = $(
                                    @(
                                        "$LaptopOU,$OUtoMove",
                                        "$DesktopOU,$OUtoMove"
                                    )[((isDesktop -ResourceName ([string]$objComputer.Name)) -eq $true)]
                                )
                    #endregion decide which chasis is the computer
                    
                    if($objComputer.DistinguishedName -inotmatch "gpo pilot"){
                       _LogMessage -LogFullPath $LogPath -tMessage $("{0} - Computer object is not in GPO Pilot OU" -f $computer) -Severity 1 
                    }
                    else{
                    
                        Add-ADGroupMember -Identity gg_pg_OSHarden_Policy -Members @($objComputer.DistinguishedName) -Confirm:$false
                        _LogMessage -LogFullPath $LogPath -tMessage $("{0} - Added to security group {1}" -f $computer, "gg_pg_OSHarden_Policy") -Severity 1 
                        
                        $objComputer | Move-ADObject -TargetPath $OUtoMove -Confirm:$false -ErrorAction Stop | Out-Null
                        _LogMessage -LogFullPath $LogPath -tMessage $("{0} - Moved to OU {1}" -f $computer, $OUtoMove) -Severity 1
                    }
                }
                catch{
                    _LogMessage -LogFullPath $LogPath -tMessage $("{0} - Unable to process - {1}" -f $computer, "Failed with error '$($_.exception.message)'") -Severity 3
                }
            }
        }
    }
    }

    
}
catch{
    Write-Error $_.Exception.Message
}