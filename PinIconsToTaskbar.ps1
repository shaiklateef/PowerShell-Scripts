<#  
.SYNOPSIS  
    Pin icons to the taskbar or unpin them.

.DESCRIPTION  
    The script may be used in it-environments (e.g. at schools), where users work with a mandatory profile (ntuser.man)
	This script pins icons (=shortcuts) to the taskbar.
	There are two mandatory parameters: <PinIconsFolder> <Action verb>
	<PinIconsFolder>  That's the folder, where you collect the icons (=shortcuts).
                      This folder path can be a local path or an unc-path. 
                      This folder might be placed in the user's homefolder,
                      so the user could manage the additional icons (=shortcuts) in the taskbar himself.
	<Action verb>     The verb has to be translated to your language.
                      english:         "Pin to Taskbar"
                      german:          "An Taskleiste anheften"
                      your language :  " ....... "
	You can find the verb in the shortcuts context menu.
	Because icons on a network-path can't be pinned to the taskbar for security reasons,
    the script copies the icons to a temporary path: LOCALAPPDATA\PinIcons\*.*
    
.NOTES  
    File Name      : PinIconsToTaskbar.ps1 
    Author         : Gerry Bammert, 11.03.2012
    Prerequisite   : PowerShell V2, Windows 7
    Copyright 2012 : There's no copyright
                     Use it, modify it ... and post it back			 
.LINK  
    Script is published at:
    http://powershell.com/cs/media/default.aspx 
.EXAMPLE  
    Example 1
	
	PinIconsToTaskbar.ps1 <PinIconsFolder> <Action verb>
	
	==> change the verb to the appropriate version of your language
	
	English verbs:
	PinIconsToTaskbar.ps1 "\\server1\homes$\$username\PinIconsFolder" "Pin to Taskbar"
	PinIconsToTaskbar.ps1 "\\server1\homes$\$username\PinIconsFolder" "Unpin from Taskbar"
	
	German verbs:
	PinIconsToTaskbar.ps1 "\\server1\homes$\$username\PinIconsFolder" "An Taskleiste anheften"
	PinIconsToTaskbar.ps1 "\\server1\homes$\$username\PinIconsFolder" "Von Taskleiste lösen"

.EXAMPLE    
    Example 2
	
	PinIconsToTaskbar.ps1 <PinIconsFolder> <Action Verb>
	
	==> change the verb to the appropriate version of your language
	
	English verbs:
	PinIconsToTaskbar.ps1 "Path to the PinIconsFolder" "Pin to Start Menu"
	PinIconsToTaskbar.ps1 "Path to the PinIconsFolder" "Unpin from Start Menu"
	
	German verbs:
	PinIconsToTaskbar.ps1 "Path to the PinIconsFolder" "An Startmenü anheften"
	PinIconsToTaskbar.ps1 "Path to the PinIconsFolder" "Von Startmenü lösen"	
#>


param([string]$PinIconsFolder = $(Throw "Missing path to <PinIconsFolder>!"),
      [string]$ActionVerb = $(Throw "Missing ActionVerb e.g. `"Pin to taskbar`""))

###### Test-Code ###################################################
# # param([string]$PinIconsFolder = "\\127.0.0.1\d$\temp\PinIcons",
# #      [string]$ActionVerb = "Von Taskleiste lösen")
	   

function Add-IconToTaskbar
{  
    param([Parameter(ValueFromPipelineByPropertyName=$true)]  
    [Alias('LinkPath')]  
    [Alias('FileName')]  
    $Path)
    begin
    {
        $shell = New-Object -ComObject Shell.Application
    }  
    process
    {  
        $parent = Split-Path $Path  
        $child = Split-Path $Path -Leaf  
        $folder = $shell.NameSpace($parent)  
        $file = $folder.ParseName($child)
		$allVerbs = $file.Verbs()
		foreach ($verb in $allVerbs)
		{
			$verbname = ($verb.name).Replace("&","")
			if ($verbname -eq $ActionVerb)
			{
				$verb.DoIt()
			}
		}
    }  
} 

#########################################################################
##    MAIN PROGRAM
#########################################################################

##### Code for testing different paths ############ 
#$PinIconsFolder = "D:\temp\PinIcons"
#$PinIconsFolder = "\\127.0.0.1\d$\temp\PinIcons"
##### Code for testing different verbs ############
#$ActionVerb = "An Taskleiste anheften"

$LocalAppData = (Get-ChildItem ENV:LOCALAPPDATA).value
$tempPinIconsFolder = $LocalAppData + "\PinIcons"
if (!(Test-Path ($tempPinIconsFolder)))
{
	New-Item -path $tempPinIconsFolder -type directory
}
$count = 0
$PinIcons = Get-Childitem -path $PinIconsFolder
$count = $PinIcons.Count
if ($count -gt 0)
{
	foreach ($icon in $PinIcons)
	{
		Copy-Item $icon.Fullname $tempPinIconsFolder -Force
	}
	$tempPinIcons = Get-Childitem -path $tempPinIconsFolder
	foreach ($icon in $tempPinIcons)
	{
		$ic = $icon.FullName
		Add-IconToTaskbar $ic
	}
}





