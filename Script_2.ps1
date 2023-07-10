
# Name		: PST file copy to n/w share utility
# Version 	: 1.1
# Creation date	: 23 Feb 2016
# Last Modified : 26 Feb 2016
# Authur 	: Arun Srinivasan
# Purpose 	: To scan for PST files and copy to remote location
# Applies	: Windows 7 / 8 / 2008 / 2012

<#
1.1 Updated folder structure creation to match the original PST file location in remote share
#>


#Define file extension and path to scan for files with specified extn
$Extensions = "*.pst"
$PSTdestination = "\\dc1ngdfile01v\packages\Copy_PST_Files" #(e.g. "\\servername\e$" or "\\servername\sharedfolder")


#Search for files and store the names
$PSTsoufiles = $null
#[array]$PSTsoufiles = Get-ChildItem -path "C:\PST-files" -Filter $Extn -recurse| select Name, fullname, length, lastwritetime

$Drives =  gwmi -query "select * from Win32_Volume where DriveType='3'" | Select -expandproperty DriveLetter

Foreach ($Drive in $Drives)
{
$drivepath = $Drive + "\*"

Foreach ($extn in $Extensions)
{ [array]$PSTsoufiles += Get-ChildItem -path $drivepath -Filter $Extn -recurse| select Name, fullname, length, lastwritetime }

}

write-host "`r`n"


#Function to Copy PST files
If ([Bool]($PSTsoufiles -ne $null))
{

For($i = 0; $i -le ($PSTsoufiles.length-1); $i++)
{

$len = $null
$PSTdestinationpath1 = $null

$st = $PSTsoufiles[$i].fullname.IndexOf("\") + 1
$en = $PSTsoufiles[$i].fullname.IndexOf(".",$st)
$len1 = $en - $st
$len = ($PSTsoufiles[$i].fullname.Substring($st, $len1)).split("\")

For($j = 0; $j -le ($len.length-2); $j++)
{ $PSTdestinationpath1 += $len[$j] + "\" }

$PSTdestinationpath = $PSTdestination + "\" + $env:username + "\" + $PSTdestinationpath1

If(!(Test-Path -Path $PSTdestinationpath))
{ New-Item -ItemType directory -Path $PSTdestinationpath | Out-Null }

$pstsouname = $null
$pstsoupath = $null
$pstsousize = $null
$pstsoumod = $null
$PSTsoulen = $null
$extn1 = $null

$PSTdesname = $null
$PSTdespath = $null
$PSTdeslen = $null
$PSTdesmod = $null

$pstsouname = $PSTsoufiles[$i].Name
$pstsoupath = $PSTsoufiles[$i].Fullname

#Defining the file name for destination
#If ($i -eq "0")
#{

#$extn1 = $pstsoupath.split(".")[-1]
#$PSTdesname = $env:username + "." + $extn1
#$PSTdespath = $PSTdestinationpath + "\" + $PSTdesname
#}
#else
#{
#$extn1 = $pstsoupath.split(".")[-1]
#$PSTdesname = $env:username + "_" + $i + "." + $extn1
#$PSTdespath = $PSTdestinationpath + "\" + $PSTdesname
#} #If condition close ($i -eq "0")

$PSTdespath = $PSTdestinationpath + $pstsouname

Write-host "Initiating copy process for '$pstsouname'. Please wait..`r`n" -fore yellow
copy-item -path $PSTsoupath -Destination $PSTdespath -force

$PSTsoulen = Get-childItem -path $PSTsoupath | select -expandproperty length
$PSTdeslen = Get-childItem -path $PSTdespath | select -expandproperty length

$PSTsoumod = Get-childItem -path $PSTsoupath | select -expandproperty lastwritetime
$PSTdesmod = Get-childItem -path $PSTdespath | select -expandproperty lastwritetime

If ([bool]($PSTsoulen -eq $PSTdeslen) -and [bool]($PSTsoumod -eq $PSTdesmod))
{ write-host "PST file(s) copied successfully from '$PSTsoupath' to '$PSTdespath'`r`n" -fore green }
else
{ write-host "PST file(s) failed to copy from '$PSTsoupath' to '$PSTdespath'`r`n" -fore red }

} #Forloop close ($i = 0; $i -le ($PSTsoufiles.length-1); $i++)

}
else
{
write-host "No PST files available in the given path`r`n" -fore yellow
} #If condition close ([Bool]($PSTsoufiles -ne $null) -and ($PSTsoufiles -ne "NA") -and (!$error))


#End of script..
