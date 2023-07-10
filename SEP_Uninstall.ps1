Get-Process -Name *outlook* | Stop-Process -Force -ErrorAction SilentlyContinue

# To get the script directory path
$script:sScriptDir = $PSScriptRoot
If (([string]$sScriptDir).Length -eq 3) { $sScriptDir = $sScriptDir.functionstring(0, 2) }
$script:sSystemDrive = [System.Environment]::ExpandEnvironmentVariables("%SystemDrive%")

function QuitScript
{
	param ($iRetVal = 0)
	try
	{
		$oShell = New-Object -ComObject "WScript.Shell"
		switch ($iRetVal)
		{
			{ ($_ -eq 0) -or ($_ -eq 3010) }{
				$oShell.Logevent(4, $iRetVal) 
			}
			default{
				$oShell.LogEvent(1, $iRetVal) 
			}
		}
		
		exit $iRetVal
	}
	catch { }
}

$x64Sources = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*
foreach ($x64Source in $x64Sources)
{

	if (($x64Source.DisplayName -like 'Symantec*EndPoint*') -and (($x64Source.DisplayVersion -like '14.*') -or ($x64Source.DisplayVersion -like '12.*')))
	{
		if (Test-Path -Path "$($x64source.InstallLocation)Bin\SymCorpUI.exe")
		{
			write-host "Uninstalling x64 Symantec Endpoint Protection $($x64Source.DisplayVersion) ... " -NoNewline
			
			Start-Process "$sScriptDir\SEPUninstaller_1.0\SEP_Service_stop_previous.cmd" -ArgumentList "/S"
            Start-Process "$sScriptDir\SEPUninstaller_1.0\SEP_Service_stop_current.cmd" -ArgumentList "/S"

            $sSMC = "HKLM:\SOFTWARE\Wow6432Node\Symantec\Symantec Endpoint Protection\SMC"
            If (Test-Path -Path $sSMC)
            {
            	Remove-ItemProperty $sSMC -Name "SmcInstData"	
            }
			$sSMCx64 = "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC"
            If (Test-Path -Path $sSMCx64)
            {
				Remove-ItemProperty  $sSMCx64 -Name "SmcInstData"	
            }
            $UninstallString = $x64Source.ModifyPath -replace "MsiExec.exe /i", "/x "
			$VersionDetail = $x64Source.DisplayVersion

            $ExecuteCMDvbs = "$sScriptDir\SEPUninstaller_1.0\SEP_Uninstall.vbs"
			Start-Process "$sSystemDrive\Windows\system32\cscript.exe" -ArgumentList "$ExecuteCMDvbs" -Windowstyle Hidden
			$ExecuteCMD = "$UninstallString" + " SYMREBOOT=ReallySuppress /qb-! /norestart /l*xv C:\FRB\Logs\Symantec_EndpointProtection_$($VersionDetail)x64_Uninstall.log"

            Start-Process "msiexec.exe" -ArgumentList $ExecuteCMD -wait
            write-host "Starting Symantec Endpoint Protection Uninstall of  ($($x64Source.DisplayName)) and Version: ($($x64Source.DisplayVersion))..."
            
            Remove-Item "$sSystemDrive\ProgramData\Microsoft\Windows\Start Menu\Programs\Symantec Endpoint Protection" -Recurse
			
			Get-ChildItem HKLM:\SOFTWARE\FRB\Applications -Recurse | Where-Object { $_.PSChildName -like 'Symantec_EndpointProtection*' } | Remove-Item -Force
			Get-ChildItem HKLM:\SOFTWARE\FRB\Applications -Recurse | Where-Object { $_.PSChildName -like 'Symantec_SylinkDrop*' } | Remove-Item -Force
			write-host "Done!" -ForegroundColor Yellow
			#QuitScript 3010
		}
	}
}