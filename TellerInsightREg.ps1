#Copy-Item -Path "C:\temp\Tellerfile.csv" -Destination C:\temp\Tellerfile.csv
$Sources = import-csv C:\Temp\TellerInsight2File.csv
#$CurrentMachine = pcid
$balancingwksID = 0
Foreach ($Source in $Sources)
{
$Compname = $Source.pcid
If ($compname -eq $CurrentMachine)
{
$WID = $Source.WorkstationID
New-Item 'HKLM:\SOFTWARE\WOW6432Node\FRBTellerInsightConfig' -Force
Set-ItemProperty HKLM:\SOFTWARE\WOW6432Node\FRBTellerInsightConfig -Name WorkstationID -Value $WID -type dword -FORCE
Set-ItemProperty HKLM:\SOFTWARE\WOW6432Node\FRBTellerInsightConfig -Name BalancingWorkstationID -Value $balancingwksID -type dword -Force
}
}
#Del C:\Temp\file.csv
#New-Item 'HKLM:\SOFTWARE\WOW6432Node\FRBTellerInsightConfig' -Force | Set-ItemProperty HKLM:\SOFTWARE\WOW6432Node\FRBTellerInsightConfig -Name Workstation_ID -Value $Compname -type dword -Force | Out-Null
cls
Rename-Item "c:\Program Files\Citrix\ICAService\WFAPI.dll" WFAPI.dll.bak