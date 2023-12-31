

$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
$macCsvPath = $dir + "\ComputerList.csv"

$macCsvList = import-Csv -Path $macCsvPath -Header "DesktopName"
#echo $macCsvList

Import-Module -name "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
cd FR1: 

foreach($item in $macCsvList){
	if( $item.DesktopName -ne $null-and $item.DesktopName -ne ""){
		
		$CollectionName = "Patch Deployment - MS14-036 - KB2881013"
		
		Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceID $(get-cmdevice -name $item.DesktopName).ResourceID
		
		echo $($item.DesktopName + " added to " + $CollectionName)
	}
}

Invoke-CMDeviceCollectionUpdate -Name $CollectionName
echo $($CollectionName + " updated successfully.")

echo $("Script complete")
cd C: