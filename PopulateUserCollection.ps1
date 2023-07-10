$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
$macCsvPath = $dir + "\PopulateUserCollection.csv"

$macCsvList = import-Csv -Path $macCsvPath -Header "EmailAddress"
#echo $macCsvList

Import-Module -name "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
cd FR1: 

$sccmServerName = 'DC1SCCMPS02V'

$Namespace = 'root\sms\site_FR1'

foreach($Item in $macCsvList){
	if( $Item.EmailAddress -ne $null-and $Item.EmailAddress -ne ""){

       $EmailAddress = $Item.EmailAddress
		
	   $CollectionName = "Install Ensenta EnsentaURL (User)"

       $ResourceID = Get-WmiObject -computername $sccmServerName -namespace $namespace -query (

       "SELECT ResourceID FROM SMS_R_User WHERE Mail = '$EmailAddress'")

       $ResourceId = $ResourceID.ResourceID
		
		Add-CMUserCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId $ResourceId
		
		echo $($EmailAddress + " added to " + $CollectionName)
	}
}

Invoke-CMUserCollectionUpdate -Name $CollectionName
echo $($CollectionName + " updated successfully.")

echo $("Script complete")
cd C: