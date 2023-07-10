# Gets time stamps for all computers in the domain that have NOT logged in since after specified date 
# Mod by Tilo 2013-08-27 
import-module activedirectory  
$domain = "corp.firstrepublic.com"  
$DaysInactive = 30  
$time = (Get-Date).Adddays(-($DaysInactive)) 

# Get all AD computers with lastLogonTimestamp less than our time 
Get-ADComputer -Filter {LastLogonTimeStamp -lt $time -and OperatingSystem -notlike "*server*"} -Properties LastLogonTimeStamp,OperatingSystem | 

# Output hostname and lastLogonTimestamp into CSV 
select-object Name,@{Name="Last Logon Timestamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}},OperatingSystem | export-csv OLD_Computer.csv -notypeinformation