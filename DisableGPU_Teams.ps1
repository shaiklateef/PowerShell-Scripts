param(
# Define parameters and values
#[string]$newWebLanguage="en-au",
[bool]$newDisableGpu=$true,
[string]$desktopConfigFile=“$env:userprofile\\AppData\Roaming\Microsoft\Teams\desktop-config.json”,
[string]$cookieFile="$env:userprofile\\AppData\Roaming\Microsoft\teams\Cookies",
[string]$registryPath="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
[string]$registryDisplayName="Microsoft Teams",
[string]$processName="Teams"
)

#Check if Teams is installed
$registryPathCheck = Get-ChildItem -Path $registryPath -Recurse | Get-ItemProperty | Where-Object {$_.DisplayName -eq $registryDisplayName } -ErrorAction SilentlyContinue
#Check if Teams process is running
$processCheck = Get-Process $processName -ErrorAction SilentlyContinue
#Read the Teams desktop config file and convert from JSON
$config = (Get-Content -Path $desktopConfigFile | ConvertFrom-Json -ErrorAction SilentlyContinue)
#Check if required parameter value is already set within Teams desktop config file
$configCheck = $config | where {($_.appPreferenceSettings.disableGpu -ne $newDisableGpu)} -ErrorAction SilentlyContinue
#Check if Teams cookie file exists
$cookieFileCheck = Get-Item -path $cookieFile -ErrorAction SilentlyContinue

#1-If Teams is installed ($registryPathCheck not null)
#2-If Teams desktop config settings current value doesn't match parameter value ($configCheck not null)
#3-If Teams process is running ($processCheck not null)
#4-Then terminate the Teams process and wait 5 seconds
if ($registryPathCheck -and $configCheck -and $processCheck)
{
    Get-Process $processName | Stop-Process -Force
    Start-Sleep 5
}

#Check if Teams process is stopped
$processCheckFinal = Get-Process $processName -ErrorAction SilentlyContinue

#1-If Teams is installed ($registryPathCheck not null)
#2-If Teams desktop config settings current value doesn't match parameter value ($configCheck not null)
#3-Then update Teams desktop config file with new parameter value
if ($registryPathCheck -and $configCheck)
{
    $config.currentWebLanguage=$newWebLanguage
    $config.appPreferenceSettings.disableGpu=$newDisableGpu
    $config | ConvertTo-Json -Compress | Set-Content -Path $desktopConfigFile -Force

#1-If Teams process is stopped ($processCheckFinal is null)
#2-If Teams cookie file exists ($cookieFileCheck not null)
#3-Then delete cookies file

    if (!$processCheckFinal -and $cookieFileCheck)
    {
        Remove-Item -path $cookieFile -Force
    }
}