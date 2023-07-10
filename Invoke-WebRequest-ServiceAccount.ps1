$Wcl = new-object System.Net.WebClient



$Wcl.Headers.Add(“user-agent”, “PowerShell Script”)



$Wcl.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials



$Wcl=New-Object System.Net.WebClient



$Creds=Get-Credential



$Wcl.Proxy.Credentials=$Creds