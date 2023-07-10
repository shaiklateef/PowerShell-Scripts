#************************************************************************************************************************
$ErrorActionPreference = "SilentlyContinue"
Set-ExecutionPolicy remotesigned -Force $ErrorActionPreference
Clear-Host
Write-Host "*******************************************************************************"  -foregroundcolor "Green"
Write-Host "File Name   : Check_SCCM_Required_Ports.ps1"                      -foregroundcolor "Green"                                                                                         
Write-Host "Purpose     : Check_SCCM_Required_Ports"                                -foregroundcolor "Green"                        
Write-Host "Version           : 1.0"                                                            -foregroundcolor "Green"
Write-Host "Requires    : PowerShell V2"                                                  -foregroundcolor "Green"    
Write-Host "*******************************************************************************"  -foregroundcolor "Green"
#******************* User Needs to be Modify the Designation Server Name and Ports Inforamtion *****************
$ServerName = "NP1SCCMDP01V.CORP.FRBNP1.COM","NP2SCCMDP01V.CORP.FRBNP2.COM","NP3SCCMDP01V.CORP.FRBNP3.COM","NPGSSCCMDEV01V.CORP.FRBNP1.COM","NPGSSCCMPS02V.FRBNPGS.COM","NPGSSCCSQLGE1V.FRBNPGS.COM","NPGSWSUS02V.FRBNPGS.COM"       # SCCM Site ServerName
$Ports = "80","135","389","443","445","636","1433","3268","3269","4022","5985","5986","8530","8531","10123","60000","DYNAMIC","ICMP"     # SCCM Distribution Point Required Ports"
#***************************************************************************************************************
$OutputPath = split-path -parent $MyInvocation.MyCommand.Definition
#$OutputPath = "C:\Temp\Check_SCCM_Required_Ports"
#***************************************************************************************************************
Remove-Item "$OutputPath\Check_SCCM_Required_Ports.CSV" -Force   #Remove Old Report
$Report= "$OutputPath\Check_SCCM_Required_Ports.CSV"             #Create New Report
$Logfile = "$OutputPath\Check_SCCM_Required_Ports.log"           #Create Log Entry
Add-Content $Logfile -Value "****************** Start Time: $(Get-Date) *******************"
Write-Host "****************** Start Time: $(Get-Date) *******************"
Hostname
Write-Host "ServerName,PortNo,Status" -foregroundcolor "Yellow"
Add-Content $logfile -Value "ServerName,PortNo,Status"
Add-Content $Report -Value "ServerName,PortNo,Status"
foreach ($Server in $ServerName)
{
      foreach ($Port in $Ports)
      {
            $Status = "Unknown"
        $Socket = New-Object Net.Sockets.TcpClient    # Create a Net.Sockets.TcpClient object to use for # Checking for open TCP ports.
            $ErrorActionPreference = 'SilentlyContinue' # Suppress error messages         
        $Socket.Connect($Server, $Port)                     # Try to connect
            $ErrorActionPreference = 'Continue'             # Make error messages visible again        
        If($Socket.Connected)                                     # Determine if we are connected.
            {
                  $Status = "Opened"
            Write-Host "$Server,$Port,$Status" -foregroundcolor "Green"
                  Add-Content $logfile -Value "$Server,$Port,$Status"
                  Add-Content $Report -Value "$Server,$Port,$Status"
            $Socket.Close()
        }
        else
            {
                  $Status = "Closed"
            Write-Host "$Server,$Port,$Status" -foregroundcolor "Red"
                  Add-Content $logfile -Value "$Server,$Port,$Status"
                  Add-Content $Report -Value "$Server,$Port,$Status"
        }
        # Apparently resetting the variable between iterations is necessary.
        $Socket = $null       
    }
}
Add-Content $Logfile -Value "****************** End Time: $(Get-Date) *******************"
Write-Host "****************** End Time: $(Get-Date) *******************"
Pause
#************************************************************************************************************************