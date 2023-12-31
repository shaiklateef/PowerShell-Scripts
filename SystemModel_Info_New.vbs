

'Set objClass = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set objClass = GetObject("winmgmts:\\.\root\cimv2")
Set colComputer = objClass.ExecQuery("Select * from Win32_ComputerSystem")
Set colItems = objClass.ExecQuery ("Select * from Win32_LogicalDisk")
Set colTimeZone = objClass.ExecQuery ("Select * from Win32_TimeZone")
Set objSysInfo = CreateObject("ADSystemInfo")
Set objComp = GetObject("LDAP://" & objSysInfo.ComputerName)

For Each objComputer in colComputer
	Wscript.Echo "System Name:  " & objComputer.Name
	Wscript.Echo "System Model: " & objComputer.Model	
	Wscript.Echo "System RAM:   " & Round((objComputer.TotalPhysicalMemory/1073741824),2)+00.01 & " GB"	
	'Wscript.Echo "System Model: " & objComputer.StandardName	
Next
For Each objItem in colItems
	If objitem.Description = "Local Fixed Disk" Then				
		Wscript.Echo "Volume Name: " & objItem.VolumeName
		Wscript.Echo "Total Hard Drive : " & Round(objItem.Size /1073741824) & " GB"
		Wscript.Echo "Available Free Space: " & Round(objItem.FreeSpace /1073741824) & " GB"		
	End If
Next

For Each objTimeZone in colTimeZone
 Wscript.Echo "Time Zone: "& objTimeZone.StandardName
Next
 Wscript.Echo "System Location: "& objSysInfo.SiteName
WScript.Quit