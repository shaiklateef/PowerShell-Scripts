net stop wuauserv

net stop bits

rd /s /q c:\windows\SoftwareDistribution

rd /s /q C:\Windows\SoftwareDistribution.bak

net start wuauserv

net start bits