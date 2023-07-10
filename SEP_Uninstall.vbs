' To Automate SEP Uninstall, by passing Uninstall Password

Set oShell = CreateObject("WScript.Shell")
Do Until SEP_Window = True
	SEP_Window = oShell.AppActivate("Please enter the uninstall password:")
	WScript.Sleep 1000
Loop
oShell.SendKeys "symantec{enter}"