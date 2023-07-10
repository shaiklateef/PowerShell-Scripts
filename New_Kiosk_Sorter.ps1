
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$ImageSorter = New-Object system.Windows.Forms.Form
$ImageSorter.ClientSize = New-Object System.Drawing.Point(400, 400)
$ImageSorter.text = "Image Sorter v1.05"
$ImageSorter.TopMost = $false

#$image1 = [System.Drawing.Image]::FromFile("C:\users\Administrator\Desktop\Avengers.jpg")
#$image2 = [System.Drawing.Image]::FromFile("C:\users\Administrator\Desktop\Avengers.jpg")
#$image3 = [System.Drawing.Image]::FromFile("C:\users\Administrator\Desktop\Avengers.jpg")
#$image4 = [System.Drawing.Image]::FromFile("C:\users\Administrator\Desktop\Avengers.jpg")

$Button1 = New-Object system.Windows.Forms.Button
$Button1.text = "Sorter Controller (Production)"
$Button1.width = 150
$Button1.height = 150
$Button1.location = New-Object System.Drawing.Point(32, 26)
$Button1.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
#$Button1.Image = $image1

$Button2 = New-Object system.Windows.Forms.Button
$Button2.text = "Sorter Controller (Core Conversion)"
$Button2.width = 150
$Button2.height = 150
$Button2.location = New-Object System.Drawing.Point(228, 26)
$Button2.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$Button3 = New-Object system.Windows.Forms.Button
$Button3.text = "Restart"
$Button3.width = 150
$Button3.height = 150
$Button3.visible = $true
$Button3.location = New-Object System.Drawing.Point(32, 217)
$Button3.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

#$Button4 = New-Object system.Windows.Forms.Button
#$Button4.text = "Notepad"
#$Button4.width = 150
#$Button4.height = 150
#$Button4.location = New-Object System.Drawing.Point(228, 217)
#$Button4.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$ImageSorter.controls.AddRange(@($Button1, $Button2, $Button3))

$Button1.Add_Click({
		'C:\Program Files (x86)\IC_Sorter\Sorter.exe'
		cd 'C:\Program Files (x86)\IC_Sorter'
		.\Sorter.exe
		
	})
$Button2.Add_Click({
		#'C:\Program Files (x86)\LR10_IC_Sorter\Sorter.exe'
		cd 'C:\Program Files (x86)\LR10_IC_Sorter'
		.\SorterController_shortcut.vbs
		
		
	})

$Button3.Add_Click({ shutdown -r -f -t 00 })
#$Button4.Add_Click({ Stop-Computer -Force })
#$Button4.Add_Click({ C:\Windows\notepad.exe })

[void]$ImageSorter.ShowDialog()



