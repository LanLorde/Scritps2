$Arch = [IntPtr]::Size


If ($Arch -eq "8"){

	(gp HKLM:\SOFTWARE\WOW6432Node\TeamViewer).ClientID

}

Else {

	(gp HKLM:\SOFTWARE\TeamViewer).ClientID

}