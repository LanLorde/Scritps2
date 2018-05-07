<#

	Auther:		Rasmus Nørgaard
	Email:		rntestermail@gmail.com
	Purpose:	Convert all *.VHDx files in a folder to *.WIM files
	Version:	1.0.0.0
	Date:		17/11-2015

#>


Write-Warning "Input the *.VHDx file path location that you want to convert.
Write the path as without the quotation marks ""C:\BlaBla\"" or as UNC-path ""\\ServerNavn\Share\folder"""
$Path = Read-Host -Prompt "Write the path here"

cls 

Write-Warning "Input the path to where you wants the converted *.Wim files to end.
Write the path as without the quotation marks ""C:\BlaBla"" or as UNC-path ""\\ServerNavn\Share\"""
$Sti = Read-Host -Prompt "Write the path here"

cls


Mkdir C:\Mount-Temp
$Mont = "C:\Mount-Temp"

cls

$Pathe = Get-ChildItem -Recurse -Path $Path | Where {($_.Extension -eq ".vhdx") -or ($_.Extension -eq ".vhd")}

ForEach ($Path in $Pathe){

	Mount-WindowsImage -ImagePath $Path.Fullname -Path "$Mont" -Index 1
	New-WindowsImage -CapturePath "$Mont" -Name "$Path" -ImagePath "$Sti\$Path.wim" -Description "$Path" -Verify
	Dismount-WindowsImage -Path "$Mont" -Discard


}
	


rm C:\Mount-Temp

Write-Host "Your VHDx files is now converted to a *.Wim file and is located at $Sti " -ForegroundColor White -BackgroundColor DarkGreen
