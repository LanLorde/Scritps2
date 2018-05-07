[CmdletBinding()]
Param (
	[Parameter(Mandatory=$True,Position=0)]
	[String]$SourceFile,
	[Parameter(Mandatory=$True,Position=1)]
	[String]$ConvertTo
)
$ErrorActionPreference = "SilentlyContinue"
If ($Error) {
	$Error.Clear()
}
If (!(Test-Path $SourceFile)) {
	Write-Host
	Write-Host "`t Cannot Perform File Conversion." -ForegroundColor "Red"
	Write-Host "`t $SourceFile Not Found in this computer." -ForegroundColor "Red"
	Write-Host
	Exit
}
Write-Host
Write-Host "`t Working. Please wait ... " -ForegroundColor "Yellow"
$OutputDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$OutputDir = $OutputDir.Trim()
If (!(Test-Path "$OutputDir\OutPutResult" -PathType Container)) {
	New-Item -Path "$OutputDir\OutPutResult" -ItemType Directory | Out-Null
}
$ResultDir = "$OutputDir\OutPutResult"
$TargetName = (Get-Item $SourceFile).BaseName
$TargetName = $TargetName.Trim()
$TargetName = "$ResultDir\$TargetName.$ConvertTo"
If (Test-Path $TargetName) {
	Remove-Item $TargetName
}
## -- Load The Required Assemblies And Get Object Reference 
[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null 
$ThisImage = New-Object System.Drawing.Bitmap($SourceFile)
$ThisImage.RotateFlip("RotateNoneFlipNone")
## -- Save The Image In Desired File Format 
$ThisImage.Save($TargetName, $ConvertTo)
If ($Error) {
	Write-Host "`t ERROR -- File conversion FAILED !!" -ForegroundColor "Red"
	Write-Host "`t $Error" -ForegroundColor "Red"
	$Error.Clear()
}
Else {
	If (Test-Path $TargetName) {
		Write-Host "`t File conversion completed successfully !!" -ForegroundColor "Yellow"
		Write-Host "`t Check $TargetName" -ForegroundColor "Yellow"
	}
}
Write-Host