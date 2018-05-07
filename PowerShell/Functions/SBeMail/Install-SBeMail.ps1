$ModuleName = "SBeMail"

$Destination = "$env:ProgramFiles\WindowsPowerShell\Modules\$ModuleName"
If ( -not (Test-Path -Path $Destination)) {
    New-Item -Path $Destination -ItemType directory -force -Confirm:$false
}
$Path = Split-Path $script:MyInvocation.MyCommand.Path
Copy-Item -Path "$Path\$ModuleName.psm1","$Path\$ModuleName.psd1" -Destination $Destination -Force -Confirm:$false
Unblock-File "$Destination\$ModuleName.psm1","$Destination\$ModuleName.psd1" -Confirm:$false 
Remove-Module $ModuleName -ErrorAction SilentlyContinue
Import-Module $ModuleName -DisableNameChecking
Get-Module