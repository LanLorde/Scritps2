# Gets time stamps for all User in the domain that have NOT logged in since after specified date
# Mod by Tilo 2013-08-27
import-module activedirectory 
$domain = Get-ADDomain | Select Name
$DaysInactive = Read-Host "Enter number of days without a login"
$time = (Get-Date).Adddays(-($DaysInactive))
 
# Get all AD User with lastLogonTimestamp less than our time and set to enable
Get-ADUser -Filter {LastLogonTimeStamp -lt $time -and enabled -eq $true} -Properties LastLogonTimeStamp |
 
# Output Name and lastLogonTimestamp into CSV
select-object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} | Sort Name