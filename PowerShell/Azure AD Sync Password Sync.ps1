####### This scirpt is used to force a full synchronization of your on prem users password with office 365########
#####Kindly replace the "Domain.com" with your domain name########
####Kindly replace TenantName.onmicrosoft.com with your tenant name ######

$Local = "mstechtalk.com"

$Remote = "UCTechTalk.onmicrosoft.com - AAD"

#Import Azure Directory Sync Module to Powershell

Import-Module AdSync

$OnPremConnector = Get-ADSyncConnector -Name $Local

Write-Output "On Prem Connector information received"

$Object = New-Object Microsoft.IdentityManagement.PowerShell.ObjectModel.ConfigurationParameter "Microsoft.Synchronize.ForceFullPasswordSync", String, ConnectorGlobal, $Null, $Null, $Null

$Object.Value = 1

$OnPremConnector.GlobalParameters.Remove($Object.Name)

$OnPremConnector.GlobalParameters.Add($Object)

$OnPremConnector = Add-ADSyncConnector -Connector $OnPremConnector

Set-ADSyncAADPasswordSyncConfiguration -SourceConnector $Local -TargetConnector $Remote -Enable $False

Set-ADSyncAADPasswordSyncConfiguration -SourceConnector $Local -TargetConnector $Remote -Enable $True 