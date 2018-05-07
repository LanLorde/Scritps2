$UserCredential = Get-Credential
$FQDN = Read-Host "Enter FQDN of Exhcange Server"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$FQDN/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session