Function Connect-EOP {
$UserCredential = Get-Credential


Try{
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
}
	Catch [System.Management.Automation.Remoting.PSRemotingTransportException] {
			$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
		}


	
Import-PSSession $Session -DisableNameChecking | Out-Null
}