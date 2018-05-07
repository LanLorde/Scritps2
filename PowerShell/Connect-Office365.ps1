<#
	.SYNOPSIS
		Connects to Office 365 Exhcange.


	.NOTES
		Author:BK

#>
Function Connect-Office365 { 
	
	
	Try {
		
        $UserName = Read-Host Enter Email Address
        $Password = Read-Host Enter Password -AsSecureString
        $UserCredential = New-Object System.Management.Automation.PSCredential($UserName, $Password)
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
		Import-PSSession $Session -DisableNameChecking | Out-Null
	}

	#Catch [System.Management.Automation.ParameterBindingValidationException] {    
		#Write-Error "Please enter credentials"
		#BREAK
	#}

	#Catch [System.Management.Automation.Remoting.PSRemotingTransportException] {
		#Write-Host "Incorrect Login Info" -ForegroundColor Red
		#BREAK
	#}

	Catch {
		Write-Host "Incorrect Login Info" -ForegroundColor Red
		BREAK
	}
	
	
	Write-Host "Successfully Connected" -BackgroundColor Yellow -ForegroundColor Black	
}