Function Connect-LocalExchange {

Try {$UserCredential = Get-Credential}

    Catch {
        Write-Host "Incorrect username or password" -ForegroundColor Red
        break
    }
    
$FQDN = Read-Host "Enter FQDN of Exhcange Server"
Try {   
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$FQDN/PowerShell/ -Authentication Kerberos -Credential $UserCredential
}
    Catch {
        Write-Host "Exchange Server FQDN incorrect or you are not authorized to complete this action."
        break
    }

Import-PSSession $Session -DisableNameChecking | Out-Null
Write-Host "Successfully connected to $FQDN" -BackgroundColor Green

}