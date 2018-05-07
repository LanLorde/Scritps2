Param(
    [string]$newpassword
)

$ErrorActionPreference = "Stop"

# Get Computer Name
$computername = "$env:computername"

# Create computer object
$computer = [ADSI]"WinNT://$computerName,computer"

# Get local users list
$userList = $computer.psbase.Children | Where-Object { $_.psbase.schemaclassname -eq 'user' } 

foreach ($user in $userList)
{
	# Create a user object in order to get its SID
	$userObject = New-Object System.Security.Principal.NTAccount($user.Name)
	$userSID = $userObject.Translate([System.Security.Principal.SecurityIdentifier])
	
	# Look for local "administrator" SID 
	if(($userSID.Value.substring(0,6) -eq "S-1-5-") -and ($userSID.Value.substring($userSID.Value.Length - 4, 4) -eq "-500"))
	{
        # change password of local administrator account
        $Error.Clear() 
        try
        {
            $user.SetPassword($newpassword)
        }
        catch [System.UnauthorizedAccessException]
        {
            Write-Host "Access Denied" -ForegroundColor Red
        }
			
	}
}
