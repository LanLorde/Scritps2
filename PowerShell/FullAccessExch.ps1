Write-Host "Give User1 Full Access Rights to User2's Exchange Mailbox"
$User1 = Read-Host "Enter User1"
$User2 = Read-Host "Enter User2"
Add-MailboxPermission -Identity $User2 -User $User1 -AccessRight FullAccess