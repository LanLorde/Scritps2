<#
Created:  January 18, 2014
Version:  2.0
Description:  On login, launch Internet Explorer to a specified URL, and then open an executable file.
Notes:  If the workstation is running Windows 7, login as the user, open "GPEDIT.MSC", and navigate to the following location.

User Configuration\Windows Settings\Scripts (Logon/Logoff)

Next, open "Logon" and go to the "PowerShell Scripts" tab.  Then, add this script to it.  However, first modify 
the following variables.  You can edit it in a text editor or in the PowerShell ISE.

$Url
$Username
$Password
$Executable

IMPORTANT:  Before running the script, open the web browser and go to the web site that will be specified in this 
script and right click in the username field and select Inspect Element.  Locate the "Input Name" and then copy and 
paste it in place of "UsernameElement" located at the end of this script (line 61).  Then, inspect the password field, 
and copy and paste the "Input Name" in place of "PasswordElement" located at the end of this script (line 62).  Then, 
right click on the submit button on the URL and locate the Input Name and then copy paste it in place of "SubmitElement"
at the end of this script (line 63).

#>

# Edit this to be the URL or IP address of the site to launch on login

$Url = “www.microsoft.com”

# Edit this to be the username

$Username=”name@microsoft.com”

# Edit this to the corresponding password

$Password=”password”

# Edit this to be the path to the executable.  Include the executable file name as well.

$Executable = "c:\windows\system32\notepad.exe"

# Invoke Internet Explorer

$IE = New-Object -com internetexplorer.application;
$IE.visible = $true;
$IE.navigate($url);

# Wait a few seconds and then launch the executable.

while ($IE.Busy -eq $true)

{

Start-Sleep -Milliseconds 2000;

}

# The following UsernameElement, PasswordElement, and LoginElement need to be modified first.  See the notes at the top
# of the script for more details.

$IE.Document.getElementById(“UsernameElement”).value = $Username
$IE.Document.getElementByID(“PasswordElement”).value=$Password
$IE.Document.getElementById(“SubmitElement”).Click()

while ($IE.Busy -eq $true)

{

Start-Sleep -Milliseconds 2000;

}

Invoke-Item $Executable