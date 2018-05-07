######## 
#OneDriveMapper
#Copyright:     Free to use, please leave this header intact 
#Author:        Jos Lieben (OGD)
#Company:       OGD (http://www.ogd.nl) 
#Script help:   http://www.lieben.nu 
#Purpose:       This script maps Onedrive for Business and maps a configurable number of Sharepoint Libraries
######## 

<#
DEV TODO:
1. modify documents folder name when redirecting to root
2. spaties goed uitzoeken in DL naam
3. enterprise management per groep
#>
 
param(
    [Switch]$asTask
)

$version = "2.53"

######## 
#Changelog: 
######## 
#V1.1: Added support for ADFS 
#V1.2: Added autoProvisioning, additional IE health checks 
#V1.3: Added checks for WebDav (WebClient) service 
#V1.4: Additional checks and automatic ProtectedMode fix 
#V1.5: Added DriveLabel parameter 
#V1.6: Form display fix (GPO bug) and Driveletter label persistence 
#V1.7: Added support for forcing a specific username and/or password, dealing with non domain joined machines using ADFS and non-standard library names 
#V1.8: Removed MaxAttempts setting, added automatic detection of changed usernames, added removal of existing failed drivemapping 
#V1.8a: Added conversion to lowercase of relevant user input to prevent problems when matching url's 
#V1.9: useADFS removed: this is now autodetected. Added sharepoint direct mapping.   
#V1.9a: added checks to verify Office is installed, Sharepoint is in Trusted Sites and WebDav file locking is disabled 
#V1.9b: added check for explorer.exe running, and option to restart it after mapping 
#V1.9c: added account splitter check (for people who use the same email for O365 and their normal MS account) 
#V2.0: enhanced the explorer.exe check to look only for own processes, added an IE runnning check and an option IE kill if found running 
#v2.1: added a check for the IE FirstRun wizard and a slight delay when restarting the IE Com Object to avoid issues with clean user profiles 
#V2.1: fixes a bug in Citrix, causing processess of other users to be returned 
#V2.1: revamped the ADFS redirection detection and login triggers to prevent slow responses to cause the script to fail 
#V2.1: improved zone map issue detection to include 3 alternate locations (machine, machine gpo, user gpo) where the registry can be saved 
#V2.1: Added detection of the 'HIDDEN' attribute of the redirection container for ADFS 
#V2.2: I got tired of the differences between attributes on the login page and the instability this causes, so several methods are implemented, don't forget to set adfsWaitTime
#V2.21: More ruthless cleanup of the COM objects used
#V2.22: Comments in Dutch -> English. Parameterised the ADFS control names for those who use a customized ADFS page. Cleanup. Additional zonecheck
#V2.23: Added a check to see IF the driveletter exists, it actually maps (approximately) to the right location, otherwise it will delete the mapping and remap it to the right location
#V2.23: Added an option to stop script execution if ADFS fails and two minor bugfixes reported by Martin Revard
#V2.24: added customization for stichting Sorg in the Netherlands to map a configurable number of Sharepoint Libraries in addition to O4B
#V2.25: solution for invisible drives when running as an admin. Make sure you set $restart_explorer to $True if you have users who are admin
#V2.26: Fixed multi-domain cookies not being registered (which causes sharepoint mappings to fail while O4B mappings work fine)
#V2.27: Fixed a bug in ProtectedMode storing values as String instead of DWord and better ADFS redirection and detection of invalid zonemaps configured through a GPO and added urlOpenAfter parameter
#V2.28: Support for Auto-Acceleration in Sharepoint Online (or O4B). https://support.office.com/en-us/article/Enable-auto-acceleration-for-your-SharePoint-Online-tenancy-74985ebf-39e1-4c59-a74a-dcdfd678ef83
#V2.29: Added a login prompt option, improved error logging and switched to IHTML3
#V2.30: Support for adding mappings based on AD Group Membership, fixed a bug where the script would unneccesarily restart explorer.exe for every mapping, changed logging to both log to file AND screen
#V2.31: switched from a reghack to running as a scheduled task when running elevated
#V2.32: added support for Office 365 MFA phonecall/text message during signin (ADFS MFA not yet supported), added compatibility with UAC disabled workstations
#V2.33: added usemailInstead parameter, if set to True the script will use the user's MAIL attribute instead of the UPN to log in to Office 365, added a small fix in the method to determine a drive path accurately
#V2.34: added a check for login.microsoftonline.com's presence in the safe sites list, automatically add safe sites
#V2.35: added a check for @ in the local login name, uses that if present vs adding a domain suffix, added support for the Authenticator App when using O365 MFA, fixed a minor bug where the script could ask for the username twice
#V2.36: added setAsHomeDir, if set to True the /home option will be used in NET USE. 
#V2.36: added redirectMyDocs, if set to True the My Documents library in Windows 7/2k8R2 will be redirected to the primary Driveletter
#V2.37: support for storing / retrieving the user's password to/from a file, removed dependency on Windows 7 library modules, potentially now supports more windows OS versions
#V2.38: added warnings if explorer is not set to restart but user wants to redirect my documents, and added explorer restart in case drive is mapped correctly but still needs to be redirected
#V2.38: added a retry when starting a required IE COM object, in case of heavy load on the machine
#V2.38: fixed doing a redirect when label and driveletter don't match and the drive was already mapped correctly
#V2.38: setAsHomedir removed, this requires an AD record which may not always be present
#V2.39: Folder redirection support for Windows 10, including My Pictures, My Music and My Videos, note: this makes redirection redirect to a subfolder instead of the root by default unless you change $redirectMyDocsName
#V2.39: autodetection of kb2846960 installation status
#V2.39: added msafed=0 parameter to the login url for O365 to avoid prompts to use a Personal account (thanks for suggesting this Dimos!)
#V2.40: Office dependency removed, added automatic detection of user login in Windows 10 Azure Ad Joined devices
#V2.41: SSO when using Windows 10 Azure AD Join (userLookupMode 3)
#V2.42: Additional username detection method added for userLookupMode 3, ESR is no longer required
#V2.43: longer SpO cookie generation wait time and logging of URL
#V2.44: bugfix in username selection mechanism when using forceUserName
#V2.45: longer AutoLogon for option 3 wait time, better SpO cookie generation check (url vs timer based)
#V2.46: handle AzureAD 'additional verification required' prompt and use $adfsWaitTime to also wait for AzureAD SSO
#V2.47: added a progress bar (thanks for the example Jeffery Hicks @ https://mcpmag.com/articles/2014/02/18/progress-bar-to-a-graphical-status-box.aspx)
#V2.47: added autostart webdav client trick
#V2.48: added automatic version check
#V2.48: added slightly more robust password caching method
#V2.48: added adfs non-UPN signin option
#V2.48: added ADFS password caching
#V2.49: retry in case of 404
#V2.49: RMUnify hints
#V2.49: progress bar shows a little more detail
#V2.49: cache user login when mode is set to 4
#V2.49: fixed a small bug when asking for password
#V2.50: logging with timestamps
#V2.50: also attempt to do SSO when userlookupMode is set to 1 or 2 instead of 3 (to make AADConnect SSO go more smoothly)
#V2.50: Powershell 2 and lower-friendly method to do web requests (JosL-WebRequest)
#V2.50: Changed display of the progress bar to be less 'in your face'
#V2.50: Automatically prevent IE firstrun wizard if needed
#V2.51: Automatically remove and re-add AzureADConnect SSO registry keys to work around sso issues with non persistent cookies
#V2.51: only do KB check for < windows 10 and < IE 11 and log OS and IE version
#V2.52: almost all IE object interactions moved to functions
#V2.53: map O4B before browsing to SpO sites to prevent cookie invalidation (in response to MS changes)

######## 
#Configuration 
######## 
$configurationID       = "00000000-0000-0000-0000-000000000000" #Don't modify this, not implemented yet, coming in the next version
$domain                = "OGD.NL"                  #This should be your domain name in O365, and your UPN in Active Directory, for example: ogd.nl 
$driveLetter           = "X:"                      #This is the driveletter you'd like to use for OneDrive, for example: Z: 
$redirectMyDocs        = $False                    #will redirect mydocuments to the mounted drive, does not properly 'undo' when disabled after being enabled
$redirectMyDocsName    = "Documents"               #This is the folder to which we will redirect under the given $driveletter, leave empty to redirect to the Root (may cause odd labels for special folders in Windows)
$driveLabel            = "OGD"                     #If you enter a name here, the script will attempt to label the drive with this value 
$O365CustomerName      = "ogd"                     #This should be the name of your tenant (example, ogd as in ogd.onmicrosoft.com) 
$logfile               = ($env:APPDATA + "\OneDriveMapper_$version.log")    #Logfile to log to 
$pwdCache              = ($env:APPDATA + "\OneDriveMapper.tmp")    #file to store encrypted password into, change to $Null to disable
$loginCache            = ($env:APPDATA + "\OneDriveMapper.tmp2")    #file to store encrypted login into, change to $Null to disable
$dontMapO4B            = $False                    #If you're only using Sharepoint Online mappings (see below), set this to True to keep the script from mapping the user's O4B (the user does still need to have a O4B license!)
$debugmode             = $False                    #Set to $True for debugging purposes. You'll be able to see the script navigate in Internet Explorer 
$userLookupMode        = 1                         #1 = Active Directory UPN, 2 = Active Directory Email, 3 = Azure AD Joined Windows 10, 4 = query user for his/her login, 5 = use workstation login name (deprecated)
$AzureAADConnectSSO    = $True                     #if set to True, will automatically remove AzureADSSO registry key before mapping, and then readd them after mapping. Otherwise, mapping fails because AzureADSSO creates a non-persistent cookie
$lookupUserGroups      = $False                    #Set this to $True if you want to map user security groups to Sharepoint Sites (read below for additional required configuration)
$forceUserName         = ''                        #if anything is entered here, userLookupMode is ignored
$forcePassword         = ''                        #if anything is entered here, the user won't be prompted for a password. This function is not recommended, as your password could be stolen from this file 
$restartExplorer       = $False                    #Set to $True if you're having any issues with drive visibility
$autoProtectedMode     = $True                     #Automatically temporarily disable IE Protected Mode if it is enabled. ProtectedMode has to be disabled for the script to function 
$adfsWaitTime          = 10                        #Amount of seconds to allow for SSO (ADFS or AzureAD or any other configured SSO provider) redirects, if set too low, the script may fail while just waiting for a slow redirect, this is because the IE object will report being ready even though it is not.  Set to 0 if using passwords to sign in.
$libraryName           = "Documents"               #leave this default, unless you wish to map a non-default library you've created 
$autoKillIE            = $True                     #Kill any running Internet Explorer processes prior to running the script to prevent security errors when mapping 
$abortIfNoAdfs         = $False                    #If set to True, will stop the script if no ADFS server has been detected during login
$adfsMode              = 1                         #1 = use whatever came out of userLookupMode, 2 = use only the part before the @ in the upn
$displayErrors         = $True                     #show errors to user in visual popups
$buttonText            = "Login"                   #Text of the button on the password input popup box
$adfsLoginInput        = "userNameInput"           #change to user-signin if using Okta, username2Txt if using RMUnify
$adfsPwdInput          = "passwordInput"           #change to pass-signin if using Okta, passwordTxt if using RMUnify
$adfsButton            = "submitButton"            #change to singin-button if using Okta, Submit if using RMUnify
$urlOpenAfter          = ""                        #This URL will be opened by the script after running if you configure it
$showConsoleOutput     = $True                     #Set this to $False to hide console output
$showElevatedConsole   = $True
$sharepointMappings    = @()
$sharepointMappings    += "https://ogd.sharepoint.com/site1/documentsLibrary,ExampleLabel,Y:"
$showProgressBar       = $True                     #will show a progress bar to the user
$versionCheck          = $True                     #will check if running the latest version, if not, this will be logged to the logfile
#for each sharepoint site you wish to map 3 comma seperated values are required, the 'clean' url to the library (see example), the desired drive label, and the driveletter
#if you wish to add more, copy the example as you see above, if you don't wish to map any sharepoint sites, simply leave as is

######## 
#Required resources, it's highly unlikely you need to change any of this
######## 
$mapresult = $False 
$protectedModeValues = @{} 
$privateSuffix = "-my" 
$script:errorsForUser = ""
$maxWaitSecondsForSpO  = 5                        #Maximum seconds the script waits for Sharepoint Online to load before mapping
$o365loginURL = "https://login.microsoftonline.com/login.srf?msafed=0"
if($sharepointMappings[0] -eq "https://ogd.sharepoint.com/site1/documentsLibrary,ExampleLabel,Y:"){           ##DO NOT CHANGE THIS
    $sharepointMappings = @()
}

function log{
    param (
        [Parameter(Mandatory=$true)][String]$text,
        [Switch]$fout,
        [Switch]$warning
    )
    if($fout){
        $text = "ERROR | $text"
    }
    elseif($warning){
        $text = "WARNING | $text"
    }
    else{
        $text = "INFO | $text"
    }
    try{
        ac $logfile "$(Get-Date) | $text"
    }catch{$Null}
    if($showConsoleOutput){
        if($fout){
            Write-Host $text -ForegroundColor Red
        }elseif($warning){
            Write-Host $text -ForegroundColor Yellow
        }else{
            Write-Host $text -ForegroundColor Green
        }
    }
}

$scriptPath = $MyInvocation.MyCommand.Definition
log -text "-----$(Get-Date) OneDriveMapper v$version - $($env:USERNAME) on $($env:COMPUTERNAME) starting-----" 

###THIS ONLY HAS TO BE CONFIGURED IF YOU WANT TO MAP USER SECURITY GROUPS TO SHAREPOINT SITES
if($lookupUserGroups){
    try{
        $groups = ([ADSISEARCHER]"samaccountname=$($env:USERNAME)").Findone().Properties.memberof -replace '^CN=([^,]+).+$','$1'
        log -text "cached user group membership because lookupUserGroups was set to True"
        #####################FOR EACH GROUP YOU WISH TO MAP TO A SHAREPOINT LIBRARY, UNCOMMENT AND REPEAT BELOW EXAMPLE, NOTE: THIS MAY FAIL IF THERE ARE REGEX CHARACTERS IN THE NAME
        #    $group = $groups -match "DLG_West District School A - Sharepoint"
        #    if($group){
        #       ###REMEMBER, THE BELOW LINE SHOULD CONTAIN 2 COMMA's to distinguish between URL, LABEL and DRIVELETTER
        #       $sharepointMappings += "https://ogd.sharepoint.com/district_west/DocumentLibraryName,West District,Y:"
        #       log -text "adding a sharepoint mapping because the user is a member of $group"
        #    }   
    }catch{
        log -text "failed to cache user group membership because of: $($Error[0])" -fout
    }
}

function getElementById{
    Param(
        [Parameter(Mandatory=$true)]$id
    )
    $localObject = $Null
    try{
        $localObject = $script:ie.document.getElementById($id)
        if($localObject.tagName -eq $Null){Throw "The element $id was not found (1) or had no tagName"}
        return $localObject
    }catch{$localObject = $Null}
    try{
        $localObject = $script:ie.document.IHTMLDocument3_getElementById($id)
        if($localObject.tagName -eq $Null){Throw "The element $id was not found (2) or had no tagName"}
        return $localObject
    }catch{
        Throw
    }
}

function JosL-WebRequest{
    Param(
        $url
    )
    $maxAttempts = 3
    $attempts=0
    while($true){
        $attempts++
        try{
            $retVal = @{}
            $request = [System.Net.WebRequest]::Create($url)
            $request.TimeOut = 5000
            $request.UserAgent = "Lieben Consultancy"
            $response = $request.GetResponse()
            $retVal.StatusCode = $response.StatusCode
            $retVal.StatusDescription = $response.StatusDescription
            $retVal.Headers = $response.Headers
            $stream = $response.GetResponseStream()
            $streamReader = [System.IO.StreamReader]($stream)
            $retVal.Content = $streamReader.ReadToEnd()
            $streamReader.Close()
            $response.Close()
            return $retVal
        }catch{
            if($attempts -ge $maxAttempts){Throw}else{sleep -s 2}
        }
    }
}

function handleAzureADConnectSSO{
    Param(
        [Switch]$initial
    )
    $failed = $False
    if($script:AzureAADConnectSSO){
        if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftazuread-sso.com\autologon" -entryName "https") -eq 1){
            log -text "ERROR: https://autologon.microsoftazuread-sso.com found in IE Local Intranet sites, Azure AD Connect SSO is only supported if you let OnedriveMapper set the registry keys! Don't set this site through GPO" -fout
            $failed = $True
        }
        if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\nsatc.net\aadg.windows.net" -entryName "https") -eq 1){
            log -text "ERROR: https://aadg.windows.net.nsatc.net found in IE Local Intranet sites, Azure AD Connect SSO is only supported if you let OnedriveMapper set the registry keys! Don't set this site through GPO" -fout
            $failed = $True
        } 
        if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftazuread-sso.com\autologon" -entryName "https") -eq 1){
            log -text "ERROR: https://autologon.microsoftazuread-sso.com found in IE Local Intranet sites, Azure AD Connect SSO is only supported if you let OnedriveMapper set the registry keys! Don't set this site through GPO" -fout
            $failed = $True
        }
        if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\nsatc.net\aadg.windows.net" -entryName "https") -eq 1){
            log -text "ERROR: https://aadg.windows.net.nsatc.net found in IE Local Intranet sites, Azure AD Connect SSO is only supported if you let OnedriveMapper set the registry keys! Don't set this site through GPO" -fout
            $failed = $True
        } 
        if($failed -eq $False){
            if($initial){
                #check AzureADConnect SSO intranet sites
                if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftazuread-sso.com\autologon" -entryName "https") -eq 1){
                    $res = remove-item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftazuread-sso.com\autologon"    
                    log -text "Automatically removed autologon.microsoftazuread-sso.com from intranet sites for this user"
                }
                if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\nsatc.net\aadg.windows.net" -entryName "https") -eq 1){
                    $res = remove-item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\nsatc.net\aadg.windows.net" 
                    log -text "Automatically removed aadg.windows.net.nsatc.net from intranet sites for this user"   
                }
            }else{
                #log results, try to automatically add trusted sites to user trusted sites if not yet added
                if((addSiteToIEZoneThroughRegistry -siteUrl "aadg.windows.net.nsatc.net" -mode 1) -eq $True){log -text "Automatically added aadg.windows.net.nsatc.net to intranet sites for this user"}
                if((addSiteToIEZoneThroughRegistry -siteUrl "autologon.microsoftazuread-sso.com" -mode 1) -eq $True){log -text "Automatically added autologon.microsoftazuread-sso.com to intranet sites for this user"}   
            }
        }
    }
}

function storeSecureString{
    Param(
        $filePath,
        $string
    )
    try{
        $stringForFile = $string | ConvertTo-SecureString -AsPlainText -Force -ErrorAction Stop | ConvertFrom-SecureString -ErrorAction Stop
        $res = Set-Content -Path $filePath -Value $stringForFile -Force -ErrorAction Stop
    }catch{
        Throw "Failed to store string: $($Error[0] | out-string)"
    }
}

function loadSecureString{
    Param(
        $filePath
    )
    try{
        $string = Get-Content $filePath -ErrorAction Stop | ConvertTo-SecureString -ErrorAction Stop
        $string = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($string)
        $string = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($string)
        if($string.Length -lt 3){throw "no valid string returned from cache"}
        return $string
    }catch{
        Throw
    }
}

function versionCheck{
    Param(
        $currentVersion
    )
    $apiURL = "http://www.lieben.nu/lieben_api.php?script=OnedriveMapper&version=$currentVersion"
    $apiKeyword = "latestOnedriveMapperVersion"
    try{
        $result = JosL-WebRequest -Url $apiURL
    }catch{
        Throw "Failed to connect to API url for version check: $apiURL $($Error[0])"
    }
    try{
        $keywordIndex = $result.Content.IndexOf($apiKeyword)
        if($keywordIndex -lt 1){
            Throw ""
        }
    }catch{
        Throw "Connected to API url for version check, but invalid API response"
    }
    $latestVersion = $result.Content.SubString($keywordIndex+$apiKeyword.Length+1,4)
    if($latestVersion -ne $currentVersion){
        Throw "OnedriveMapper version mismatch, current version: v$currentVersion, latest version: v$latestVersion"
    }
}

function startWebDavClient{
    $Source = @" 
        using System;
        using System.Text;
        using System.Security;
        using System.Collections.Generic;
        using System.Runtime.Versioning;
        using Microsoft.Win32.SafeHandles;
        using System.Runtime.InteropServices;
        using System.Diagnostics.CodeAnalysis;
        namespace JosL.WebClient{
        public static class Starter{
            [StructLayout(LayoutKind.Explicit, Size=16)]
            public class EVENT_DESCRIPTOR{
                [FieldOffset(0)]ushort Id = 1;
                [FieldOffset(2)]byte Version = 0;
                [FieldOffset(3)]byte Channel = 0;
                [FieldOffset(4)]byte Level = 4;
                [FieldOffset(5)]byte Opcode = 0;
                [FieldOffset(6)]ushort Task = 0;
                [FieldOffset(8)]long Keyword = 0;
            }

            [StructLayout(LayoutKind.Explicit, Size = 16)]
            public struct EventData{
                [FieldOffset(0)]
                internal UInt64 DataPointer;
                [FieldOffset(8)]
                internal uint Size;
                [FieldOffset(12)]
                internal int Reserved;
            }

            public static void startService(){
                Guid webClientTrigger = new Guid(0x22B6D684, 0xFA63, 0x4578, 0x87, 0xC9, 0xEF, 0xFC, 0xBE, 0x66, 0x43, 0xC7);
                long handle = 0;
                uint output = EventRegister(ref webClientTrigger, IntPtr.Zero, IntPtr.Zero, ref handle);
                bool success = false;
                if (output == 0){
                    EVENT_DESCRIPTOR desc = new EVENT_DESCRIPTOR();
                    unsafe{
                        uint writeOutput = EventWrite(handle, ref desc, 0, null);
                        success = writeOutput == 0;
                        EventUnregister(handle);
                    }
                }
            }

            [DllImport("Advapi32.dll", SetLastError = true)]
            public static extern uint EventRegister(ref Guid guid, [Optional] IntPtr EnableCallback, [Optional] IntPtr CallbackContext, [In][Out] ref long RegHandle);
            [DllImport("Advapi32.dll", SetLastError = true)]
            public static extern unsafe uint EventWrite(long RegHandle, ref EVENT_DESCRIPTOR EventDescriptor, uint UserDataCount, EventData* UserData);
            [DllImport("Advapi32.dll", SetLastError = true)]
            public static extern uint EventUnregister(long RegHandle);
        }
    }
"@ 
    try{
        log -text "Attempting to automatically start the WebDav client without elevation..."
        $compilerParameters = New-Object System.CodeDom.Compiler.CompilerParameters
        $compilerParameters.CompilerOptions="/unsafe"
        Add-Type -TypeDefinition $Source -Language CSharp -CompilerParameters $compilerParameters
        [JosL.WebClient.Starter]::startService()
        log -text "Start Service Command completed without errors"
        sleep -s 1
        if((Get-Service -Name WebClient).status -eq "Running"){
            log -text "detected that the webdav client is now running!"
        }else{
            log -text "but the webdav client is still not running! Please set the client to automatically start!" -fout
        }
    }catch{
        Throw "Failed to start the webdav client :( $($Error[0])"
    }
}

function Pause{
   Read-Host 'Press Enter to continue...' | Out-Null
}

$domain = $domain.ToLower() 
$O365CustomerName = $O365CustomerName.ToLower() 
#for people that don't RTFM, fix wrongly entered customer names:
$O365CustomerName = $O365CustomerName -Replace ".onmicrosoft.com",""
$forceUserName = $forceUserName.ToLower() 
$finalURLs = @()
$finalURLs += "https://portal.office.com"
$finalURLs += "https://outlook.office365.com"
$finalURLs += "https://outlook.office.com"
$finalURLs += "https://$($O365CustomerName)-my.sharepoint.com"
$finalURLs += "https://$($O365CustomerName).sharepoint.com"
$finalURLs += "https://www.office.com"

function redirectMyDocuments{
    Param(
        $driveLetter
    )
    $dl = "$($driveLetter)\"
    $myDocumentsNewPath = Join-Path -Path $dl -ChildPath $redirectMyDocsName
    $myPicturesNewPath = Join-Path -Path $myDocumentsNewPath -ChildPath "Pictures"
    $myVideosNewPath = Join-Path -Path $myDocumentsNewPath -ChildPath "Videos"
    $myMusicNewPath = Join-Path -Path $myDocumentsNewPath -ChildPath "Music"
    #create folders if necessary
    $waitedTime = 0    
    while($true){
        try{
            if(![System.IO.Directory]::Exists($myDocumentsNewPath)){
                $res = New-Item $myDocumentsNewPath -ItemType Directory -ErrorAction Stop
                Sleep -Milliseconds 200
            }   
            if(![System.IO.Directory]::Exists($myPicturesNewPath)){
                $res = New-Item $myPicturesNewPath -ItemType Directory -ErrorAction Stop
                Sleep -Milliseconds 200
            } 
            if(![System.IO.Directory]::Exists($myVideosNewPath)){
                $res = New-Item $myVideosNewPath -ItemType Directory -ErrorAction Stop
                Sleep -Milliseconds 200
            } 
            if(![System.IO.Directory]::Exists($myMusicNewPath)){
                $res = New-Item $myMusicNewPath -ItemType Directory -ErrorAction Stop
            } 
            break
        }catch{
            sleep -s 2
            $waitedTime+=2
            if($waitedTime -gt 15){
                log -text "Failed to redirect document libraries because we could not create folders in the target path $dl $($Error[0])" -fout
                return $False              
            }      
        }
    }
    try{
        log -text "Retrieving current document library configuration"
        $lib = "$Env:appdata\Microsoft\Windows\Libraries\Documents.library-ms"
        $content = get-content -LiteralPath $lib
    }catch{
        log -text "Failed to retrieve document library configuration, will not be able to redirect $($Error[0])" -fout
        return $False
    }
    #Method 1 (works for Win7/8/2008R2)
    try{
        $strip = $false
        $count = 0
        foreach($line in $content){
            if($line -like "*<searchConnectorDescriptionList>*"){$strip = $True}
            if($strip){$content[$count]=$Null}
            $count++
        }
        $content+="<searchConnectorDescriptionList>"
        $content+="<searchConnectorDescription>"
        $content+="<isDefaultSaveLocation>true</isDefaultSaveLocation>"
        $content+="<isSupported>false</isSupported>"
        $content+="<simpleLocation>"
        $content+="<url>$myDocumentsNewPath</url>"
        $content+="</simpleLocation>"
        $content+="</searchConnectorDescription>"
        $content+="</searchConnectorDescriptionList>"
        $content+="</libraryDescription>"
        Set-Content -Value $content -Path $lib -Force -ErrorAction Stop
        log -text "Modified $lib"
    }catch{
        log -text "Failed to redirect document library $($Error[0])" -fout
        return $False
    }
    #Method 2 (Windows 10+)
    try{   
        $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "Personal" -value $myDocumentsNewPath -ErrorAction Stop        
        $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "{F42EE2D3-909F-4907-8871-4C22FC0BF756}" -value $myDocumentsNewPath -ErrorAction Stop  
        $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "My Video" -value $myVideosNewPath -ErrorAction SilentlyContinue
        $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "{35286A68-3C57-41A1-BBB1-0EAE73D76C95}" -value $myVideosNewPath -ErrorAction SilentlyContinue
        $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "My Music" -value $myMusicNewPath -ErrorAction SilentlyContinue
        $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "{A0C69A99-21C8-4671-8703-7934162FCF1D}" -value $myMusicNewPath -ErrorAction SilentlyContinue
        $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "My Pictures" -value $myPicturesNewPath -ErrorAction SilentlyContinue
        $res = Set-ItemProperty "hkcu:\software\microsoft\windows\currentversion\explorer\User Shell Folders" -Name "{0DDD015D-B06C-45D5-8C4C-F59713854639}" -value $myPicturesNewPath -ErrorAction SilentlyContinue
        log -text "Modified explorer shell registry entries"
    }catch{
        log -text "Failed to redirect document library $($Error[0])" -fout
        return $False
    }
    log -text "Redirected My Documents to $myDocumentsNewPath"
    return $True
}

function checkIfAtO365URL{
    param(
        [String]$url,
        [Array]$finalURLs
    )
    foreach($item in $finalURLs){
        if($url.StartsWith($item)){
            return $True
        }
    }
    return $False
}

#region basicFunctions
function lookupUPN{ 
    if($userLookupMode -eq 2){
        try{
            $userMail = ([ADSISEARCHER]"samaccountname=$($env:USERNAME)").Findone().Properties.mail
            if($userMail){
                return $userMail
            }else{Throw $Null}
        }catch{
            log -text "Failed to lookup email, active directory connection failed, please change userLookupMode" -fout
            $script:errorsForUser += "Could not connect to your corporate network.`n"
            abort_OM 
        }
    }
    if($userLookupMode -eq 1){
        try{ 
            $objDomain = New-Object System.DirectoryServices.DirectoryEntry 
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
            $objSearcher.SearchRoot = $objDomain 
            $objSearcher.Filter = "(&(objectCategory=User)(SAMAccountName=$Env:USERNAME))"
            $objSearcher.SearchScope = "Subtree"
            $objSearcher.PropertiesToLoad.Add("userprincipalname") | Out-Null 
            $results = $objSearcher.FindAll() 
            return $results[0].Properties.userprincipalname 
        }catch{ 
            log -text "Failed to lookup username, active directory connection failed, please change userLookupMode" -fout
            $script:errorsForUser += "Could not connect to your corporate network.`n"
            abort_OM 
        }
    }
}

function CustomInputBox(){ 
    Param(
        [String]$title,
        [String]$message,
        [Switch]$password
    )
    if($forcePassword.Length -gt 2 -and $password) { 
        return $forcePassword 
    } 
    $objBalloon = New-Object System.Windows.Forms.NotifyIcon  
    $objBalloon.BalloonTipIcon = "Info" 
    $objBalloon.BalloonTipTitle = "OneDriveMapper"  
    $objBalloon.BalloonTipText = "OneDriveMapper - www.lieben.nu" 
    $objBalloon.Visible = $True  
    $objBalloon.ShowBalloonTip(10000) 
 
    $userForm = New-Object 'System.Windows.Forms.Form' 
    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState' 
    $Form_StateCorrection_Load= 
    { 
        $userForm.WindowState = $InitialFormWindowState 
    }  
    $userForm.Text = "$title" 
    $userForm.Size = New-Object System.Drawing.Size(350,200) 
    $userForm.StartPosition = "CenterScreen" 
    $userForm.AutoSize = $False 
    $userForm.MinimizeBox = $False 
    $userForm.MaximizeBox = $False 
    $userForm.SizeGripStyle= "Hide" 
    $userForm.WindowState = "Normal" 
    $userForm.FormBorderStyle="Fixed3D" 
    $userForm.KeyPreview = $True 
    $userForm.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$userForm.Close()}})   
    $OKButton = New-Object System.Windows.Forms.Button 
    $OKButton.Location = New-Object System.Drawing.Size(105,110) 
    $OKButton.Size = New-Object System.Drawing.Size(95,23) 
    $OKButton.Text = $buttonText 
    $OKButton.Add_Click({$userForm.Close()}) 
    $userForm.Controls.Add($OKButton) 
    $userLabel = New-Object System.Windows.Forms.Label 
    $userLabel.Location = New-Object System.Drawing.Size(10,20) 
    $userLabel.Size = New-Object System.Drawing.Size(300,50) 
    $userLabel.Text = "$message" 
    $userForm.Controls.Add($userLabel)  
    $objTextBox = New-Object System.Windows.Forms.TextBox 
    if($password) {$objTextBox.UseSystemPasswordChar = $True }
    $objTextBox.Location = New-Object System.Drawing.Size(70,75) 
    $objTextBox.Size = New-Object System.Drawing.Size(180,20) 
    $userForm.Controls.Add($objTextBox)  
    $userForm.Topmost = $True 
    $userForm.TopLevel = $True 
    $userForm.ShowIcon = $True 
    $userForm.Add_Shown({$userForm.Activate();$objTextBox.focus()}) 
    $InitialFormWindowState = $userForm.WindowState 
    $userForm.add_Load($Form_StateCorrection_Load) 
    [void] $userForm.ShowDialog() 
    return $objTextBox.Text 
} 
 
function labelDrive{ 
    Param( 
    [String]$lD_DriveLetter, 
    [String]$lD_MapURL, 
    [String]$lD_DriveLabel 
    ) 
 
    #try to set the drive label 
    if($lD_DriveLabel.Length -gt 0){ 
        log -text "A drive label has been specified, attempting to set the label for $($lD_DriveLetter)"
        try{ 
            $regURL = $lD_MapURL.Replace("\","#") 
            $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($regURL)" 
            $Null = New-Item -Path $path -Value "default value" –Force -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_CommentFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromReg" -Value $lD_DriveLabel -ErrorAction SilentlyContinue
            $regURL = $regURL.Replace("DavWWWRoot#","") 
            $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($regURL)" 
            $Null = New-Item -Path $path -Value "default value" –Force -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_CommentFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromDesktopINI" -ErrorAction SilentlyContinue
            $Null = New-ItemProperty -Path $path -Name "_LabelFromReg" -Value $lD_DriveLabel -ErrorAction SilentlyContinue
            log -text "Label has been set to $($lD_DriveLabel)" 
 
        }catch{ 
            log -text "Failed to set the drive label registry keys: $($Error[0]) " -fout
        } 
 
    } 
} 

function restart_explorer{
    log -text "Restarting Explorer.exe to make the drive(s) visible" 
    #kill all running explorer instances of this user 
    $explorerStatus = Get-ProcessWithOwner explorer 
    if($explorerStatus -eq 0){ 
        log -text "no instances of Explorer running yet, at least one should be running" -warning
    }elseif($explorerStatus -eq -1){ 
        log -text "ERROR Checking status of Explorer.exe: unable to query WMI" -fout
    }else{ 
        log -text "Detected running Explorer processes, attempting to shut them down..." 
        foreach($Process in $explorerStatus){ 
            try{ 
                Stop-Process $Process.handle | Out-Null 
                log -text "Stopped process with handle $($Process.handle)" 
            }catch{ 
                log -text "Failed to kill process with handle $($Process.handle)" -fout
            } 
        } 
    } 
} 

function fixElevationVisibility{
    #check if a task already exists for this script
    if($showElevatedConsole){
        $createTask = "schtasks /Create /SC ONCE /TN OnedriveMapper /IT /RL LIMITED /F /TR `"Powershell.exe -ExecutionPolicy ByPass -File '$scriptPath' -asTask`" /st 00:00"    
    }else{
        $createTask = "schtasks /Create /SC ONCE /TN OnedriveMapper /IT /RL LIMITED /F /TR `"Powershell.exe -ExecutionPolicy ByPass -WindowStyle Hidden -File '$scriptPath' -asTask`" /st 00:00"
    }
    $res = Invoke-Expression $createTask
    if($res -NotMatch "ERROR"){
        log -text "Scheduled a task to run OnedriveMapper unelevated because this script cannot run elevated"
        $runTask = "schtasks /Run /TN OnedriveMapper /I"
        $res = Invoke-Expression $runTask
        if($res -NotMatch "ERROR"){
            log -text "Scheduled task started"
        }else{
            log -text "Failed to start a scheduled task to run OnedriveMapper without elevation because: $res" -fout
        }
    }else{
        log -text "Failed to schedule a task to run OnedriveMapper without elevation because: $res" -fout
    }
}

function MapDrive{ 
    Param( 
    [String]$MD_DriveLetter, 
    [String]$MD_MapURL, 
    [String]$MD_DriveLabel 
    ) 
    $LASTEXITCODE = 0
    log -text "Mapping target: $($MD_MapURL)" 
    try{$del = NET USE $MD_DriveLetter /DELETE /Y 2>&1}catch{$Null}
    try{$out = NET USE $MD_DriveLetter $MD_MapURL /PERSISTENT:YES 2>&1}catch{$Null}
    if($out -like "*error 67*"){
        log -text "ERROR: detected string error 67 in return code of net use command, this usually means the WebClient isn't running" -fout
    }
    if($out -like "*error 224*"){
        log -text "ERROR: detected string error 224 in return code of net use command, this usually means your trusted sites are misconfigured or KB2846960 is missing" -fout
    }
    if($LASTEXITCODE -ne 0){ 
        log -text "Failed to map $($MD_DriveLetter) to $($MD_MapURL), error: $($LASTEXITCODE) $($out) $del" -fout
        $script:errorsForUser += "$MD_DriveLetter could not be mapped because of error $($LASTEXITCODE) $($out) d$del`n"
        return $False 
    } 
    if([System.IO.Directory]::Exists($MD_DriveLetter)){ 
        #set drive label 
        $Null = labelDrive $MD_DriveLetter $MD_MapURL $MD_DriveLabel
        log -text "$($MD_DriveLetter) mapped successfully`n" 
        if($redirectMyDocs -and $driveLetter -eq $MD_DriveLetter){
            $res = redirectMyDocuments -driveLetter $MD_DriveLetter
        }
        return $True 
    }else{ 
        log -text "failed to contact $($MD_DriveLetter) after mapping it to $($MD_MapURL), check if the URL is valid. Error: $($error[0]) $out" -fout
        return $False 
    } 
} 
 
function revertProtectedMode(){ 
    log -text "autoProtectedMode is set to True, reverting to old settings" 
    try{ 
        for($i=0; $i -lt 5; $i++){ 
            if($protectedModeValues[$i] -ne $Null){ 
                log -text "Setting zone $i back to $($protectedModeValues[$i])" 
                Set-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500"  -Value $protectedModeValues[$i] -Type Dword -ErrorAction SilentlyContinue 
            } 
        } 
    } 
    catch{ 
        log -text "Failed to modify registry keys to change ProtectedMode back to the original settings: $($Error[0])" -fout
    } 
} 

function abort_OM{ 
    if($showProgressBar) {
        $progressbar1.Value = 100
        $label1.text="Done!"
        Sleep -Milliseconds 500
        $form1.Close()
    }
    #find and kill all active COM objects for IE
    try{
        $script:ie.Quit() | Out-Null
    }catch{}
    $shellapp = New-Object -ComObject "Shell.Application"
    $ShellWindows = $shellapp.Windows()
    for ($i = 0; $i -lt $ShellWindows.Count; $i++)
    {
      if ($ShellWindows.Item($i).FullName -like "*iexplore.exe")
      {
        $del = $ShellWindows.Item($i)
        try{
            $Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($del)  2>&1 
        }catch{}
      }
    }
    try{
        $Null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shellapp) 
    }catch{}
    if($autoProtectedMode){ 
        revertProtectedMode 
    } 
    handleAzureADConnectSSO
    log -text "OnedriveMapper has finished running"
    if($restartExplorer){
        restart_explorer
    }else{
        log -text "restartExplorer is set to False, if you're redirecting My Documents, it won't show until next logon" -warning
    }
    if($urlOpenAfter){Start-Process iexplore.exe $urlOpenAfter}
    if($displayErrors){
        if($errorsForUser){ 
            $OUTPUT= [System.Windows.Forms.MessageBox]::Show($errorsForUser, "Onedrivemapper Error" , 0) 
            $OUTPUT2= [System.Windows.Forms.MessageBox]::Show("You can always use https://portal.office.com to access your data", "Need a workaround?" , 0) 
        }
    }
    Exit 
} 
 
function askForPassword{
    $askAttempts = 0
    do{ 
        $askAttempts++ 
        log -text "asking user for password" 
        try{ 
            $password = CustomInputBox "Microsoft Office 365 OneDrive" "Please enter your password for Office 365" -password
        }catch{ 
            log -text "failed to display a password input box, exiting. $($Error[0])" -fout
            abort_OM              
        } 
    } 
    until($password.Length -gt 0 -or $askAttempts -gt 2) 
    if($askAttempts -gt 2) { 
        log -text "user refused to enter a password, exiting" -fout
        $script:errorsForUser += "You did not enter a password, we will not be able to connect to Onedrive`n"
        abort_OM 
    }else{ 
        return $password 
    }
}

function retrievePassword{ 
    Param(
        [switch]$forceNewPassword
    )

    if($forceNewPassword){
        $password = askForPassword
        if($pwdCache){
            try{
                $res = storeSecureString -filePath $pwdCache -string $password
                log -text "Stored user's new password to user password cache file $pwdCache"
            }catch{
                log -text "Error storing user password to user password cache file ($($Error[0] | out-string)" -fout
            }
        }    
        return $password
    }
    if($pwdCache){
        try{
            $res = loadSecureString -filePath $pwdCache
            log -text "Retrieved user password from cache $pwdCache"
            return $res
        }catch{
            log -text "Failed to retrieve user password from cache: $($Error[0])" -fout
        }
    }
    $password = askForPassword
    if($pwdCache){
        try{
            $res = storeSecureString -filePath $pwdCache -string $password
            log -text "Stored user's new password to user password cache file $pwdCache"
        }catch{
            log -text "Error storing user password to user password cache file ($($Error[0] | out-string)" -fout
        }
    }
    return $password
} 
 
function askForUserName{ 
    $askAttempts = 0
    do{ 
        $askAttempts++ 
        log -text "asking user for login" 
        try{ 
            $login = CustomInputBox "Microsoft Office 365 OneDrive" "Please enter your login name for Office 365"
        }catch{ 
            log -text "failed to display a login input box, exiting $($Error[0])" -fout
            abort_OM              
        } 
    } 
    until($login.Length -gt 0 -or $askAttempts -gt 2) 
    if($askAttempts -gt 3) { 
        log -text "user refused to enter a login name, exiting" -fout
        $script:errorsForUser += "You did not enter a login name, script cannot continue`n"
        abort_OM 
    }else{ 
        return $login 
    } 
}

function retrieveLogin{ 
    Param(
        [switch]$forceNewLogin
    )

    if($forceNewLogin){
        $login = askForUserName
        if($loginCache){
            try{
                $res = storeSecureString -filePath $loginCache -string $login
                log -text "Stored user's new login to user login cache file $loginCache"
            }catch{
                log -text "Error storing user login to user login cache file ($($Error[0] | out-string)" -fout
            }
        }
        return [String]$login.ToLower()
    }
    if($loginCache){
        try{
            $res = loadSecureString -filePath $loginCache
            log -text "Retrieved user login from cache $loginCache"
            return [String]$res.ToLower()
        }catch{
            log -text "Failed to retrieve user login from cache: $($Error[0])" -fout
        }
    }
    $login = askForUserName
    if($loginCache){
        try{
            $res = storeSecureString -filePath $loginCache -string $login
            log -text "Stored user's new login to user login cache file $loginCache"
        }catch{
            log -text "Error storing user login to user login cache file ($($Error[0] | out-string)" -fout
        }
    }
    return [String]$login.ToLower()
} 

function askForCode{ 
    $askAttempts = 0
    do{ 
        $askAttempts++ 
        log -text "asking user for SMS or App code" 
        try{ 
            $login = CustomInputBox "Microsoft Office 365 OneDrive" "Please enter the SMS or Authenticator App code you have received on your cellphone"
        }catch{ 
            log -text "failed to display a code input box, exiting $($Error[0])" -fout
            abort_OM              
        } 
    } 
    until($login.Length -gt 0 -or $askAttempts -gt 2) 
    if($askAttempts -gt 3) { 
        log -text "user refused to enter an SMS code, exiting" -fout
        $script:errorsForUser += "You did not enter an SMS code, script cannot continue`n"
        abort_OM 
    }else{ 
        return $login 
    } 
}

function Get-ProcessWithOwner { 
    param( 
        [parameter(mandatory=$true,position=0)]$ProcessName 
    ) 
    $ComputerName=$env:COMPUTERNAME 
    $UserName=$env:USERNAME 
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($(New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$('ProcessName','UserName','Domain','ComputerName','handle')))) 
    try { 
        $Processes = Get-wmiobject -Class Win32_Process -ComputerName $ComputerName -Filter "name LIKE '$ProcessName%'" 
    } catch { 
        return -1 
    } 
    if ($Processes -ne $null) { 
        $OwnedProcesses = @() 
        foreach ($Process in $Processes) { 
            if($Process.GetOwner().User -eq $UserName){ 
                $Process |  
                Add-Member -MemberType NoteProperty -Name 'Domain' -Value $($Process.getowner().domain) 
                $Process | 
                Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value $ComputerName  
                $Process | 
                Add-Member -MemberType NoteProperty -Name 'UserName' -Value $($Process.GetOwner().User)  
                $Process |  
                Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers 
                $OwnedProcesses += $Process 
            } 
        } 
        return $OwnedProcesses 
    } else { 
        return 0 
    } 
 
} 
#endregion

function addMapping(){
    Param(
        [String]$driveLetter,
        [String]$url,
        [String]$label
    )
    $mapping = "" | Select-Object driveLetter, URL, Label, alreadyMapped
    $mapping.driveLetter = $driveLetter
    $mapping.url = $url
    $mapping.label = $label
    $mapping.alreadyMapped = $False
    log -text "Adding to mapping list: $driveLetter ($url)"
    return $mapping
}

#this function checks if a given drivemapper is properly mapped to the given location, returns true if it is, otherwise false
function checkIfLetterIsMapped(){
    Param(
        [String]$driveLetter,
        [String]$url
    )
    if([System.IO.Directory]::Exists($driveLetter)){ 
        #check if mapped path is to at least the personal folder on Onedrive for Business, username detection would require a full login and slow things down
        #Ignore DavWWWRoot, as this does not consistently appear in the actual URL
        try{
            [string]$mapped_URL = @(Get-WMIObject -query "Select * from Win32_NetworkConnection Where LocalName = '$driveLetter'")[0].RemoteName.Replace("DavWWWRoot\","").Replace("@SSL","")
        }catch{
            log -text "problem detecting network path for $driveLetter, $($Error[0])" -fout
        }
        [String]$url = $url.Replace("DavWWWRoot\","").Replace("@SSL","")
        if($mapped_URL.StartsWith($url)){
            log -text "the mapped url for $driveLetter ($mapped_URL) matches the expected URL of $url, no need to remap"
            return $True
        }else{
            log -text "the mapped url for $driveLetter ($mapped_URL) does not match the expected partial URL of $url"
            return $False
        } 
    }else{
        log -text "$driveLetter is not yet mapped"
        return $False
    }
}

function waitForIE{
    do {sleep -m 100} until (-not ($script:ie.Busy))
}

function checkIfMFASetupIsRequired{
    try{
        $found_Tfa = (getElementById -id "tfa_setupnow_button").tagName
        #two factor was required but not yet set up
        log -text "Failed to log in at $($script:ie.LocationURL) because you have not set up two factor authentication methods while you are required to." -fout
        $script:errorsForUser += "Cannot continue: you have not yet set up two-factor authentication at portal.office.com"
        abort_OM 
    }catch{$Null}    
}

function checkIfCOMObjectIsHealthy{
    #check if the COM object is healthy, otherwise we're running into issues 
    if($script:ie.HWND -eq $null){ 
        log -text "ERROR: the browser object was Nulled during login, this means IE ProtectedMode or other security settings are blocking the script, check if you have correctly configure Trusted Sites." -fout
        $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
        abort_OM 
    } 
}

#Returns True if there was an error (and logs the error), returns False if no error was detected
function checkErrorAtLoginValue{
    Param(
        [String]$mode #msonline = microsoft, #adfs = adfs of client
    )
    if($mode -eq "msonline"){
        try{
            $found_ErrorControl = (getElementById -id "error_code").value
        }catch{$Null}
    }elseif($mode -eq "adfs"){
        try{
            $found_ErrorControl = (getElementById -id "errorText").innerHTML
        }catch{$Null}
    }
    if($found_ErrorControl){
        if($mode -eq "msonline"){
            switch($found_ErrorControl){
                "InvalidUserNameOrPassword" {
                    log -text "Detected an error at $($ie.LocationURL): invalidUsernameOrPassword" -fout
                    $script:errorsForUser += "The password or login you're trying to use is invalid`n"
                }
                default{                
                    log -text "Detected an error at $($ie.LocationURL): $found_ErrorControl" -fout
                    $script:errorsForUser += "Office365 reported an error: $found_ErrorControl`n"
                }
            }
            return $True
        }elseif($mode -eq "adfs"){
            if($found_ErrorControl.Length -gt 1){
                log -text "Detected an error at $($ie.LocationURL): $found_ErrorControl" -fout
                return $True
            }else{
                return $False
            }
        }
    }else{
        return $False
    }
}

function checkIfMFAControlsArePresent{
    Param(
        [Switch]$withoutADFS
    )
    try{
        $found_TfaWaiting = (getElementById -id "tfa_results_container").tagName
    }catch{$found_TfaWaiting = $Null}
    if($found_TfaWaiting){return $True}else{return $False}
}

#region loginFunction
function login(){ 
    log -text "Login attempt at Office 365 signin page" 
    #AzureAD SSO check if a tile exists for this user
    $skipNormalLogin = $False
    if($userLookupMode -le 3){
        try{
            $lookupQuery = $userUPN -replace "@","_"
            $lookupQuery = $lookupQuery -replace "\.","_"
            $userTile = getElementById -id $lookupQuery
            $skipNormalLogin = $True
            log -text "detected SSO option for OnedriveMapper through AzureAD, attempting to login automatically"
            $userTile.Click()
            waitForIE
            Sleep -m 500
            waitForIE
            if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
                #we've been logged in, we can abort the login function 
                log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
                return $True             
            }else{
                $skipNormalLogin = $False
            }
        }catch{
            $skipNormalLogin = $False
            log -text "failed to use Azure AD SSO for Workplace Joined devices" -fout
        }
    }
    
    if(!$skipNormalLogin){
        #click to open up the login menu 
        try{
            (getElementById -id "use_another_account").Click()
            log -text "Found sign in elements type 1 on Office 365 login page, proceeding" 
        }catch{
            log -text "Failed to find signin element type 1 on Office 365 login page, trying next method. Error details: $($Error[0])"
        }
        try{
            (getElementById -id "use_another_account_link").click() 
            log -text "Found sign in elements type 2 on Office 365 login page, proceeding" 
        }catch{
            log -text "Failed to find signin element type 2 on Office 365 login page, trying next method. Error details: $($Error[0])"
        }
        try{
            $Null = getElementById -id "cred_keep_me_signed_in_checkbox"
        }catch{
            log -text "Failed to find signin element type 3 at $($script:ie.LocationURL). You may have to upgrade to a later Powershell version or Install Office. Attempting to log in anyway, this will likely fail. Error details: $($Error[0])" -fout
        }
        waitForIE 
        if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
            #we've been logged in, we can abort the login function 
            log -text "user detected as logged in, login function succeeded but mapping will probably fail, final url: $($script:ie.LocationURL)" 
            return $True             
        }
 
        if($userLookupMode -eq 4){
            $userName = (retrieveLogin)
        }else{
            $userName = $userUPN
        }
        #attempt to trigger redirect to detect if we're using ADFS automatically 
        try{ 
            log -text "attempting to trigger a redirect to SSO Provider using method 1" 
            $checkBox = getElementById -id "cred_keep_me_signed_in_checkbox"
            if($checkBox.checked -eq $False){
                $checkBox.click() 
                log -text "Signin Option persistence selected"
            }else{
                log -text "Signin Option persistence was already selected"
            }
            if($checkBox.checked -eq $False){
                log -text "the cred_keep_me_signed_in_checkbox is not selected! This may result in error 224" -fout
            }
            (getElementById -id "cred_userid_inputtext").value = $userName       
            waitForIE 
            (getElementById -id "cred_password_inputtext").click() 
            waitForIE
        }catch{ 
            log -text "Failed to find the correct controls at $($script:ie.LocationURL) to log in by script, check your browser and proxy settings or check for an update of this script. $($Error[0])" -fout
            $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
            abort_OM  
        } 
    }
    sleep -s 2 

    #update progress bar
    if($showProgressBar) {
        $script:progressbar1.Value = 35
        $script:form1.Refresh()
    }

    $redirWaited = 0 
    while($True){ 
        sleep -m 500 

        checkIfMFASetupIsRequired

        checkIfCOMObjectIsHealthy

        #If ADFS or Azure automatically signs us on, this will trigger
        if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
            log -text "Detected an url that indicates we've been signed in automatically: $($script:ie.LocationURL)"
            $useADFS = $True
            break            
        }

        #this is the ADFS login control ID, modify this in the script setup if you have a custom IdP
        try{
            $found_ADFSControl = getElementById -id $adfsLoginInput
        }catch{
            $found_ADFSControl = $Null
            log -text "Waited $redirWaited of $adfsWaitTime seconds for SSO redirect. While looking for $adfsLoginInput at $($script:ie.LocationURL). If you're not using SSO this message is expected."
        }

        $redirWaited += 0.5 
        #found ADFS control
        if($found_ADFSControl){
            log -text "ADFS Control found, we were redirected to: $($script:ie.LocationURL)" 
            $useADFS = $True
            break            
        } 

        if($redirWaited -ge $adfsWaitTime){ 
            log -text "waited for more than $adfsWaitTime to get redirected by SSO provider, attempting normal signin" 
            $useADFS = $False    
            break 
        } 
    }  
    
    #update progress bar
    if($showProgressBar) {
        $script:progressbar1.Value = 40
        $script:form1.Refresh()
    }       

    #if not using ADFS, sign in 
    if($useADFS -eq $False){ 
        if($abortIfNoAdfs){
            log -text "abortIfNoAdfs was set to true, SSO provider was not detected, script is exiting" -fout
            $script:errorsForUser += "Onedrivemapper cannot login because SSO provider is not available`n"
            abort_OM
        }
        if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
            #we've been logged in, we can abort the login function 
            log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
            return $True             
        }
        $pwdAttempts = 0
        while($pwdAttempts -lt 3){
            $pwdAttempts++
            try{ 
                $checkBox = getElementById -id "cred_keep_me_signed_in_checkbox"
                if($checkBox.checked -eq $False){
                    $checkBox.click() 
                    log -text "Signin Option persistence selected"
                }else{
                    log -text "Signin Option persistence was already selected"
                }
                if($pwdAttempts -gt 1){
                    if($userLookupMode -eq 4){
                        $userName = (retrieveLogin -forceNewUsername)
                        (getElementById -id "cred_userid_inputtext").value = $userName
                    }
                    (getElementById -id "cred_password_inputtext").value = retrievePassword -forceNewPassword
                }else{
                    (getElementById -id "cred_password_inputtext").value = retrievePassword 
                }
                (getElementById -id "cred_sign_in_button").click() 
                waitForIE
            }catch{ 
                if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
                    #we've been logged in, we can abort the login function 
                    log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
                    return $True 
                } 
                log -text "Failed to find the correct controls at $($ie.LocationURL) to log in by script, check your browser and proxy settings or check for an update of this script (2). $($Error[0])" -fout
                abort_OM
            }
            Sleep -s 1
            waitForIE
            #check if the error field does not appear, if it does not our attempt was succesfull
            if((checkErrorAtLoginValue -mode "msonline") -eq $False){
                break   
            }
            $script:errorsForUser = $Null
        }

        #Office 365 two factor is required (SMS NOT YET SUPPORTED)
        if((checkIfMFAControlsArePresent -withoutADFS)){ 
            $waited = 0
            $maxWait = 90
            $loop = $True
            $MfaCodeAsked = $False
            while($loop){
                Sleep -s 2
                $waited+=2
                #check if on the MFA page, otherwise we're past the page already
                if((checkIfMFAControlsArePresent -withoutADFS)){ 
                    log -text "Waited for $waited seconds for user to complete Multi-Factor Authentication $found_TfaWaiting"
                }else{
                    log -text "Multi-Factor Authentication completed in $waited seconds"
                    $loop = $False
                }
                #check for SMS/App input field container
                try{
                    $found_MfaCode = getElementById -id "tfa_code_container"
                }catch{
                    $found_MfaCode = $Null
                }
                #if field is visible and we haven't asked before, ask for the text/app message code, otherwise user is probably using the phonecall method
                if($found_MfaCode -ne $Null -and $found_MfaCode.ariaHidden -ne $True -and $MfaCodeAsked -eq $False){
                    $MfaCodeAsked = $True
                    $code = askForCode
                    (getElementById -id "tfa_code_inputtext").value = $code
                    waitForIE
                    (getElementById -id "tfa_signin_button").click() 
                    waitForIE
                }
                if($waited -ge $maxWait){
                    log -text "Failed to log in at $($script:ie.LocationURL) because multi-factor authentication was not completed in time." -fout
                    $script:errorsForUser += "Cannot continue: you have not completed multi-factor authentication in the maximum alotted time"
                    abort_OM 
                }
            }

        }
    }else{ 
        #check if logged in now automatically after ADFS redirect 
        if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
            #we've been logged in, we can abort the login function 
            log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
            return $True 
        } 
    } 
 
    waitForIE
 
    #Check if we arrived at a 404, or an actual page 
    if($script:ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*") { 
        log -text "We received a 404 error after our signin attempt, retrying...." -fout
        $script:ie.navigate("https://portal.office.com")
        waitForIE
        if($script:ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*") { 
            log -text "We received a 404 error again, aborting" -fout
            $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
            abort_OM     
        }     
    } 

    #check if logged in now 
    if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
        #we've been logged in, we can abort the login function 
        log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
        return $True 
    }else{ 
        if($useADFS){ 
            log -text "ADFS did not automatically sign us on, attempting to enter credentials at $($script:ie.LocationURL)" 
            $pwdAttempts = 0
            while($pwdAttempts -lt 3){
                $pwdAttempts++
                try{ 
                    if($userLookupMode -eq 4 -and $pwdAttempts -gt 1){
                        $userName = (retrieveLogin -forceNewUsername)
                    }
                    if($adfsMode -eq 1){
                        (getElementById -id $adfsLoginInput).value = $userName
                    }else{
                        (getElementById -id $adfsLoginInput).value = ($userName.Split("@")[0])
                    }
                    if($pwdAttempts -gt 1){
                        (getElementById -id $adfsPwdInput).value = retrievePassword -forceNewPassword
                    }else{
                        (getElementById -id $adfsPwdInput).value = retrievePassword 
                    }
                    (getElementById -id $adfsButton).click() 
                    waitForIE 
                    sleep -s 1 
                    waitForIE  
                    if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
                        #we've been logged in, we can abort the login function 
                        log -text "login detected, login function succeeded, final url: $($script:ie.LocationURL)" 
                        return $True 
                    } 
                }catch{ 
                    log -text "Failed to find the correct controls at $($script:ie.LocationURL) to log in by script, check your browser and proxy settings or modify this script to match your ADFS form. Error details: $($Error[0])" -fout
                    $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
                    abort_OM 
                }
                #check if the error field does not appear, if it does not our attempt was succesfull
                if((checkErrorAtLoginValue -mode "adfs") -eq $False){
                    break   
                }
            }
 
            waitForIE   
            #check if logged in now         
            if((checkIfAtO365URL -url $ie.LocationURL -finalURLs $finalURLs)){
                #we've been logged in, we can abort the login function 
                log -text "login detected, login function succeeded, final url: $($ie.LocationURL)" 
                return $True 
            }else{ 
                log -text "We attempted to login with ADFS, but did not end up at the expected location. Detected url: $($ie.LocationURL), expected URL: $($baseURL)" -fout
                $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
                abort_OM 
            } 
        }else{ 
            log -text "We attempted to login without using ADFS, but did not end up at the expected location. Detected url: $($ie.LocationURL), expected URL: $($baseURL)" -fout
            $script:errorsForUser += "Mapping cannot continue because we could not log in to Office 365`n"
            abort_OM 
        } 
    } 
} 
#endregion

#return -1 if nothing found, or value if found
function checkRegistryKeyValue{
    Param(
        [String]$basePath,
        [String]$entryName
    )
    try{$value = (Get-ItemProperty -Path "$($basePath)\" -Name $entryName -ErrorAction Stop).$entryName
        return $value
    }catch{
        return -1
    }
}

function addSiteToIEZoneThroughRegistry{
    Param(
        [String]$siteUrl,
        [Int]$mode=2 #1=intranet, 2=trusted sites
    )
    try{
        $components = $siteUrl.Split(".")
        $count = $components.Count
        if($count -gt 3){
            $old = $components
            $components = @()
            $subDomainString = ""
            for($i=0;$i -le $count-3;$i++){
                if($i -lt $count-3){$subDomainString += "$($old[$i])."}else{$subDomainString += "$($old[$i])"}
            }
            $components += $subDomainString
            $components += $old[$count-2]
            $components += $old[$count-1]    
        }
        if($count -gt 2){
            $res = New-Item "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\$($components[1]).$($components[2])" -ErrorAction SilentlyContinue 
            $res = New-Item "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\$($components[1]).$($components[2])\$($components[0])" -ErrorAction SilentlyContinue
            $res = New-ItemProperty "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\$($components[1]).$($components[2])\$($components[0])" -Name "https" -value $mode -ErrorAction Stop
        }else{
            $res = New-Item "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\$($components[0]).$($components[1])" -ErrorAction SilentlyContinue 
            $res = New-ItemProperty "hkcu:\software\microsoft\windows\currentversion\internet settings\zonemap\domains\$($components[0]).$($components[1])" -Name "https" -value $mode -ErrorAction Stop
        }
    }catch{
        return -1
    }
    return $True
}

function checkWebClient{
    if((Get-Service -Name WebClient).Status -ne "Running"){ 
        #attempt to auto-start if user is admin
        if($isElevated){
            $res = Start-Service WebClient -ErrorAction SilentlyContinue
        }else{
            #use another trick to autostart the client
            try{
                startWebDavClient
            }catch{
                log -text "CRITICAL ERROR: OneDriveMapper detected that the WebClient service was not started, please ensure this service is always running!`n" -fout
                $script:errorsForUser += "$MD_DriveLetter could not be mapped because the WebClient service is not running`n"
            }
        }
    } 
}


#check if the script is running elevated, run via scheduled task if UAC is not disabled
If (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){   
    log -text "Script elevation level: Administrator" -fout
    $scheduleTask = $True
    $isElevated = $True
    if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -entryName "EnableLUA") -eq 0){
        log -text "NOTICE: $($BaseKeypath)\EnableLua found in registry and set to 0, you have disabled UAC, the script does not need to bypass by using a scheduled task"    
        $scheduleTask = $False                
    }    
    if($asTask){
        log -text "Already running as task, but still elevated, will attempt to map normally but drives may not be visible" -fout
        $scheduleTask = $False
    }
    checkWebClient
    if($scheduleTask){
        $Null = fixElevationVisibility
        Exit
    }
}else{
    log -text "Script elevation level: User"
    $isElevated = $False
    checkWebClient
}

#load windows libraries to display things to the user 
try{ 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
}catch{ 
    log -text "Error loading windows forms libraries, script will not be able to display a password input box" -fout
} 

#load settings from OnedriveMapper Configurator if licensed
if($configurationID -ne "00000000-0000-0000-0000-000000000000"){
    try{
        log -text "configurationID set to $configurationID, retrieving associated settings from lieben.nu..."
        $rawSettingsResponse = JosL-WebRequest -url "http://www.lieben.nu/om_api_v1.php?cid=$configuratonID"
        log -text "settings retrieved, processing..."
    }catch{
        log -text "failed to retrieve settings from lieben.nu using $configurationID because of $($Error[0])" -fout
        log -text "Will attempt to use default / in-script settings but will likely fail"
    }
}

#show a progress bar if set to True
if($showProgressBar) {
    #title for the winform
    $Title = "OnedriveMapper v$version"
    #winform dimensions
    $height=39
    $width=400
    #winform background color
    $color = "White"

    #create the form
    $form1 = New-Object System.Windows.Forms.Form
    $form1.Text = $title
    $form1.Height = $height
    $form1.Width = $width
    $form1.BackColor = $color
    $form1.ControlBox = $false
    $form1.MaximumSize = New-Object System.Drawing.Size($width,$height)
    $form1.MinimumSize = new-Object System.Drawing.Size($width,$height)
    $form1.Size = new-Object System.Drawing.Size($width,$height)

    $form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None 
    #display center screen
    $form1.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $screen = ([System.Windows.Forms.Screen]::AllScreens | where {$_.Primary}).WorkingArea
    $form1.Location = New-Object System.Drawing.Size(($screen.Right - $width), ($screen.Bottom - $height))
    $form1.Topmost = $True 
    $form1.TopLevel = $True 

    # create label
    $label1 = New-Object system.Windows.Forms.Label
    $label1.text="OnedriveMapper v$version is connecting your drives..."
    $label1.Name = "label1"
    $label1.Left=0
    $label1.Top= 9
    $label1.Width= $width
    $label1.Height=17
    $label1.Font= "Verdana"
    # create label
    $label2 = New-Object system.Windows.Forms.Label
    $label2.Name = "label2"
    $label2.Left=0
    $label2.Top= 0
    $label2.Width= $width
    $label2.Height=7
    $label2.backColor= "#CC99FF"

    #add the label to the form
    $form1.controls.add($label1) 
    $form1.controls.add($label2) 
    $progressBar1 = New-Object System.Windows.Forms.ProgressBar
    $progressBar1.Name = 'progressBar1'
    $progressBar1.Value = 0
    $progressBar1.Style="Continuous" 
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = $width
    $System_Drawing_Size.Height = 10
    $progressBar1.Size = $System_Drawing_Size   
    
    $progressBar1.Left = 0
    $progressBar1.Top = 29
    $form1.Controls.Add($progressBar1)
    $form1.Show()| out-null  
    $form1.Focus() | out-null 
    $progressbar1.Value = 5
    $form1.Refresh()
}

#do a version check if allowed
if($versionCheck){
    #update progress bar
    try{
        versionCheck -currentVersion $version
        log -text "NOTICE: you are running the latest (v$version) version of OnedriveMapper"
    }catch{
        if($showProgressBar) {
            $form1.controls["Label1"].Text = "OnedriveMapper version outdated :("
            $form1.Refresh()
            Sleep -s 1
            $form1.controls["Label1"].Text = "OnedriveMapper v$version is connecting your drives..."
            $form1.Refresh()
        }
        log -text "ERROR: $($Error[0])" -fout
    }
}

#get IE version on this machine
$ieVersion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Internet Explorer').svcVersion
if($ieVersion -eq $Null){
    $ieVersion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Internet Explorer').Version
    $ieVersion = $ieVersion.Split(".")[1]
}else{
    $ieVersion = $ieVersion.Split(".")[0]
}

#get OSVersion
$windowsVersion = ([System.Environment]::OSVersion.Version).Major

log -text "You are running on Windows $windowsVersion with IE $ieVersion"

#check if KB2846960 is installed or not
try{
    $res = get-hotfix -id kb2846960 -erroraction stop
    log -text "KB2846960 detected as installed"
}catch{
    if($ieVersion -eq 10 -and $windowsVersion -lt 10){
        log -text "KB2846960 is not installed on this machine, if you're running IE 10 on anything below Windows 10, you may not be able to map your drives until you install KB2846960" -warning
    }
}

#Check if Zone Configuration is on a per machine or per user basis, then check the zones 
$privateZoneFound = $False
$publicZoneFound = $False
$msOnlineZoneFound = $False

#check if zone enforcement is set to machine only
$reg_HKLM = checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings" -entryName "Security HKLM only"
if($reg_HKLM -eq -1){
    log -text "NOTICE: HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Security HKLM only not found in registry, your zone configuration could be set on both levels" 
}elseif($reg_HKLM -eq 1){
    log -text "NOTICE: HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Security HKLM only found in registry and set to 1, your zone configuration is set on a machine level"    
}

#check if sharepoint and onedrive are set as safe sites at the user level
if($reg_HKLM -ne 1){
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -eq 2){
        log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on user level"  
        $privateZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -eq 2){
        log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on user level (through GPO)"  
        $privateZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -eq 2){
        log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on user level"  
        $publicZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -eq 2){
        log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on user level (through GPO)" 
        $publicZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftonline.com" -entryName "https") -eq 2){
        log -text "NOTICE: *.microsoftonline.com found in IE Trusted Sites on user level"  
        $msOnlineZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftonline.com" -entryName "https") -eq 2){
        log -text "NOTICE: *.microsoftonline.com found in IE Trusted Sites on user level (through GPO)" 
        $msOnlineZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\office.com" -entryName "https") -eq 2){
        log -text "NOTICE: *.office.com found in IE Trusted Sites on user level"  
        $officeZoneFound = $True        
    }
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\office.com" -entryName "https") -eq 2){
        log -text "NOTICE: *.office.com found in IE Trusted Sites on user level (through GPO)" 
        $officeZoneFound = $True        
    }
}

#check if sharepoint and onedrive are set as safe sites at the machine level
if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -eq 2){
    log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on machine level"
    $privateZoneFound = $True 
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)$($privateSuffix)" -entryName "https") -eq 2){
    log -text "NOTICE: $($O365CustomerName)$($privateSuffix).sharepoint.com found in IE Trusted Sites on machine level (through GPO)"  
    $privateZoneFound = $True        
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -eq 2){
    log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on machine level"  
    $publicZoneFound = $True    
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\sharepoint.com\$($O365CustomerName)" -entryName "https") -eq 2){
    log -text "NOTICE: $($O365CustomerName).sharepoint.com found in IE Trusted Sites on machine level (through GPO)"  
    $publicZoneFound = $True    
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftonline.com" -entryName "https") -eq 2){
    log -text "NOTICE: *.microsoftonline.com found in IE Trusted Sites on machine level"  
    $msOnlineZoneFound = $True    
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoftonline.com" -entryName "https") -eq 2){
    log -text "NOTICE: *.microsoftonline.com found in IE Trusted Sites on machine level (through GPO)"  
    $msOnlineZoneFound = $True    
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\office.com" -entryName "https") -eq 2){
    log -text "NOTICE: *.office.com found in IE Trusted Sites on machine level"  
    $officeZoneFound = $True    
}
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\office.com" -entryName "https") -eq 2){
    log -text "NOTICE: *.office.com found in IE Trusted Sites on machine level (through GPO)"  
    $officeZoneFound = $True    
}

#log results, try to automatically add trusted sites to user trusted sites if not yet added
if($publicZoneFound -eq $False){
    log -text "Possible critical error: $($O365CustomerName).sharepoint.com not found in IE Trusted Sites on user or machine level, the script will likely fail" -fout
    if((addSiteToIEZoneThroughRegistry -siteUrl "$O365CustomerName.sharepoint.com") -eq $True){log -text "Automatically added $O365CustomerName.sharepoint.com to trusted sites for this user"}
}
if($privateZoneFound -eq $False){
    log -text "Possible critical error: $($O365CustomerName)$($privateSuffix).sharepoint.com not found in IE Trusted Sites on user or machine level, the script will likely fail" -fout
    if((addSiteToIEZoneThroughRegistry -siteUrl "$($O365CustomerName)$($privateSuffix).sharepoint.com") -eq $True){log -text "Automatically added $($O365CustomerName)$($privateSuffix).sharepoint.com to trusted sites for this user"}
}
if($msOnlineZoneFound -eq $False){
    log -text "Possible critical error: *.microsoftonline.com not found in IE Trusted Sites on user or machine level, the script will likely fail" -fout
    if((addSiteToIEZoneThroughRegistry -siteUrl "microsoftonline.com") -eq $True){log -text "Automatically added *.microsoftonline.com to trusted sites for this user"}
}
if($officeZoneFound -eq $False){
    log -text "Possible critical error: *.office.com not found in IE Trusted Sites on user or machine level, the script will likely fail" -fout
    if((addSiteToIEZoneThroughRegistry -siteUrl "office.com") -eq $True){log -text "Automatically added *.office.com to trusted sites for this user"}
}

#Check if IE FirstRun is disabled at computer level
if((checkRegistryKeyValue -basePath "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main" -entryName "DisableFirstRunCustomize") -ne 1){
    log -text "HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\DisableFirstRunCustomize not found or not set to 1 registry, if script hangs this may be due to the First Run popup in IE, checking at user level..." -warning
    #Check if IE FirstRun is disabled at user level
    if((checkRegistryKeyValue -basePath "HKCU:\Software\Microsoft\Internet Explorer\Main" -entryName "DisableFirstRunCustomize") -ne 1){
        log -text "HKCU:\Software\Microsoft\Internet Explorer\Main\DisableFirstRunCustomize not found or not set to 1 registry, if script hangs this may be due to the First Run popup in IE, attempting to autocorrect..." -warning
        try{
            $res = New-ItemProperty "hkcu:\software\microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -value 1 -ErrorAction Stop
            log -text "automatically prevented IE firstrun Popup"
        }catch{
            log -text "failed to automatically add a registry key to prevent the IE firstrun wizard from starting"
        }
    }    
}

#get user login 
if($forceUserName.Length -gt 2){ 
    log -text "A username was already specified in the script configuration: $($forceUserName)" 
    $userUPN = $forceUserName 
    $userLookupMode = 0
}else{
    switch($userLookupMode){
        1 {    
            log -text "userLookupMode is set to 1 -> checking Active Directory UPN" 
            $userUPN = (lookupUPN).ToLower()
        }
        2 {
            log -text "userLookupMode is set to 2 -> checking Active Directory email address" 
            $userUPN = (lookupUPN).ToLower()  
        }
        3 {
        #Windows 10
             try{
                $basePath = "HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\AADNGC"
                if((test-path $basePath) -eq $False){
                    log -text "userLookupMode is set to 3, but the registry path $basePath does not exist! Using method 2" -fout
                    $basePath = "HKCU:\Software\Classes\Local Settings\Software\Microsoft\SettingSyncHost.exe\WinMSIPC"
                    if((test-path $basePath) -eq $False){
                        log -text "userLookupMode is set to 3, but the registry path $basePath does not exist! Using method 3" -fout 
                        $objUser = New-Object System.Security.Principal.NTAccount($Env:USERNAME)
                        $strSID = ($objUser.Translate([System.Security.Principal.SecurityIdentifier])).Value
                        $basePath = "HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\$strSID\IdentityCache\$strSID"
                        if((test-path $basePath) -eq $False){
                            log -text "userLookupMode is set to 3, but the registry path $basePath does not exist! All lookup modes exhausted, exiting" -fout
                            abort_OM   
                        }
                        $userId = (Get-ItemProperty -Path $basePath -Name UserName).UserName
                    }else{
                        $userId = @(Get-ChildItem $basePath)[0].Name | Split-Path -Leaf
                    }
                }else{
                    $basePath = @(Get-ChildItem $basePath)[0].Name -Replace "HKEY_CURRENT_USER","HKCU:"
                    $userId = (Get-ItemProperty -Path $basePath -Name UserId).UserId
                }
                if($userId -and $userId -like "*@*"){
                    log -text "userLookupMode is set to 3, we detected $userId in $basePath"
                    $userUPN = ($userId).ToLower()
                }else{
                    log -text "userLookupMode is set to 3, but we failed to detect a username at $basePath" -fout
                    abort_OM
                }
             }catch{
                log -text "userLookupMode is set to 3, but we failed to detect a proper username" -fout
                abort_OM
             }
        }
        4 {
            #handled later on in the script
        }
        5 {
            if(([Environment]::UserName).Split("@").Count -eq 2){
                $userUPN = (([Environment]::UserName)).ToLower()
                log -text "userlookupMode is set to 4 -> Using $userUPN from the currently locally logged in username" 
            }else{
                $userUPN = (([Environment]::UserName)+"@"+$domain).ToLower()
                log -text "userlookupMode is set to 4 -> Using $userUPN from the currently logged in username + $domain" 
            }    
        }
        default {
            log -text "userLookupMode is not properly configured" -fout
            abort_OM
        }
    } 
}

#update progress bar
if($showProgressBar) {
    $progressbar1.Value = 10
    $form1.Refresh()
}

#region flightChecks


#Check if required HTML parsing libraries have been installed 
if([System.IO.File]::Exists("$(${env:ProgramFiles(x86)})\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll") -eq $False){ 
    log -text "Microsoft Office installation not detected"
}  
 
#Check if webdav locking is enabled
if((checkRegistryKeyValue -basePath "HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\" -entryName "SupportLocking") -ne 0){
    log -text "ERROR: WebDav File Locking support is enabled, this could cause files to become locked in your OneDrive or Sharepoint site" -fout 
} 

#report/warn file size limit
$sizeLimit = [Math]::Round((checkRegistryKeyValue -basePath "HKLM:\SYSTEM\CurrentControlSet\Services\WebClient\Parameters\" -entryName "FileSizeLimitInBytes")/1024/1024)
log -text "Maximum file upload size is set to $sizeLimit MB" -warning

#check if any zones are configured with Protected Mode through group policy (which we can't modify) 
$BaseKeypath = "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" 
for($i=0; $i -lt 5; $i++){ 
    $curr = Get-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500" -ErrorAction SilentlyContinue | select -exp 2500 
    if($curr -ne $Null -and $curr -ne 3){ 
        log -text "IE Zone $i protectedmode is enabled through group policy, autoprotectedmode cannot disable it. This will likely cause the script to fail." -fout
    }
} 

#endregion
 
#translate to URLs 
$mapURLpersonal = ("\\"+$O365CustomerName+"$($privateSuffix).sharepoint.com@SSL\DavWWWRoot\personal\") 
if($dontMapO4B){
    $baseURL = ("https://"+$O365CustomerName+".sharepoint.com") 
}else{
    $baseURL = ("https://"+$O365CustomerName+$privateSuffix+".sharepoint.com") 
}

$desiredMappings = @() #array with mappings to be made

#add the O4B mapping first, with an incorrect URL that will be updated later on because we haven't logged in yet and can't be sure of the URL
if($dontMapO4B){
    log -text "Not mapping O4B because dontMapO4B is set to True"
}else{
    $desiredMappings += addMapping -driveLetter $driveLetter -url $mapURLpersonal -label $driveLabel
}

$WebAssemblyloaded = $True
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Web")
if(-NOT [appdomain]::currentdomain.getassemblies() -match "System.Web"){
    log -text "Error loading System.Web library to decode sharepoint URL's, mapped sharepoint URL's may become read-only. $($Error[0])" -fout
    $WebAssemblyloaded = $False
}

#update progress bar
if($showProgressBar) {
    $progressbar1.Value = 15
    $form1.Refresh()
}

#add any desired Sharepoint Mappings
$sharepointMappings | % {
    $data = $_.Split(",")
    if($data[0] -and $data[1] -and $data[2]){
        if($WebAssemblyloaded){
            $add = [System.Web.HttpUtility]::UrlDecode($data[0])
        }else{
            $add = $data[0]
        }
        $add = $add.Replace("https://","\\") 
        $add = $add.Replace("/_layouts/15/start.aspx#","")
        $add = $add.Replace("sharepoint.com/","sharepoint.com@SSL\DavWWWRoot\") 
        $add = $add.Replace("/","\") 
        $desiredMappings += addMapping -driveLetter $data[2] -url $add -label $data[1]
    }
}

$continue = $False
$countMapping = 0
#check if any of the mappings we should make is already mapped and update the corresponding property
$desiredMappings | % {
    if((checkIfLetterIsMapped -driveLetter $_.driveletter -url $_.url)){
        $desiredMappings[$countMapping].alreadyMapped = $True
        if($redirectMyDocs -and $_.driveletter -eq $driveLetter) {
            $res = redirectMyDocuments -driveLetter $driveLetter
        }
    }
    $countMapping++
}
 
if(@($desiredMappings | where-object{$_.alreadyMapped -eq $False}).Count -le 0){
    log -text "no unmapped or incorrectly mapped drives detected"
    abort_OM    
}

#update progress bar
if($showProgressBar) {
    $progressbar1.Value = 20
    $form1.Refresh()
}

handleAzureADConnectSSO -initial

log -text "Base URL: $($baseURL) `n" 

#Start IE and stop it once to make sure IE sets default registry keys 
if($autoKillIE){ 
    #start invisible IE instance 
    $tempIE = new-object -com InternetExplorer.Application 
    $tempIE.visible = $debugmode 
    sleep 2 
 
    #kill all running IE instances of this user 
    $ieStatus = Get-ProcessWithOwner iexplore 
    if($ieStatus -eq 0){ 
        log -text "no instances of Internet Explorer running yet, at least one should be running" -warning
    }elseif($ieStatus -eq -1){ 
        log -text "Checking status of iexplore.exe: unable to query WMI" -fout
    }else{ 
        log -text "autoKillIE enabled, stopping IE processes" 
        foreach($Process in $ieStatus){ 
                Stop-Process $Process.handle -ErrorAction SilentlyContinue
                log -text "Stopped process with handle $($Process.handle)"
        } 
    } 
}else{ 
    log -text "ERROR: autoKillIE disabled, IE processes not stopped. This may cause the script to fail for users with a clean/new profile" -fout
} 

if($autoProtectedMode){ 
    log -text "autoProtectedMode is set to True, disabling ProtectedMode temporarily" 
    $BaseKeypath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" 
     
    #store old values and change new ones 
    try{ 
        for($i=0; $i -lt 5; $i++){ 
            $curr = Get-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500" -ErrorAction SilentlyContinue| select -exp 2500 
            if($curr -ne $Null){ 
                $protectedModeValues[$i] = $curr 
                log -text "Zone $i was set to $curr, setting it to 3" 
            }else{
                $protectedModeValues[$i] = 0 
                log -text "Zone $i was not yet set, setting it to 3" 
            }
            Set-ItemProperty -Path "$($BaseKeypath)\$($i)\" -Name "2500"  -Value "3" -Type Dword -ErrorAction Stop
        } 
    } 
    catch{ 
        log -text "Failed to modify registry keys to autodisable ProtectedMode $($error[0])" -fout
    } 
}else{
    log -text "autoProtectedMode is set to False, IE ProtectedMode will not be disabled temporarily" -fout
}

#start invisible IE instance 
$COMFailed = $False
try{ 
    $script:ie = new-object -com InternetExplorer.Application -ErrorAction Stop
    $script:ie.visible = $debugmode 
}catch{ 
    log -text "failed to start Internet Explorer COM Object, check user permissions or already running instances. Will retry in 30 seconds. $($Error[0])" -fout
    $COMFailed = $True
} 

#retry above if failed
if($COMFailed){
    Sleep -s 30
    try{ 
        $script:ie = new-object -com InternetExplorer.Application -ErrorAction Stop
        $script:ie.visible = $debugmode 
    }catch{ 
        log -text "failed to start Internet Explorer COM Object a second time, check user permissions or already running instances $($Error[0])" -fout
        $errorsForUser += "Mapping cannot continue because we could not start your browser"
        abort_OM 
    }
}

#update progress bar
if($showProgressBar) {
    $progressbar1.Value = 25
    $form1.Refresh()
}

#navigate to the base URL of the tenant's Sharepoint to check if it exists 
try{ 
    $script:ie.navigate("https://login.microsoftonline.com/logout.srf")
    waitForIE
    sleep -s 1
    waitForIE
    $script:ie.navigate($o365loginURL) 
    waitForIE
}catch{ 
    log -text "Failed to browse to the Office 365 Sign in page, this is a fatal error $($Error[0])`n" -fout
    $errorsForUser += "Mapping cannot continue because we could not contact Office 365`n"
    abort_OM 
} 
 
#check if we got a 404 not found 
if($script:ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*") { 
    log -text "Failed to browse to the Office 365 Sign in page, 404 error detected, exiting script" -fout
    $errorsForUser += "Mapping cannot continue because we could not start the browser`n"
    abort_OM 
} 
 
checkIfCOMObjectIsHealthy

if($script:ie.LocationURL.StartsWith($o365loginURL)){
    log -text "Starting logon process at: $($script:ie.LocationURL)" 
}else{
    log -text "For some reason we're not at the logon page, even though we tried to browse there, we'll probably fail now but let's try one final time. URL: $($script:ie.LocationURL)" -fout
    $script:ie.navigate($o365loginURL) 
}

#update progress bar
if($showProgressBar) {
    $progressbar1.Value = 30
    $form1.Refresh()
}

#log in 
if((checkIfAtO365URL -url $script:ie.LocationURL -finalURLs $finalURLs)){
    log -text "Detected an url that indicates we've been signed in automatically: $($script:ie.LocationURL), but we did not select sign in persistence, this may cause an error when mapping" -fout
}else{ 
    #Check and log if Explorer is running 
    $explorerStatus = Get-ProcessWithOwner explorer 
    if($explorerStatus -eq 0){ 
        log -text "no instances of Explorer running yet, expected at least one running" -warning
    }elseif($explorerStatus -eq -1){ 
        log -text "Checking status of explorer.exe: unable to query WMI" -fout
    }else{ 
        log -text "Detected running explorer process" 
    } 
    $res = login
    $script:ie.navigate($baseURL) 
    waitForIE
    do {sleep -m 100} until ($script:ie.ReadyState -eq 4 -or $script:ie.ReadyState -eq 0)  
    Sleep -s 2
} 

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 45
    $script:form1.Refresh()
}

$url = $script:ie.LocationURL 
$timeSpent = 0
#find username
if($dontMapO4B -eq $False){
    while($url.IndexOf("/personal/") -eq -1){
        log -text "Attempting to detect username at $url, waited for $timeSpent seconds" 
        $script:ie.navigate($baseURL)
        waitForIE
        if($timeSpent -gt 60){
            log -text "Failed to get the username from the URL for over $timeSpent seconds while at $url, aborting" -fout 
            $errorsForUser += "Mapping cannot continue because we cannot detect your username`n"
            abort_OM 
        }
        Sleep -s 2
        $timeSpent+=2
        $url = $script:ie.LocationURL
    }
    try{
        $start = $url.IndexOf("/personal/")+10 
        $end = $url.IndexOf("/",$start) 
        $userURL = $url.Substring($start,$end-$start) 
        $mapURL = $mapURLpersonal + $userURL + "\" + $libraryName 
    }catch{
        log -text "Failed to get the username while at $url, aborting" -fout
        $errorsForUser += "Mapping cannot continue because we cannot detect your username`n"
        abort_OM 
    }
    $desiredMappings[0].url = $mapURL 
    log -text "Detected user: $($userURL)"
    log -text "Onedrive cookie generated, mapping drive..."
    $mapresult = MapDrive $desiredMappings[0].driveLetter $desiredMappings[0].url $desiredMappings[0].label
} 

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 50
    $script:form1.Refresh()
    $maxAdded = 40
    $added = 0
}

log -text "Current location: $($script:ie.LocationURL)" 
foreach($spMapping in $sharepointMappings){
    $data = $spMapping.Split(",")
    $desiredMapping = $Null
    [Array]$desiredMapping = @($desiredMappings | where{$_.alreadyMapped -eq $False -and $_.driveLetter -eq $data[2] -and $_})
    if($desiredMapping.Count -ne 1){
        continue
    }
    log -text "browsing to Sharepoint site to validate existence and set a cookie: $($data[0])"
    if($data[0] -and $data[1] -and $data[2]){
        $spURL = $data[0] #URL to browse to
        $script:ie.navigate($spURL) #check the URL
        $waited = 0
        waitForIE
        while($($ie.LocationURL) -notlike "$spURL*"){
            sleep -s 1
            $waited++
            log -text "waited $waited seconds to load $spURL, currently at $($ie.LocationURL)"
            if($waited -ge $maxWaitSecondsForSpO){
                log -text "waited longer than $maxWaitSecondsForSpO seconds to load $spURL! This mapping may fail" -fout
                break
            }
        }
        if($script:ie.Document.IHTMLDocument2_url -like "res://ieframe.dll/http_404.htm*" -or $script:ie.HWND -eq $null) { 
            log -text "Failed to browse to Sharepoint URL $spURL.`n" -fout
        } 
        log -text "Current location: $($script:ie.LocationURL)" 
    }
    #update progress bar
    if($showProgressBar) {
        if($added -le $maxAdded){
            $script:progressbar1.Value += 10
            $script:form1.Refresh()
        }
        $added+=10
    }
    log -text "SpO cookie generated, attempting to map drive"
    $mapresult = MapDrive $desiredMapping[0].driveLetter $desiredMapping[0].url $desiredMapping[0].label
}

#update progress bar
if($showProgressBar) {
    $script:progressbar1.Value = 100
    $script:form1.Refresh()
}

abort_OM