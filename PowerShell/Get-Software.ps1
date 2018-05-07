Function Get-Software {
<#
.Synopsis
   Function to get the installed Software from a local or remote Windows System.

.DESCRIPTION
   Function to get the installed Software from a local or remote Windows System.
   This function offers Objects with informations similar to the Installed Software Window-view in the Windows Systemmanagemend.
   This function will return information about each windows installer-based application out of the Registry 'Uninstall' key (32 and 64 Bit).
   In order for a registry key to be opened remotely, both the server and client machines must be running the remote registry service, and have remote administration enabled.
   If the remote registry service is not accessible then you can use this function with PowerShell remoting.

   ADVICE: The WMI/CIM Win32_Product class is considered evil!

       It is not recomende to use WMi/CIM and the class Win32_Product for this task!
       Unfortunately, Win32_Product uses a provider DLL that validates the consistency of every installed MSI package on the computer.
       If a Software installation is inconsistent the Sofware starts to try to repair themself with a reinstallation!
       That makes it very, very slow and in case of a inconsistent Software the task will hang!
       
       If you use the Computername Parameter of this Funktion please registry service of one Computer connect to (many) other Computers even If you use PowerShell Jobs with this Function.
       This is very slow!
       If you like to query many Computers use this Funktion with PowerShell remoting so that it is executed on the target systems!
       See Invoke-Command Example in section Examples

.PARAMETER Computername
    Specifies the target computer for the operation.
    The local machine registry is opened if Computername is an empty String "".
    Default is an empty String "".

.PARAMETER IncludeEmptyDisplaynames
     Returns even Software Objects with empty Displaynames
     Default is to NOT return Software Objects with empty Displaynames

.EXAMPLE
   Get-Software
   
   Retrieves the Installed Software from the locale System

.EXAMPLE
   Get-Software -IncludeEmptyDisplaynames
   
   Retrieves the Installed Software from the locale System include Software with empty displaynames

.EXAMPLE
   Get-Software -ComputerName 'Server1'

   Retrieves the Installed Software from the remote System 'Server1' out of the registry
   In order for a registry key to be opened remotely, both the server and client machines must be running the remote registry service, and have remote administration enabled.

.EXAMPLE
    Invoke-Command -ComputerName 'Server1','Server2','Server3' -ScriptBlock ${Function:Get-Software}

.INPUTS
   You can pipe the ComputerName(s) as input to Get-InstalledSoftware

.OUTPUTS
    PSObject with following NoteProperties:
        ComputerName
        AuthorizedCDFPrefix
        Comments
        Contact
        DisplayVersion
        HelpLink
        HelpTelephone
        InstallDate
        InstallLocation
        InstallSource
        ModifyPath
        Publisher
        Readme
        Size
        EstimatedSize
        UninstallString
        URLInfoAbout
        URLUpdateInfo
        VersionMajor
        VersionMinor
        WindowsInstalle
        Version
        Language
        DisplayName

.NOTES
   Author: Peter Kriegel
   Version: 2.0.2.
   13.January.2014
   HTTP://www.Admin-Source.de
#>
 
    [Cmdletbinding()]
    Param(
        [Parameter(Position=0,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,ValueFromRemainingArguments=$False)]
        [ValidateNotNull()]
        [String[]]$Computername = @(''),
        [Switch]$IncludeEmptyDisplaynames
    )
 
     Begin {
 
         # sub function to convert the registry values to an Object
         Function Convert-RegistryUninstallSubKeyToObject {
            param (
                [String]$Computername,
                [microsoft.win32.registrykey]$SubKey
            )

            # create New Object with empty Properties
            $obj = New-Object PSObject | Select-Object ComputerName,AuthorizedCDFPrefix,Comments,Contact,DisplayVersion,HelpLink,HelpTelephone,InstallDate,InstallLocation,InstallSource,ModifyPath,Publisher,Readme,Size,EstimatedSize,UninstallString,URLInfoAbout,URLUpdateInfo,VersionMajor,VersionMinor,WindowsInstaller,Version,Language,DisplayName
 
            $obj.ComputerName = $Computername
            $obj.AuthorizedCDFPrefix = $SubKey.GetValue('AuthorizedCDFPrefix')
            $obj.Comments = $SubKey.GetValue('Comments')
            $obj.Contact = $SubKey.GetValue('Contact')
            $obj.DisplayVersion = $SubKey.GetValue('DisplayVersion')
            $obj.HelpLink = $SubKey.GetValue('HelpLink')
            $obj.HelpTelephone = $SubKey.GetValue('HelpTelephone')
            $obj.InstallDate = $SubKey.GetValue('InstallDate')
            $obj.InstallLocation = $SubKey.GetValue('InstallLocation')
            $obj.InstallSource = $SubKey.GetValue('InstallSource')
            $obj.ModifyPath = $SubKey.GetValue('ModifyPath')
            $obj.Publisher = $SubKey.GetValue('Publisher')
            $obj.Readme = $SubKey.GetValue('Readme')
            $obj.Size = $SubKey.GetValue('Size')
            $obj.EstimatedSize = $SubKey.GetValue('EstimatedSize')
            $obj.UninstallString = $SubKey.GetValue('UninstallString')
            $obj.URLInfoAbout = $SubKey.GetValue('URLInfoAbout')
            $obj.URLUpdateInfo = $SubKey.GetValue('URLUpdateInfo')
            $obj.VersionMajor = $SubKey.GetValue('VersionMajor')
            $obj.VersionMinor = $SubKey.GetValue('VersionMinor')
            $obj.WindowsInstaller = $SubKey.GetValue('WindowsInstaller')
            $obj.Version = $SubKey.GetValue('Version')
            $obj.Language = $SubKey.GetValue('Language')
            $obj.DisplayName = $SubKey.GetValue('DisplayName')
            
            # return Object
            $obj
        }

    } # end Begin block       

    Process {    
        foreach($pc in $Computername){
 
            $UninstallPathes = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
            
            ForEach($UninstallKey in $UninstallPathes) {
                #Create an instance of the Registry Object and open the HKLM base key
                Try {
                    $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$pc)
                } Catch {
                    $_
                    Continue
                }
 
                #Drill down into the Uninstall key using the OpenSubKey Method
                $regkey=$reg.OpenSubKey($UninstallKey)

                If(-not $regkey) {
                    Write-Error "Subkey not found in registry: HKLM:\\$UninstallKey `non Machine: $pc"
                }

                #Retrieve an array of string that contain all the subkey names
                $subkeys=$regkey.GetSubKeyNames()
    
                #Open each Subkey and use GetValue Method to return the required values for each
                foreach($key in $subkeys){
 
                    $thisKey=$UninstallKey+"\\"+$key
 
                    $thisSubKey=$reg.OpenSubKey($thisKey)
 
                    # prevent Objects with empty DisplayName
                    if (-not $thisSubKey.getValue("DisplayName") -and (-not $IncludeEmptyDisplaynames)) { continue }
 
                    # convert registry values to an Object
                    Convert-RegistryUninstallSubKeyToObject -Computername $PC -SubKey $thisSubKey
 
                }  # End ForEach $key
 
                $reg.Close()
                 
            } # End ForEach $UninstallKey
        } # End ForEach $pc
    } # end Process block

    End {
    
    } # end End block
 
} #end Function