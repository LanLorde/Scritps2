function Get-VssWriters {
<# 
 .Synopsis
  Function to get information about VSS Writers on one or more computers

 .Description
  Function will parse information from VSSAdmin tool and return object containing
  WriterName, StateID, StateDesc, and LastError
  Function will display a progress bar while it retrives information from different
  computers.

 .Parameter ComputerName
  This is the name (not IP address) of the computer. 
  If absent, localhost is assumed.

 .Example
  Get-VssWriters
  This example will return a list of VSS Writers on localhost

 .Example
  # Get VSS Writers on localhost, sort list by WriterName
  $VssWriters = Get-VssWriters | Sort "WriterName" 
  $VssWriters | FT -AutoSize # Displays it on screen
  $VssWriters | Out-GridView # Displays it in GridView
  $VssWriters | Export-CSV ".\myReport.csv" -NoTypeInformation # Exports it to CSV

 .Example
  # Get VSS Writers on the list of $Computers, sort list by ComputerName
  $Computers = "xHost11","notThere","xHost12"
  $VssWriters = Get-VssWriters -ComputerName $Computers -Verbose | Sort "ComputerName" 
  $VssWriters | Out-GridView # Displays it in GridView
  $VssWriters | Export-CSV ".\myReport.csv" -NoTypeInformation # Exports it to CSV

 .Example
  # Reports any errors on VSS Writers on the computers listed in MyComputerList.txt, sorts list by ComputerName
  $Computers = Get-Content ".\MyComputerList.txt"
  $VssWriters = Get-VssWriters $Computers -Verbose | 
    Where { $_.StateDesc -ne 'Stable' } | Sort "ComputerName" 
  $VssWriters | Out-GridView # Displays it in GridView
  $VssWriters | Export-CSV ".\myReport.csv" -NoTypeInformation # Exports it to CSV 
 
 .Example
  # Get VSS Writers on all computers in current AD domain, sort list by ComputerName
  $Computers = (Get-ADComputer -Filter *).Name
  $VssWriters = Get-VssWriters $Computers -Verbose | Sort "ComputerName" 
  $VssWriters | Out-GridView # Displays it in GridView
  $VssWriters | Export-CSV ".\myReport.csv" -NoTypeInformation # Exports it to CSV

 .Example
  # Get VSS Writers on all Hyper-V hosts in current AD domain, sort list by ComputerName
  $FilteredComputerList = $null
  $Computers = (Get-ADComputer -Filter *).Name 
  Foreach ($Computer in $Computers) {
      if (Get-WindowsFeature -ComputerName $Computer -ErrorAction SilentlyContinue | 
        where { $_.Name -eq "Hyper-V" -and $_.InstallState -eq "Installed"}) {
          $FilteredComputerList += $Computer
      }
  }
  $VssWriters = Get-VssWriters $FilteredComputerList -Verbose | Sort "ComputerName" 
  $VssWriters | Out-GridView # Displays it in GridView
  $VssWriters | Export-CSV ".\myReport.csv" -NoTypeInformation # Exports it to CSV

 .OUTPUT
  Scripts returns a PS Object with the following properties:
    ComputerName                                                                                                                                                                           
    WriterName  
    StateID                                                                                                                                                                                
    StateDesc                                                                                                                                                                              
    LastError                                                                                                                                                                              

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 09/17/2014

#>

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$false,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [ValidateNotNullorEmpty()]
            [String[]]$ComputerName = $env:COMPUTERNAME
    )
    
    $Writers = @()
    $k = 0
    foreach ($Computer in $ComputerName) {
        try {
            Write-Verbose "Getting VssWriter information from computer $Computer"
            $k++
            $Progress = "{0:N0}" -f ($k*100/$ComputerName.count)
            Write-Progress -Activity "Processing computer $Computer ... $k out of $($ComputerName.count) computers" `
                -PercentComplete $Progress -Status "Please wait" -CurrentOperation "$Progress% complete"

            $RawWriters = Invoke-Command -ComputerName $Computer -ErrorAction Stop -ScriptBlock { 
                return (VssAdmin List Writers)
            } 

            for ($i=0; $i -lt ($RawWriters.Count-3)/6; $i++) {
                $Writer = New-Object -TypeName psobject
                $Writer| Add-Member "ComputerName" $Computer
                $Writer| Add-Member "WriterName" $RawWriters[($i*6)+3].Split("'")[1]
                $Writer| Add-Member "StateID" $RawWriters[($i*6)+6].SubString(11,1)
                $Writer| Add-Member "StateDesc" $RawWriters[($i*6)+6].SubString(14,$RawWriters[($i*6)+6].Length - 14)
                $Writer| Add-Member "LastError" $RawWriters[($i*6)+7].SubString(15,$RawWriters[($i*6)+7].Length - 15)
                $Writers += $Writer 
            }

            Write-Debug "Done"
        } catch {
            Write-Warning "Computer $Computer is offline, does not exist, or cannot be contacted"
        }
    }
    return $Writers
}