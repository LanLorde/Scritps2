###Import Necessary Modules
cls
ImportSystemModules
Import-Module ActiveDirectory

function Create-Account
{
###Manipulate the Data
$name = $firstname + " " + $lastname
If (($firstname).Length -lt 3)
  {$SamAccountName = $firstname + $lastname}
Else
  {$SamAccountName = $firstname.Substring(0,3) + $lastname}
$SchoolUPN = $school.Substring(0,1) + "-Students"
$OUPath = "OU=$year,OU=$SchoolUPN,OU=$school,DC=msd60,DC=maercker,DC=org"
IF ([adsi]::Exists("LDAP://$OUPath") -like "*False*")
{
  Write-Host "The year $year at school $school does not have an OU associated with it." -ForegroundColor Red
  Write-Host "OU searched was $OUPath. Exiting" -ForegroundColor Red
  Break
}
$UserExist = Get-ADUser -filter {sAMAccountName -eq $SamAccountName}
If ($UserExist -ne $NULL)
{
  cls
  Write-Host "Account $SamAccountName already exists. Exiting."
  Break
}

$UPNAccount = $SamAccountName + "@maercker.org"
$SamAccountName = $SamAccountName.ToLower()
$UPNAccount = $UPNAccount.ToLower()

###Create the Account
    New-ADUser -SamAccountName $SamAccountName -UserPrincipalName $UPNAccount -Country "US" -DisplayName $name -Enabled $True -GivenName $firstname -Surname $lastname -Name $name -Path $OUPath -AccountPassword (ConvertTo-SecureString -AsPlainText "Abcd1234!" -Force) -EmailAddress $UPNAccount -OtherAttributes @{'Pager'=$StudentID} -Confirm:$False
    Add-ADGroupMember ("$" + "$year") $SamAccountName -Confirm:$False
    Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$StudentID" -Force)
}

###Input Information
cls
$MassImport = Read-Host 'Will this be a mass import? Y or N?'
If ($MassImport -eq 'Y')
{
  Do {$ImportFile = Read-Host 'What is the exact location of the file? Ex: C:\Data\students.csv?'} until ((Test-Path $ImportFile) -Like "*True*")
      $csv_info = Import-Csv -Path $ImportFile
      $csv_info | foreach {$school = $_.SchoolName
        $year = $_.YearGraduate
        $firstname = $_.FirstName
        $lastname = $_.LastName
        $StudentID = $_.IDNumber
        Create-Account
      }
}
Else
{
  Do {$school = Read-Host 'What School?  Maercker, Westview, or Holmes?'} until ($school -eq 'Maercker' -or $school -eq 'Westview' -or $school -eq 'Holmes')
  Do {$year = Read-Host 'What year (Type four digits) is the student graduating?'} until (!([string]::IsNullOrEmpty($year)))
  $firstname = Read-Host 'First Name?'
  $lastname = Read-Host 'Last Name?'
  $StudentID = Read-host 'Student ID?'
  Create-Account
}
