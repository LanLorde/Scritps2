#Requires -Version 4
#Requires -Modules DnsClient


function Log {
<# 
 .Synopsis
  Function to log input string to file and display it to screen

 .Description
  Function to log input string to file and display it to screen. Log entries in the log file are time stamped. Function allows for displaying text to screen in different colors.

 .Parameter String
  The string to be displayed to the screen and saved to the log file

 .Parameter Color
  The color in which to display the input string on the screen
  Default is White
  Valid options are
    Black
    Blue
    Cyan
    DarkBlue
    DarkCyan
    DarkGray
    DarkGreen
    DarkMagenta
    DarkRed
    DarkYellow
    Gray
    Green
    Magenta
    Red
    White
    Yellow

 .Parameter LogFile
  Path to the file where the input string should be saved.
  Example: c:\log.txt
  If absent, the input string will be displayed to the screen only and not saved to log file

 .Example
  Log -String "Hello World" -Color Yellow -LogFile c:\log.txt
  This example displays the "Hello World" string to the console in yellow, and adds it as a new line to the file c:\log.txt
  If c:\log.txt does not exist it will be created.
  Log entries in the log file are time stamped. Sample output:
    2014.08.06 06:52:17 AM: Hello World

 .Example
  Log "$((Get-Location).Path)" Cyan
  This example displays current path in Cyan, and does not log the displayed text to log file.

 .Example 
  "$((Get-Process | select -First 1).name) process ID is $((Get-Process | select -First 1).id)" | log -color DarkYellow
  Sample output of this example:
    "MDM process ID is 4492" in dark yellow

 .Example
  log "Found",(Get-ChildItem -Path .\ -File).Count,"files in folder",(Get-Item .\).FullName Green,Yellow,Green,Cyan .\mylog.txt
  Sample output will look like:
    Found 520 files in folder D:\Sandbox - and will have the listed foreground colors

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/06/2014
  v1.1 - 12/01/2014 - added multi-color display in the same line

#>

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [String[]]$String, 
        [Parameter(Mandatory=$false,
                   Position=1)]
            [ValidateSet("Black","Blue","Cyan","DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","Red","White","Yellow")]
            [String[]]$Color = "Green", 
        [Parameter(Mandatory=$false,
                   Position=2)]
            [String]$LogFile,
        [Parameter(Mandatory=$false,
                   Position=3)]
            [Switch]$NoNewLine
    )

    if ($String.Count -gt 1) {
        $i=0
        foreach ($item in $String) {
            if ($Color[$i]) { $col = $Color[$i] } else { $col = "White" }
            Write-Host "$item " -ForegroundColor $col -NoNewline
            $i++
        }
        if (-not ($NoNewLine)) { Write-Host " " }
    } else { 
        if ($NoNewLine) { Write-Host $String -ForegroundColor $Color[0] -NoNewline }
            else { Write-Host $String -ForegroundColor $Color[0] }
    }

    if ($LogFile.Length -gt 2) {
        "$(Get-Date -format "yyyy.MM.dd hh:mm:ss tt"): $($String -join " ")" | Out-File -Filepath $Logfile -Append 
    } else {
        Write-Verbose "Log: Missing -LogFile parameter. Will not save input string to log file.."
    }
}


function Check-EmailFormat {
    [CmdletBinding()] 
    Param([Parameter(Mandatory=$true, Position=0)][String]$EmailAddress)
    
    $Result = $null   
    if (-not $EmailAddress.Contains('.')) { # Must have one or more '.'
        $Result = "Bad email address format '$EmailAddress' - reason missing '.'" 
    } 
    $DaCount = (0..($EmailAddress.length - 1) | where { $EmailAddress[$_] -eq '@'}).Count 
    if ($DaCount -ne 1) { # Must have one '@' only
        $Result = "Bad email address format '$EmailAddress' - reason '@' appeared '$DaCount' times" 
    }
    $Result
}


function Send-Email {
<# 
 .SYNOPSIS
  Function to send email directly to recipient email server

 .DESCRIPTION
  Function is intended for bulk email sending. Attachments and many standard SMTP features are not implemented.
  Function sends email directly acting as an SMTP (sending) server. Not intended to act as SMTP relay.

 .PARAMETER From
  Sender's email address

 .PARAMETER SenderName
  Optional sender's full name

 .PARAMETER To
  Recipient email address

 .PARAMETER Subject
  Email title or subject

 .PARAMETER Body
  Email body/text. HTML is OK. 

 .PARAMETER MyFQDN
  The fully qualified domain name of this machine that's sending the email

 .PARAMETER SMTPLog
  Optional path to file where the script will log its screen output.
  Logging can build up large log files. It's not recommend as a common practice. 

 .PARAMETER ShowSMTP
  This switch is set to $false by default. When set to $true it will display SMTP commands and responses. 

 .PARAMETER TimeOut
  Duration in seconds that the script uses to wait on a recipient email server to respond.

 .PARAMETER DKIMDomain
  Internet domain name used for DKIM message signing - see http://www.dkim.org/ for more details on DKIM
  
 .PARAMETER DKIMSelector
  see http://www.dkim.org/ for more details on DKIM
  
 .PARAMETER BouncyCastle
  Path to BouncyCastle Crypto DLL, for example 'C:\Sandbox\BouncyCastle.Crypto.dll'
  This script uses the free BouncyCastle Crypto DLL for .NET that can be downloaded from the page http://www.bouncycastle.org/csharp/
  Direct download link is http://www.bouncycastle.org/csharp/download/bccrypto-net-1.7-bin.zip
 
 .PARAMETER pemFileName  
  This is the file the contains your private RSA key for the domain name provided in the DKIMDomain parameter
  For example: 'C:\Sandbox\OurCampaignOnline.com.DKIM'
  Example file content:
    -----BEGIN RSA PRIVATE KEY-----
    MIICXgIBAAKBgQDOiu1LRC5FOS2BBs8XpP3VSAKpTwyjkDOAQoNjebgNUsEo05sy
    yrc9ZbWWdHd7hMlWA5RKGvwMzHXs5BesuhfeGkd272fybqlzyT+iOZprryF6iJkj
    LF8hvGRO7cLukuRt9ggVZkEu2+Lj+lMD1Ep2355d0UztoiopTIbn12rCmQIDAQAB
    AoGAXP1LbLGbq2rcw9SO9HRCG/45xIRkildn+Hz5rpWkecsiUAFFRI7kBO5/3Oc+
    yrc9ZbWWdHd7hMlWA5RKGvwMzHXs5BesuhfeGkd272fybqlzyT+iOZprryF6iJkj
    LF8hvGRO7cLukuRt9ggVZkEu2+Lj+lMD1Ep2355d0UztoiopTIbn12rCmQIDAQAB
    fhoH4jvepDullBH6OOAnIttXjKKg2syIrFYuJ5xy0tdzgIyuuMR3WfUDLztlX38n
    SmfEs2xhAkEAzxR02F4xwR1GUKII5Gz7FQyny1eAYa8RUz6Qu52LVUGDdFWmaPVv
    yrc9ZbWWdHd7hMlWA5RKGvwMzHXs5BesuhfeGkd272fybqlzyT+iOZprryF6iJkj
    LF8hvGRO7cLukuRt9ggVZkEu2+Lj+lMD1Ep2355d0UztoiopTIbn12rCmQIDAQAB
    Ud+DzcB2RuMSKnYmqeZA/aMGYQyZXMXbARQkUgRLGj8Si0IPzeNyp1eZ8Z5m8Yro
    vmuTU5EuuUz4puPIIRbpAkEAwj87SpnkCGg8W1SAeGpi0PqqXOoNM9RswLsElnr/
    3/kSmyXzc+aUIAlW/DdDmGviYmH5sVZrO3otspQqb6bXmA==
    -----END RSA PRIVATE KEY-----

 .EXAMPLE
  Send-Email -From samb@mydomain.com -To samb@yourdomain.com -Subject 'Test Message' -Body 'this is the email body text' -MyFQDN 'mail.thismachinedomain.com'

 .EXAMPLE
  $Params = @{
    From        = 'samb@senderdomain.com' 
    SenderName  = 'Sam Boutros'
    To          = 'samb@recipientdomain.com'
    Subject     = 'IOPS test 3'
    MyFQDN      = 'mail.mememe.com'
    Body        = @"
<div align=center>
<font color=blue><strong>Here's the IOPS test result:</strong></font>
<table width=700 border=0>
<tr>
<td align=center>
<div align=center>
<img src='https://superwidgets.files.wordpress.com/2015/01/veeam-azure02.jpg' width='700' border='0' >
</div>
</td>
</tr>
</table>
</div>
"@
  }; Send-Email @Params
  This is an example of sending embedded image

 .EXAMPLE
  $Params = @{
    From        = 'samb@smartwebapps.com'  
    SenderName  = 'Sam Boutros'
    To          = 'samb@townsware.com'
    Subject     = 'Thank you notes'
    MyFQDN      = 'mail.smartwebapps.com'
    SMTPLog     = '.\log1.txt'
    Body        = @"
<div align=center>
<font color=red>
<strong> 
<a href = 'https://www.youtube.com/watch?v=VdXPVPSgEMc'> 
<img src='https://superwidgets.files.wordpress.com/2015/01/thankyounotes.jpg' alt='Thank you notes' border=0>
</a> 
</strong>
</font>
</div>
"@
    }; Send-Email @Params
  This is an example of sending an embedded video

 .EXAMPLE
    foreach ($Recipient in 'sam2@@townsware.com','samb@townsware.com') {    
        $Params = @{
            From        = 'samb@smartwebapps.com' 
            SenderName  = 'Sam Boutros'
            To          = $Recipient
            Subject     = 'Thank you notes 2'
            MyFQDN      = 'mail.smartwebapps.com'
            SMTPLog     = '.\log1.txt'
            Body        = @"
<div align=center>
<font color=red>
<strong> 
<a href = 'https://www.youtube.com/watch?v=VdXPVPSgEMc'> 
<img src='https://superwidgets.files.wordpress.com/2015/01/thankyounotes.jpg' alt='Thank you notes 2' border=0>
</a> 
</strong>
</font>
</div>
"@
        }
        try { Send-Email @Params -ErrorAction Stop } catch { "Skipping '$Recipient' - $_" }
    }
  This example sends the email to the 2 recipients listed on the first line. The try statement is important to allow processing good emails while skipping bad ones.

 .EXAMPLE
    foreach ($Recipient in (Import-Csv '.\emailList.csv').'EmailAddress') {    
        $Params = @{
            From        = 'samb@smartwebapps.com' 
            SenderName  = 'Sam Boutros'
            To          = $Recipient
            Subject     = 'Thank you notes'
            MyFQDN      = 'mail.smartwebapps.com'
            SMTPLog     = '.\log1.txt'
            Body        = @"
<div align=center>
<font color=red>
<strong> 
<a href = 'https://www.youtube.com/watch?v=VdXPVPSgEMc'> 
<img src='https://superwidgets.files.wordpress.com/2015/01/thankyounotes.jpg' alt='Thank you notes' border=0>
</a> 
</strong>
</font>
</div>
"@
        }
        try { Send-Email @Params -ErrorAction Stop } catch { "Skipping '$Recipient' - $_" }
    }
  This example sends the email to all recipients listed in the '.\emailList.csv' file.
  Assuming the '.\emailList.csv' file has a column labeled 'EmailAddress'

 .OUTPUTS
  The script returns an object with the following properties:
    RecipientEmail  : example samb@townsware.com
    RecipientServer : example aspmx3.googlemail.com
    SenderEmail     : example samb@smartwebapps.com
    SenderServer    : example mail.smartwebapps.com
    ReplyCode       : example 221
    StatusCode      : example 2.0.0
    ReplyText       : example 221 2.0.0 closing connection k3si7199658wjf.70 - gsmtp

 .LINK
  https://superwidgets.wordpress.com/category/powershell/

 .NOTES
  Function by Sam Boutros
  v1.0 - 1/27/2015
  v1.1 - 2/05/2015 - Added DKIM message signing. DomainKeys Identified Mail (DKIM) lets an organization take responsibility for an email message that is in transit.  
  DKIM provides a method for validating a domain name identity that is associated with a message through cryptographic authentication.
  see http://www.dkim.org/ for more details on DKIM
  Special thanks to Dave Wyatt for providing guidance with DKIM implementation http://powershell.org/wp/forums/topic/dkim-signing/

#>
    [CmdletBinding()] 
    Param(
        [Parameter(Mandatory=$true,  Position=0)] [String]$From,
        [Parameter(Mandatory=$false, Position=1)] [String]$SenderName,
        [Parameter(Mandatory=$true,  Position=2)] [String]$To,
        [Parameter(Mandatory=$true,  Position=3)] [String]$Subject,
        [Parameter(Mandatory=$true,  Position=4)] [String]$Body,
        [Parameter(Mandatory=$true,  Position=5)] [String]$MyFQDN,
        [Parameter(Mandatory=$false, Position=6)] [String]$SMTPLog,
        [Parameter(Mandatory=$false, Position=7)] [Switch]$ShowSMTP = $false,
        [Parameter(Mandatory=$false, Position=8)] [ValidateRange(1,60)][Int]$Timeout = 5, # Seconds
        [Parameter(Mandatory=$false, Position=9)] [String]$DKIMDomain,
        [Parameter(Mandatory=$false, Position=10)][String]$DKIMSelector,
        [Parameter(Mandatory=$false, Position=11)][String]$BouncyCastle, # 'C:\Sandbox\BouncyCastle.Crypto.dll'
        [Parameter(Mandatory=$false, Position=12)][String]$pemFileName   # 'C:\Sandbox\OurCampaignOnline.com.DKIM'

    )

    Begin {
        $Script:ReturnedStatus = @()
        $Script:Go = $true
        if ($Result = Check-EmailFormat $From) { throw $Result }
        if ($Result = Check-EmailFormat $To)   { throw $Result } 
        try { $MX = Resolve-DnsName $To.Split('@')[1] -Type MX -EA 1 } catch { throw $_ }
        if (-not $MX.NameExchange) { throw "No Mail servers found for recipient domain '$($To.Split('@')[1])'" }
        [array]$MX = $MX.NameExchange
        if ($ShowSMTP) { log "Checking recipient email server(s)" -LogFile $SMTPLog } 
        for ($i=0; $i -lt $MX.Count; $i++) {
            try { $Socket = New-Object System.Net.Sockets.TcpClient($MX[$i],25) -EA 1; break } catch {}            
        }
        if (-not $Socket) { throw "All mail servers '$($MX -join ', ')' for recipient domain are down" }
        if ($ShowSMTP) { log "Sending to email server '$($MX[$i])'" -LogFile $SMTPLog }
        if (-not $From.StartsWith('<')) { $From = '<' + $From }
        if (-not $From.EndsWith('>'))   { $From = $From + '>' }
        if (-not $To.StartsWith('<'))   { $To = '<' + $To }
        if (-not $To.EndsWith('>'))     { $To = $To + '>' }

        # DKIM related verifications/error checking
        $DKIM = $true
        if ($DKIMDomain) {
            if (-not $DKIMSelector) { $DKIMError += "Missing DKIMSelector parameter `r`n"; $DKIM =$false }
            if (-not (Test-Path $BouncyCastle)) { 
                $DKIMError += "BouncyCastle DLL '$BouncyCastle' does not exist `r`n"
                $DKIM =$false 
            } else {
                $BouncyCastle = (Get-Item -Path $BouncyCastle).FullName
            }
            try { 
                [Reflection.Assembly]::LoadFile($BouncyCastle) | Out-Null
            } catch {
                $DKIMError += "Failed to load BouncyCastle assembly '$BouncyCastle', make sure this is the BouncyCastle DLL and it's not corrupt `r`n"
                $DKIM =$false 
            }
            if (-not (Test-Path $pemFileName)) { $DKIMError += "DKIM private key file '$pemFileName' does not exist `r`n"; $DKIM =$false }
        } else {
            $DKIM = $false
        }
        if (-not $DKIM) { $DKIMError += "Not adding DKIM header.. `r`n"; log $DKIMError Yellow $SMTPLog }
    }

    Process {
        $Stream   = $Socket.GetStream() 
        $Writer   = New-Object System.IO.StreamWriter($Stream) 
        $Buffer   = New-Object System.Byte[] 1024 
        $Encoding = New-Object System.Text.ASCIIEncoding 

        # Update ReturnedStatus
        $UpdateReturnedStatus = {
            $Props = [ordered]@{
                RecipientEmail  = $To.Substring(1,$To.Length-2) 
                RecipientServer = $MX[$i]
                SenderEmail     = $From.Substring(1,$From.Length-2)
                SenderServer    = $MyFQDN
                DKIM            = $DKIM
                ReplyCode       = $(if ($Response.Length -gt 3) { $Response.Substring(0,3) })
                StatusCode      = $(if ($Response.Length -gt 9) { $Response.Substring(4,5) })
                ReplyText       = $Response.Substring(0,$Response.Length-2)
            }
            $Script:ReturnedStatus += New-Object -TypeName psobject -Property $Props
        }

        # Script block to send SMTP command
        $InvokeSMTPCommand = {
            if ($ShowSMTP) { log "Command:  $Command" -LogFile $SMTPLog }
            $Writer.WriteLine($Command) 
            $Writer.Flush() 
            $SendingTime = 0    
            do {
                Start-Sleep -Milliseconds 100
                $SendingTime += 0.1 # seconds
                while($Stream.DataAvailable) { $Read = $Stream.Read($Buffer, 0, 1024) }
                $Response = $Encoding.GetString($Buffer, 0, $Read)
                if ($SendingTime -gt $Timeout) { log 'Timed out..' Yellow $LogFile; break }
            } while (-not $Response)
            if ($Response) {
                switch ($Response.Substring(0,1)) { 
                    4 { # Transient error
                        log "Response: $Response" Yellow $SMTPLog
                        & $UpdateReturnedStatus
                        $Script:Go = $false
                    } 
                    5 { # Permanent error
                        log "Response: $Response" Magenta $SMTPLog 
                        & $UpdateReturnedStatus
                        $Script:Go = $false
                    } 
                    default { 
                        if ($ShowSMTP) { log "Response: $Response" -LogFile $SMTPLog } 
                        if ($Command -eq 'QUIT' -and $Script:Go) {
                            log "Succecded" -LogFile $SMTPLog 
                            & $UpdateReturnedStatus
                        } # if email sent successfully
                    } # default switch
                } # switch based on first number of ReplyCode
            } # if Response
        } # InvokeSMTPCommand

        # Send email
        log "Emailing $To..." -LogFile $SMTPLog -NoNewLine
        $Command = "EHLO $MyFQDN";         & $InvokeSMTPCommand
        if ($Script:Go) { $Command = "MAIL FROM:$From"     ; & $InvokeSMTPCommand }
        if ($Script:Go) { $Command = "RCPT TO:$To"         ; & $InvokeSMTPCommand }
        if ($Script:Go) { $Command = 'DATA'                ; & $InvokeSMTPCommand }
        if ($Script:Go) { 
            
            # Build SMTP body
            $Command  = "from: $SenderName $From `r`n"
            $Command += "mime-version: 1.0 `r`n"
            $Command += "to: $To `r`n"
            $Command += "X-Priority: 1 `r`n"
            $Command += "Priority: urgent `r`n"
            $Command += "Importance: high `r`n"
            $Command += "date: $(Get-Date) `r`n"
            $Command += "subject: $Subject `r`n"
            $Command += "content-type: text/html; charset=us-ascii `r`n"
            $Command += "message-id: <$([Guid]::NewGuid().Guid).$(Get-Date -f yyyyMMdd.hhmmsstt)@$MyFQDN> `r`n"
            $Command += "$Body `r`n"

            # Add DKIM header
            if ($DKIM) {
                # Compute the hash of $Command (String)
                $Bytes = [System.Text.Encoding]::UTF8.GetBytes($Command)
                $Sha   = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
                $Hash  = $Sha.ComputeHash($Bytes)

                # Read the private key 
                $FileStream  = [System.IO.File]::OpenText($pemFileName)
                $pemReader   = New-Object Org.BouncyCastle.OpenSsl.PemReader($FileStream)
                try { 
                    $KeyPair = $pemReader.ReadObject() 
                    $FileStream.Close()

                    # Sign the message 
                    $Signer = [Org.BouncyCastle.Security.SignerUtilities]::GetSigner('RSA')
                    $Signer.Init($true, $KeyPair.Private)
                    $Signer.BlockUpdate($Hash, 0, $Hash.Count)
                    $Signature = $Signer.GenerateSignature()

                    # Build the DKIM header
                    $DKIMHeader  = 'DKIM-Signature: v=1; a=rsa-sha256; c=relaxed/relaxed; q=dns/txt; ' 
                    $DKIMHeader += "d=$DKIMDomain; s=$DKIMSelector; "
                    $DKIMHeader += 'h=from:mime-version:to:date:subject:content-type:message-id; '
                    $DKIMHeader += "bh=$([System.Convert]::ToBase64String($Hash)); "
                    $DKIMHeader += "b=$([System.Convert]::ToBase64String($Signature)) `r`n"
                    $Command = $DKIMHeader + $Command             
                } catch {
                    log "Failed to DKIM-sign the message, check if the private key file '$pemFileName' is a valid RSA signature file" Yellow $SMTPLog
                }
            }

            # Send SMTP command
            if ($ShowSMTP) { log "Command:  $Command" -LogFile $SMTPLog }
            $Writer.WriteLine($Command) 
            $Writer.Flush()     
            Start-Sleep -Milliseconds 100
        }
        if ($Script:Go) { $Command = "`r`n.`r`n" ; & $InvokeSMTPCommand }
        if ($Script:Go) { $Command = 'QUIT'      ; & $InvokeSMTPCommand }
    } # Process

    End {
        $Writer.Close() 
        $Stream.Close()
        $Script:ReturnedStatus
    }
}