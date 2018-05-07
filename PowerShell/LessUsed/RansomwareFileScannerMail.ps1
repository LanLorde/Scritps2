# RansomwareFileScanner.ps1
# Scan for Ransomware generated files and send an e-mail if something was detected
# You have to edit $SearchDir and the E-mail variables!
# Script by Tim Buntrock

# Date format
$date = get-date -format d.M.yyyy_HH.mm

# Create log directory
New-Item -Path C:\Log -ItemType directory -Force
# Specify the path
$SearchDir = "C:\Test"
# Search for specified extensions
$RansomwareList = get-childitem $SearchDir -Recurse -include *.locky,*.key,*.ecc,*.ezz,*.exx,*.zzz,*.xyz,*.aaa,*.abc,*.ccc,*.vvv,*.xxx,*.ttt,*.micro,*.encrypted,*.locked,*.crypto,_crypt,*.crinf,*.r5a,*.xrtn,*.XTBL,*.crypt,*.R16M01D05,*.pzdc,*.good,*.LOL!,*.OMG!,*.RDM,*.RRK,*.encryptedRSA,*.crjoker,*.EnCiPhErEd,*.LeChiffre,*.keybtc@inbox_com,*.0x0,*.bleep,*.1999,*.vault,*.HA3,*.toxcrypt,*.magic,*.SUPERCRYPT,*.CTBL,*.CTB2,*.locky,HELPDECRYPT.TXT,HELP_YOUR_FILES.TXT,HELP_TO_DECRYPT_YOUR_FILES.txt,RECOVERY_KEY.txt,HELP_RESTORE_FILES.txt,HELP_RECOVER_FILES.txt,HELP_TO_SAVE_FILES.txt,DecryptAllFiles.txt,DECRYPT_INSTRUCTIONS.TXT,INSTRUCCIONES_DESCIFRADO.TXT,How_To_Recover_Files.txt,YOUR_FILES.HTML,YOUR_FILES.url,encryptor_raas_readme_liesmich.txt,Help_Decrypt.txt,DECRYPT_INSTRUCTION.TXT,HOW_TO_DECRYPT_FILES.TXT,ReadDecryptFilesHere.txt,Coin.Locker.txt,_secret_code.txt,About_Files.txt,Read.txt,DECRYPT_ReadMe.TXT,DecryptAllFiles.txt,FILESAREGONE.TXT,IAMREADYTOPAY.TXT,HELLOTHERE.TXT,READTHISNOW!!!.TXT,SECRETIDHERE.KEY,IHAVEYOURSECRET.KEY,SECRET.KEY,HELPDECYPRT_YOUR_FILES.HTML,help_decrypt_your_files.html,HELP_TO_SAVE_FILES.txt,RECOVERY_FILES.txt,RECOVERY_FILE.TXT,RECOVERY_FILE*.txt,HowtoRESTORE_FILES.txt,HowtoRestore_FILES.txt,howto_recover_file.txt,restorefiles.txt,howrecover+*.txt,_how_recover.txt,recoveryfile*.txt,recoverfile*.txt,recoveryfile*.txt,Howto_Restore_FILES.TXT,help_recover_instructions+*.txt,_Locky_recover_instructions.txt
# Save the matching files to a text file
$RansomwareList |ft fullname |out-file C:\Log\RansomwareFileList$date.txt

$RansomwareListMail = "C:\Log\RansomwareFileList$date.txt"

if(($RansomwareList).length -eq 0 )

{
exit
}

else
{

# Set E-mail variables
$EmailFrom = "server@tim.tester"
$EmailTo = "Tim@tim.tester"
$Subject = "Ransomware Files detected on $env:computername"
$Body = "Find attached the Ransomware File Report for $env:computername."
$SMTPServer = "smtp01.tim.tester"
 
# Send Email and log
Send-MailMessage -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Priority High -To $EmailTo -From $EmailFrom -Attachments $RansomwareListMail

}