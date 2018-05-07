$Win81AndW2K12R2x64 = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win8.1AndW2K12R2-KB3191564-x64.msu"
$W2K12x64 = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/W2K12-KB3191565-x64.msu"
$Win7AndW2K8R2x64 = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win7AndW2K8R2-KB3191566-x64.zip"
$Win81x86 = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win8.1-KB3191564-x86.msu"
$Win7x86 = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win7-KB3191566-x86.zip"

[System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null

Function UnZip-Archive ([string]$Source, [string]$Destination)  {

	[System.IO.Compression.ZipFile]::ExtractToDirectory($Source, $Destination)

}



$version = (gcim win32_operatingsystem).version
$bitness = (gcim win32_operatingsystem).OSArchitecture


If ($version -like "10.0*") {
    Write-Host "Windows Managment Framwork 5.0 or later already installed"
}

    ElseIF ($version -like "6.3*" -and $bitness -eq "64-bit") {
            Start-BitsTransfer -Source $Win81AndW2K12R2x64 -Destination ~\Desktop\WMF5.1.msu
            ~\Desktop\WMF5.1.msu /quiet /norestart
    }

    ElseIF ($version -like "6.3*" -and $bitness -eq "32-bit") {
            Start-BitsTransfer -Source $Win81x86 -Destination ~\Desktop\WMF5.1.msu
            ~\Desktop\WMF5.1.msu /quiet /norestart
    }

    ElseIF ($version -like "6.2*" -and $bitness -eq "64-bit") {
            Start-BitsTransfer -Source $W2K12x64 -Destination ~\Desktop\WMF5.1.msu
            ~\Desktop\WMF5.1.msu /quiet /norestart
    }
    
    ElseIF ($version -like "6.1*" -and $bitness -eq "64-bit") {
            Start-BitsTransfer -Source $Win7AndW2K8R2x64 -Destination ~\Desktop\WMF5.1.zip
			Sleep 1
            UnZip-Archive -Source $ENV:USERPROFILE\Desktop\WMF5.1.zip -Destination $ENV:USERPROFILE\Desktop\WMF
			CD $ENV:USERPROFILE\Desktop\WMF
			.\Win7AndW2K8R2-KB3191566-x64.msu /quiet /norestart

    }   