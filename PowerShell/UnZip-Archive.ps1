[System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null

Function UnZip-Archive ([string]$Source, [string]$Destination)  {

	[System.IO.Compression.ZipFile]::ExtractToDirectory($Source, $Destination)

}

