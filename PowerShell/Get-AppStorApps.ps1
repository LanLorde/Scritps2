Function Get-AppStorApps {
	Get-AppxPackage | Sort Name | Select Name, PackageFullName | FT
}