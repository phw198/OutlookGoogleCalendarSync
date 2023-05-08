[CmdletBinding()]
param(
	[String]$BuildType
)

if ($BuildType -eq "Release") {
	$version = (Get-Item  OutlookGoogleCalendarSync.exe).VersionInfo.FileVersion
	if ($version -notmatch "\.0$") {
		$zipFile = "v$version.zip"
		& 'C:\Program Files\7-Zip\7z.exe' a $zipFile OutlookGoogleCalendarSync.exe OutlookGoogleCalendarSync.pdb
		Copy-Item $zipFile Z:\
	}
}
