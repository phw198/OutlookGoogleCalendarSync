[CmdletBinding()]
param(
	[String]$BuildType
)

function getCRC($data) {
	[String]$regex = "^CRC32  for data:\s+(\w+)$"
	$line = ($data | Select -SkipLast 1) -match $regex 
	if (![String]::IsNullOrEmpty($line)) {
		if ($line[0] -match $regex) {
			return $Matches[1]
		}
	}
}

if ($BuildType -eq "Release") {
    $pinFile = "C:\temp\pin.txt"
    if (Test-Path $pinFile) {
        Write-Host (Get-Content $pinFile)
    } else {
        Write-Host "Create a file containing the PIN in $pinFile"
    }

    & '..\..\..\..\src\packages\squirrel.windows.1.9.0\tools\signtool.exe' sign /a /n "Open Source Developer, Paul Woolcock" /tr http://time.certum.pl/ /td sha256 /fd sha256 /v OutlookGoogleCalendarSync.exe
    
	$version = (Get-Item OutlookGoogleCalendarSync.exe).VersionInfo.FileVersion
	if ($version -notmatch "\.0$") {
		$zipFile = "v$version.zip"
		& 'C:\Program Files\7-Zip\7z.exe' a $zipFile OutlookGoogleCalendarSync.exe OutlookGoogleCalendarSync.pdb
		Copy-Item $zipFile Z:\

		$output = & 'C:\Program Files\7-Zip\7z.exe' h $zipFile
		$zipCRC = getCRC $output
		$output = & 'C:\Program Files\7-Zip\7z.exe' t $zipFile -scrc OutlookGoogleCalendarSync.exe
		$exeCRC = getCRC $output

		Write-Host "Zip = ``$zipCRC`` Exe = ``$exeCrc``"
	}
}
