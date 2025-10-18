set RELEASE=2.11.6-alpha
set RELEASEDIR=Release

REM Check for new NodaTime DB @ http://nodatime.org/tzdb/latest.txt

cd src
del Releases\OutlookGoogleCalendarSync-%RELEASE%-full.nupkg
del Releases\OutlookGoogleCalendarSync-%RELEASE%-delta.nupkg
..\nuget.exe pack OutlookGoogleCalendarSync\OutlookGoogleCalendarSync.nuspec
cd ..

REM Sign the stand-alone OGCS executable
REM src\packages\squirrel.windows.1.9.0\tools\signtool.exe sign /a /n "Paul Woolcock" /tr http://time.certum.pl/ /td sha256 /fd sha256 /v src\OutlookGoogleCalendarSync\bin\%RELEASEDIR%\OutlookGoogleCalendarSync.exe

REM In VS Package Manager
REM PM> Install-Package squirrel.windows -Version 1.9.0
REM PM> packages\squirrel.windows.1.9.0\tools\Squirrel --releasify OutlookGoogleCalendarSync.2.11.6-alpha.nupkg --no-msi --loadingGif=..\docs\images\ogcs128x128-animated.gif

REM Sign the Squirrel install executable
REM src\packages\squirrel.windows.1.9.0\tools\signtool.exe sign /a /n "Paul Woolcock" /tr http://time.certum.pl/ /td sha256 /fd sha256 /v src\Releases\Setup.exe

REM Build ZIP
PAUSE
del src\Releases\OGCS_Setup.exe
rename src\Releases\Setup.exe OGCS_Setup.exe

cd src\OutlookGoogleCalendarSync\bin\%RELEASEDIR%
del Portable_OGCS_v2.11.6.zip
REM https://documentation.help/7-Zip/update1.htm
"c:\Program Files\7-Zip\7z.exe" u Portable_OGCS_v2.11.5.zip -u- -up0q0r2x2y2z1w2!Portable_OGCS_v2.11.6.zip *.dll *.ps1 ErrorReportingTemplate.json logger.xml tzdb.nzd OutlookGoogleCalendarSync.exe OutlookGoogleCalendarSync.exe.config OutlookGoogleCalendarSync.pdb Console\* 

"c:\Program Files\7-Zip\7z.exe" e -y Portable_OGCS_v2.11.5.zip Microsoft.Office.Interop.Outlook.DLL
"c:\Program Files\7-Zip\7z.exe" e -y Portable_OGCS_v2.11.5.zip stdole.dll
"c:\Program Files\7-Zip\7z.exe" e -y Portable_OGCS_v2.11.5.zip "Windows Defender SmartScreen Unblock.ps1"

"c:\Program Files\7-Zip\7z.exe" a Portable_OGCS_v2.11.6.zip Microsoft.Office.Interop.Outlook.DLL
"c:\Program Files\7-Zip\7z.exe" a Portable_OGCS_v2.11.6.zip stdole.dll
"c:\Program Files\7-Zip\7z.exe" a Portable_OGCS_v2.11.6.zip "Windows Defender SmartScreen Unblock.ps1"

cd ..\..\..\..
