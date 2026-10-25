set NEW_VERSION=3.0.5
set PREVIOUS_VERSION=3.0.4
set RELEASE=%NEW_VERSION%-alpha
set RELEASEDIR=Release-v3

REM Check for new NodaTime DB @ http://nodatime.org/tzdb/latest.txt

cd src
del Releases\OutlookGoogleCalendarSync-%RELEASE%-full.nupkg
del Releases\OutlookGoogleCalendarSync-%RELEASE%-delta.nupkg

powershell.exe -NoProfile -ExecutionPolicy Bypass -File OutlookGoogleCalendarSync\scripts\signRelease.ps1 -ReleaseDirectory OutlookGoogleCalendarSync\bin\%RELEASEDIR%
if errorlevel 1 exit /b %errorlevel%

..\nuget.exe pack OutlookGoogleCalendarSync\OutlookGoogleCalendarSync.nuspec
if errorlevel 1 exit /b %errorlevel%

REM Sign the stand-alone OGCS executable <= VS post-release build now does this
REM "c:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe" sign /as /n "Paul Woolcock" /tr http://time.certum.pl/ /td sha256 /fd sha256 /v OutlookGoogleCalendarSync\bin\%RELEASEDIR%\OutlookGoogleCalendarSync.exe

REM Build the Squirrel installer
REM PM> Install-Package squirrel.windows -Version 1.9.0
packages\squirrel.windows.1.9.0\tools\Squirrel.exe --releasify OutlookGoogleCalendarSync.%RELEASE%.nupkg --no-msi --loadingGif=..\docs\images\ogcs128x128-animated.gif
if errorlevel 1 exit /b %errorlevel%

REM Sign Tool executable location on Win7
REM set SIGNTOOL="c:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe"
REM Sign Tool executable location on Win10
set SIGNTOOL="c:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe"
%SIGNTOOL% sign /as /n "Paul Woolcock" /tr http://time.certum.pl/ /td sha256 /fd sha256 /v Releases\Setup.exe
if errorlevel 1 exit /b %errorlevel%

REM Build ZIP
PAUSE
del Releases\OGCS_Setup.exe
rename Releases\Setup.exe OGCS_Setup.exe

REM exit /b
REM PAUSE

cd OutlookGoogleCalendarSync\bin\%RELEASEDIR%
del Portable_OGCS_v%NEW_VERSION%.zip
REM https://documentation.help/7-Zip/update1.htm
"c:\Program Files\7-Zip\7z.exe" u Portable_OGCS_v%PREVIOUS_VERSION%.zip -u- -up0q0r2x2y2z1w2!Portable_OGCS_v%NEW_VERSION%.zip *.dll *.ps1 ErrorReportingTemplate.json logger.xml tzdb.nzd OutlookGoogleCalendarSync.exe OutlookGoogleCalendarSync.exe.config OutlookGoogleCalendarSync.pdb Console\* 

"c:\Program Files\7-Zip\7z.exe" e -y Portable_OGCS_v%PREVIOUS_VERSION%.zip Microsoft.Office.Interop.Outlook.DLL
"c:\Program Files\7-Zip\7z.exe" e -y Portable_OGCS_v%PREVIOUS_VERSION%.zip stdole.dll
"c:\Program Files\7-Zip\7z.exe" e -y Portable_OGCS_v%PREVIOUS_VERSION%.zip "Windows Defender SmartScreen Unblock.ps1"

"c:\Program Files\7-Zip\7z.exe" a Portable_OGCS_v%NEW_VERSION%.zip Microsoft.Office.Interop.Outlook.DLL
"c:\Program Files\7-Zip\7z.exe" a Portable_OGCS_v%NEW_VERSION%.zip stdole.dll
"c:\Program Files\7-Zip\7z.exe" a Portable_OGCS_v%NEW_VERSION%.zip "Windows Defender SmartScreen Unblock.ps1"

cd ..\..\..\..
