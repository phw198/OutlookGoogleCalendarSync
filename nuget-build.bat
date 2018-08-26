set RELEASE=2.7.6-alpha

cd src
del Releases\OutlookGoogleCalendarSync-%RELEASE%-full.nupkg
del Releases\OutlookGoogleCalendarSync-%RELEASE%-delta.nupkg
..\nuget.exe pack OutlookGoogleCalendarSync\OutlookGoogleCalendarSync.nuspec
cd ..

REM Sign the stand-alone OGCS executable
REM src\packages\squirrel.windows.1.7.8\tools\signtool.exe sign /n "Open Source Developer, Paul Woolcock" /tr http://time.certum.pl/ /td sha256 /fd sha256 /v src\OutlookGoogleCalendarSync\bin\Release\OutlookGoogleCalendarSync.exe

REM In VS Package Manager
REM PM> Install-Package squirrel.windows -Version 1.7.8
REM PM> packages\squirrel.windows.1.7.8\tools\Squirrel --releasify OutlookGoogleCalendarSync.2.7.6-alpha.nupkg --no-msi --loadingGif=..\docs\images\ogcs128x128-animated.gif

REM Sign the Squirrel install executable
REM src\packages\squirrel.windows.1.7.8\tools\signtool.exe sign /n "Open Source Developer, Paul Woolcock" /tr http://time.certum.pl/ /td sha256 /fd sha256 /v src\Releases\Setup.exe

REM Build ZIP
PAUSE
cd src\OutlookGoogleCalendarSync\bin\Release
"c:\Program Files\7-Zip\7z.exe" u Portable_OGCS_v2.7.5.zip -u- -up0q0r2x2y2z1w2!Portable_OGCS_v2.7.6.zip *.dll *.ps1 ErrorReportingTemplate.json logger.xml tzdb.nzd OutlookGoogleCalendarSync.exe OutlookGoogleCalendarSync.exe.config Console\* 

"c:\Program Files\7-Zip\7z.exe" e Portable_OGCS_v2.7.5.zip Microsoft.Office.Interop.Outlook.DLL
"c:\Program Files\7-Zip\7z.exe" a Portable_OGCS_v2.7.6.zip Microsoft.Office.Interop.Outlook.DLL

cd ..\..\..\..