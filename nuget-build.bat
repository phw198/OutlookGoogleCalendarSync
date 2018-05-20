set RELEASE=2.7.4-alpha

cd src
del Releases\OutlookGoogleCalendarSync-%RELEASE%-full.nupkg
del Releases\OutlookGoogleCalendarSync-%RELEASE%-delta.nupkg
..\nuget.exe pack OutlookGoogleCalendarSync\OutlookGoogleCalendarSync.nuspec
cd ..

REM Sign the stand-alone OGCS executable
REM src\packages\squirrel.windows.1.7.8\tools\signtool.exe sign /n "Open Source Developer, Paul Woolcock" /tr http://time.certum.pl/ /td sha256 /fd sha256 /v src\OutlookGoogleCalendarSync\bin\Release\OutlookGoogleCalendarSync.exe

REM In VS Package Manager
REM PM> Install-Package squirrel.windows -Version 1.7.8
REM PM> packages\squirrel.windows.1.7.8\tools\Squirrel --releasify OutlookGoogleCalendarSync.2.7.4-alpha.nupkg --no-msi --loadingGif=..\docs\images\ogcs128x128-animated.gif

REM Sign the Squirrel install executable
REM src\packages\squirrel.windows.1.7.8\tools\signtool.exe sign /n "Open Source Developer, Paul Woolcock" /tr http://time.certum.pl/ /td sha256 /fd sha256 /v src\Releases\Setup.exe

