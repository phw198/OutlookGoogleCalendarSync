set RELEASE=2.7.2-alpha

cd src
del Releases\OutlookGoogleCalendarSync-%RELEASE%-full.nupkg
del Releases\OutlookGoogleCalendarSync-%RELEASE%-delta.nupkg
..\nuget.exe pack OutlookGoogleCalendarSync\OutlookGoogleCalendarSync.nuspec
cd ..

REM In VS Package Manager
REM PM> Install-Package squirrel.windows -Version 1.7.8
REM PM> packages\squirrel.windows.1.7.8\tools\Squirrel --releasify OutlookGoogleCalendarSync.2.7.2-alpha.nupkg --no-msi --loadingGif=..\docs\images\ogcs128x128-animated.gif