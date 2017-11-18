cd src
..\nuget.exe pack OutlookGoogleCalendarSync\OutlookGoogleCalendarSync.nuspec

REM In VS Package Manager
REM PM> Install-Package squirrel.windows -Version 1.7.8
REM PM> packages\squirrel.windows.1.7.8\tools\Squirrel --releasify OutlookGoogleCalendarSync.2.6.4-alpha.nupkg --no-msi --loadingGif=..\docs\images\ogcs128x128-animated.gif