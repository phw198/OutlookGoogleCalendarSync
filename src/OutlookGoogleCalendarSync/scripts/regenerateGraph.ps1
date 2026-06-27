Write-Host "Regenerating Custom Microsoft Graph Client via Kiota..." -ForegroundColor Cyan

# Latest verson of the properly mastered API YAML:-
# https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/master/openapi/v1.0/openapi.yaml

.\kiota.exe generate `
  -d "openapi.yaml" `
  -l CSharp `
  -c GraphServiceClient `
  -n OutlookGoogleCalendarSync.Outlook.Graph.CustomClient `
  -o ..\Outlook.Graph\CustomClient `
  --include-path "/me/calendars" `
  --include-path "/me/calendars/**" `
  --include-path "/me/events" `
  --include-path "/me/events/**" `
  --include-path "/rootODataError" `
  --backing-store --exclude-backward-compatible `
  --clean-output

Write-Host "Generation complete. Visual Studio will auto-patch ODataError on next build." -ForegroundColor Green
