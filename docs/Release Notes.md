| Release Notes |
|:---|

# v2.9.3.0 - Alpha

**Enhancements**
- The arrival of Profiles brings multi-calendar sync! :tada: 

**Bugfix**
- Don't error if Google event has no "popup" notification
- Rolled back incompatible SharpCompress DLL
- Remove alt-tab icon when minimised to system tray

# v2.9.2.0 - Alpha

**Enhancements**
- Don't steal focus for messageboxes, flash icon instead
- Improved UX for category/colour map tests
- Proactively offer hiding of splash screen following donation

**Bugfix**
- Fix missing Google default notification in Outlook appointment
- Fix "the given key was not in the dictionary"

# v2.9.1.0 - Alpha

**Enhancements**
- More reliable method to open browser
- Option to show and sync with hidden Google calendars
- Update of third-party DLL files

**Bugfix**
- Don't alert for alpha releases if user not opted in
- Correctly sync Google "this and following events" recurring event changes
- Fix end time of all day recurring items in non-GMT timezone

# v2.9.0.0 - Beta

**Enhancements** rolled in from Alpha releases
- Customisable attendee limit for sync
- Option to exclude sync of declined Google invitiations
- Option to exclude sync of Google 'goals'
- Option to set exact type of availablity in Outlook (eg Out of Office)
- Provision for custom timezone mapping when required for meeting organiser in different timezone
- Ability to specify custom maps between Google colours and Outlook categories
- Additional colour sync setting to only sync single category to Outlook
- Retrieve all Google calendars, not just first 30
- Order the Google calendar dropdown: primary, writeable, read-only
- Warn if read-only calendar selected for sync; deny for two-way sync
- Sync deletion in Outlook on first two-way sync (not second)
- New /t command line option to append custom text to application title
- Set Outlook organisers as having "accepted" in Google

----

# v2.8.7.0 - Alpha

**Enhancements**
- Customisable attendee limit for sync
- Set "No" as default for deletion prompts
- Advanced hidden configuration option to disconnect from Outlook between syncs
- For Push Sync, check Outlook is actually running

**Bugfix**
- Hide quota exhausted notification after correct amount of time
- Get new Google access token if expiring imminently
- Ensure milestone window is always responsive
- Actually save calendar item if "force" flag is set
- Download all delta install file(s) to allow for successful upgrade
- Check if install files are already downloaded
- Regression bug syncs every minute after manual sync, even with no schedule

# v2.8.6.0 - Alpha

**Enhancements**
- Optionally exclude declined Google invitiations from sync
- Additional colour sync setting to only sync single category to Outlook
- Don't set Outlook category to default Google calendar colour (unless mapped)
- Retrieve all Google calendars, not just first 30

**Bugfix**
- Outlook 2003: Retain recurring items starting before, but spanning sync date range
- Cater for nothing being returned when getting Google item recurrences
- Handle Google error code 410 [Gone]
- Correctly delay next sync when quota exhausted
- Google calendar colours in dropdown may be offset incorrectly by 1
- Handle failure to get email address from Exchange meeting attendee
- Fixed regression of network failure preventing future automated sync
- Telemetry popup box every time OGCS starts in system tray
- Default Google reminder errors if only email notification set on event

# v2.8.5.0 - Alpha

**Enhancements**
- Improved detection of Outlook versions after 2016
- Option to exclude sync of Google 'goals'
- Option to set exact type of availablity in Outlook (eg Out of Office)
- Sync deletion in Outlook on first two-way sync (not second)
- Log exact code location of any errors
- Log Google 500 errors as `FAIL` to avoid error reporting
- Further masking of Outlook profile email address in log file

**Bugfix**
- Quota limit error updated to match changed Google error
- Reclaim calendar entries in both Google _and_ Outlook at start of sync
- Proper calculation of recurrence end time from organiser's local time
- Compare recurrence rule if number of elements have changed
- Compare attendee cloaked email addresses properly after config change
- Handle configured category mappings for categories that no longer exist in Outlook
- Don't error if log file deletion fails
- Improved support for Outlook2003

# v2.8.4.0 - Alpha

**Enhancements**
- Ability to specify custom maps between Google colours and Outlook categories
- Detect movement of Outlook appointments into/out of sync window bounds
- Improved detection of Outlook versions after 2016

**Bugfix**
- Broken upgrade mechanism (fixed from this release onwards)
- Save timezone mappings on exit of config screen
- Issue reading ErrorReporting.json file when starting multiple instances of OGCS
- Silently reconnect to Outlook if disconnected and using Push Sync
- Syncing Google reminder but none exists
- Properly report Google API exception errors

# v2.8.3.0 - Alpha

**Enhancements**
- Provision for custom timezone mapping when required for meeting organiser in different timezone
- Handle when corporate policy/AV blocks access to current user's name in Outlook

**Bugfix**
- Handle mailboxes that have no Deleted Items folder
- Check configured Outlook mailbox and calendar are stil available on startup
- Properly detect sync direction when forcing attributes in target calendar
- Only update category/colour if configured to
- Don't attempt to sync colours for recurring series exceptions
- Take into account DST for UTC offsets
- Handle recurrence end dates with no time element
- Ensure start/end Event values are populated
- Encode CSVs to UTF8
- Fix error reporting DLL reference

# v2.8.2.0 - Alpha

**Enhancements**
- Use any configured proxy for _all_ web calls
- New `/t` command line option to append custom text to application title
- Don't attempt update of recurring Google series owned by another
- Pop-up message boxes now associated with main application form
- Improved feedback on API quota exhausted
- Ability to disable telemetry

**Bugfix**
- Convert recurrence end date to local time
- Error reporting DLLs updated

# v2.8.1.0 - Alpha

**Enhancements**
- Set Outlook organisers as having "accepted" in Google
- Order the Google calendar dropdown: primary, writeable, read-only
- Warn if read-only calendar selected for sync; deny for two-way sync
- Display custom Google calendar name, if set
- Third-party DLLs updated

**Bugfix**
- Better error handling if Outlook closed mid-sync
- Fix missing last occurrence in Google for recurring series from Outlook
- Report correct COM error code and redirect to wiki for published solution
- Don't crash after upgrade if user doesn't have access to system registry
- Fixed subscribers not picking up correct API key
- Only allow single Error Reporting window
- Don't allow sync if Error Reporting window displaying

# v2.8.0.0 - Beta

**Enhancements** rolled in from Alpha releases
- Sync colours / categories
- Option to force particular colour for synced items
- Option to not sync Outlook invites you have yet to responded to
- Syncing a common calendar to/from more than one other now supported!
- Added option for users to automatically feedback errors
- Command line parameters to support multi-instance OGCS
- Better ability to cancel a running sync
- OGCS stays responsive whilst Oauth process takes place; can be cancelled
- Detect Windows system timezone changes
- Improved Push sync mechanism
- Collapsible sections added to Sync Options configuration screen
- Properly retrieve meeting organiser's timezone
- G->O: Allow synced items to be assigned specific category (not just colour)
- Show authorised Google account on Settings tab
- Better timing of auto-retry for when quota renewed after API exhausted
- Option to configure the browser's User Agent in the GUI
- Options to use (or not) Outlook/Google calendar defaults for reminders
- Don't update last sync date if sync unsuccessful
- New window for social links; option to suppress for donors
- Redirect users with COM errors (bad Office installs) to wiki help page
- Use the OGCS logo for system tray notifications
----
# v2.7.9.0 - Alpha

**Enhancements**
- Redirect users with COM errors (bad Office installs) to wiki help page
- If GAL is blocked don't report this as an error.
- Don't access Outlook appointment organiser if GAL blocked by policy
- Silently fail if check for OGCS update errors.
- Use the OGCS logo for system tray notifications

**Bugfix**
- Copy and paste / click and drag Outlook appointment causing null references
- Check for custom Google reminders before setting default
- Improved handling of Exchange errors when obtaining attendee mail addresses
- Incorrect timezone offset for appointments in Outlook 2003

# v2.7.8.0 - Alpha

**Enhancements**
- Remove Google+ related links and API calls
- Don't update last sync date if sync unsuccessful
- New window for social links; option to suppress for donors
- Faster matching of calendar items
- New advanced setting to extirpate all OGCS custom properties

**Bugfix**
- Don't try and update master Event custom attributes
- Only compare colour/category if set to be synced
- Email address cloaking
- Don't reselect Google calendar when re-downloading list

# v2.7.7.0 - Alpha

**Enhancements**
- Better timing of auto-retry for when quota renewed after API exhausted
- Option to configure the browser's User Agent in the GUI
- Options to use (or not) Outlook/Google calendar defaults for reminders

**Bugfix**
- Persist a push sync through a restart of Outlook
- Persist selection of correct calendar when alternate mailbox is temporarily unavailable
- Select correct alternate mailbox on startup
- Handle a Google event having an end date _before_ the start date
- Reminder DND window not always applying correctly

# v2.7.6.0 - Alpha

**Enhancements**
- Syncing a common calendar to/from more than one other now supported!
- Improved logic to determine meeting organiser's timezone
- Don't delete Google event immediately after reclaiming it
- New FAIL logging level added, below Error
- Nicer output when authenticating with Google and reclaiming items
- Show authorised Google account on Settings tab

**Bugfix**
- Fixed XML namespace for obfuscation sync direction
- Stopped error reporting obstructing Squirrel events
- Fixed crash on OGCS automated startup 
- Don't try and retrieve Google event if synced from different calendar (avoid 404 errors)

# v2.7.5.0 - Alpha

**Enhancements**
- Properly retrieve meeting organiser's timezone
- Suggest manual start of Outlook if OGCS not permitted
- G->O: Allow synced items to be assigned specific category (not just colour)
- Added option for users to automatically feedback errors

**Bugfix**
- Google token expiry not calculated in UTC
- Regression: Properly release Outlook if no Outlook GUI

# v2.7.4.0 - Alpha

**Enhancements**
- Sync colours / categories
- Option to force particular colour for synced items
- Collapsible sections added to Sync Options configuration screen

# v2.7.3.0 - Alpha

**Enhancements**
- Improved Push sync mechanism
- General error handling improvements
- First code-signed release of OGCS! (Should reduce and finally eliminate anti-virus false positives)

**Bugfix**
- Cope with copy and paste of Outlook appointments
- Stop Push syncs continually firing

# v2.7.2.0 - Alpha

**Enhancements**
- Handle Apple iCloud changing immutable Outlook IDs
- Option to not sync Outlook invites you have yet to responded to
- Command line parameters to support multi-instance OGCS
- Detect Windows system timezone changes

**Bugfix**
- Only redirect to wiki once per COM error
- Turn on support for TLS1.1 and 1.2 (GitHub removed support for TLS1.0)

# v2.7.1.0 - Alpha

**Enhancements**
- Google DLLs updated to latest releases; extraneous DLLs removed
- Better ability to cancel a running sync
- OGCS stays responsive whilst Oauth process takes place; can be cancelled
- Code refactor to better prepare for future developments

**Bugfix**
- Translate annual recurrences into 12 monthly recurrences (Google apps work with this better)
- Disconnection of Google account
- Increase tolerance of when to compare recurring series exceptions
- Fixed CSV file creation and format issues

# v2.7.0.0 - Beta

:spiral_notepad: If upgrading from v2.6.0, this release will require you to reauthorise OGCS to make changes to your Google calendar.  
:pushpin: If you are using your own API key, you will need to enable Google+ in your project.

**Enhancements** rolled in from Alpha releases
- Sync output now HTML, not plain text!
- Ability to set all synced items as "Available" in target calendar
- Option to cloak Google Event attendee email addresses (prevent Google sending unwanted notifications)
- New upgrade alert window containing "What's New?" improvement details + option to skip a release.
- Help/F1 to see online user guide.
- Google authorisation process streamlined - no longer needs manual copy+paste of code.
- If Outlook address book (GAL) blocked by corporate policy, remove OGCS features in order that basic sync works.
- Only access GAL on start-up if user requested sync of meeting attendees (which can trigger Outlook security pop-up).
- Don't block subsequent scheduled syncs if network drop out caused transient failure.
- Show categories from alternative mailboxes
- More reliable setting of next sync + restarting push sync.
- Remove user configuration files upon uninstall
- Feedback on settings Save button
- Skip sync of appointment body/description, if access is denied
- Cope with Google still using obsolete Calcutta timezone.
- Additional info/tip when changing "What" attributes to be synced.
- Handle Outlook and Google calendars being in different timezones.
- Automatically invert category selection when changing between include/exclude.
- New pseudo Outlook category "No category assigned".
----
# v2.6.6.0 - Alpha

**Enhancements**
- New upgrade alert window containing "What's New?" improvement details.  
- Option to skip a release.  
- Help/F1 to see online user guide.  
- Don't block subsequent scheduled syncs if network drop out caused transient failure.  
- Continue switch to MD5.  

**Bugfix**
- Updating weekday recurrence interval back into Outlook.  
- Remember obfuscation sync direction.  
- Sync freezes if OGCS starts in notification tray.  

# v2.6.5.0 - Alpha

**Enhancements**
- Sync output now HTML, not plain text!
- Change Sync button text when shift-clicking.

**Bugfix**
- Updating weekday recurrence interval back into Outlook.

# v2.6.4.0 - Alpha

**Enhancements**
- Show categories from alternative mailboxes
- More reliable setting of next sync + restarting push sync.
- Improved upgrade experience + error handling
- Remove user configuration files upon uninstall
- Better calendar security

**Bugfix**
- Forgetting obfuscation rules
- Filter on items with multiple categories may not work

# v2.6.3.0 - Alpha

**Enhancements**
- Ability to set all synced items as "Available" in target calendar
- Choose if all synced items or just created items are enforced as available and/or private.
- Feedback on settings `Save` button
- Skip sync of appointment body/description, if access is denied
- Cope with Google still using obsolete Calcutta timezone.

**Breaking Change**
- If items were already configured to sync as "Private", this may need reconfiguring.

**Bugfix**
- Forgetting `Add attendees` setting
- Give user feedback when manually checking for update

# v2.6.2.0 - Alpha

**Enhancements**
- Option to cloak Google Event attendee email addresses.
- Additional info/tip when changing "What" attributes to be synced.
- Handle Outlook and Google calendars being in different timezones.
- Automatically invert category selection when changing between include/exclude.
- New pseudo Outlook category "No category assigned".

**Bugfix**
- Detect copied Outlook items and remove OGCS metadata from copy.
- Handle failure to update notification tray icon if not present.

# v2.6.1.0 - Alpha

:spiral_notepad: This release will require you to reauthorise OGCS to make changes to your Google calendar.  
:pushpin: If you are using your own API key, you will need to enable Google+ in your project.

**Enhancements**
- If Outlook address book (GAL) blocked by corporate policy, remove OGCS features in order that basic sync works.
- Only access GAL on start-up if user requested sync of meeting attendees (which can trigger Outlook security pop-up).
- Code refactor into new Authenticator class.
- Google authentication no longer needs manual copy+paste of code.
- GoogleAuthorizationCode form removed.
- Replaced deprecated method to retrieve email address from authenticator's account.
- Removed CodePlex to Squirrel migration code.
- Update Google Auth libraries to v1.6.0
- Update Google Calendar library to v1.5.0.55

# v2.6.0.0 - Beta

:boom: This release completes the migration to GitHub.

**Enhancements Rolled in from Alphas**
- Handle startup errors better when Outlook is too busy to attach to.
- Option to delay start of application upon login.
- Suggest startup delay if Outlook still unresponsive after timeout.
- Option to force all items as private in target calendar.
- Notification bubble help for Outlook security pop-up.
- Push users to wiki for COM errors for known fixes.
- Updated bit.ly links to point to GitHub.
- Added animated logo during installation.
- Added logo to Control Panel > Uninstall Programs.
- Reset settings to default if XML file corrupted.
- Option to provide feedback when uninstalling.
- Uninstalls CodePlex ClickOnce app.
- All references to CodePlex removed.
- Improved management of API keys.
----
# v2.5.3.0 - Alpha

**Enhancements**
- Reset settings to default if XML file corrupted.
- Notification bubble help for Outlook security pop-up.
- Handle startup errors better when Outlook is too busy to attach to.
- Suggest startup delay if Outlook still unresponsive after timeout.
- Push users to wiki for COM errors for known fixes.
- Updated bit.ly links to point to GitHub.

**Bugfix**
- Handle Outlook calendars with the same name in a mailbox.
- Only set annual recurrence if INTERVAL > 1
- Subject obfuscation failing

# v2.5.2.0 - Alpha

**Enhancements**
- Option to force all items as private in target calendar.
- Option to delay start of application upon login.
- Added animated logo during installation.
- Added logo to Control Panel > Uninstall Programs.

**Bugfix**
- No log file for startup on login / Squirrel start.
- Set a default sync direction on virgin install.

# v2.5.1.0 - Alpha

**Enhancements**
- First native GitHub release (using Squirrel). :tada:
- Option to provide feedback when uninstalling.
- Uninstalls CodePlex ClickOnce app.
- All references to CodePlex removed.
- Improved management of API keys.

**Bugfix**
- Continual retry to delete recurring all-day events
- Failure to retrieve master for Google recurring event.
- Saves failing to Google after adding calendar ID to event.
- Cannot find shared Outlook folder.
- Catch failure when checking for update on GitHub.
- DND for reminders.

# v2.5.0.0 - Beta
:boom: This will be the final<sup>1</sup> CodePlex release before we move home to GitHub.  

**Bugfix Release on CodePlex/ClickOnce**
- Pass explicit Squirrel install path (not relative to ClickOnce): show-stopper for moving to Squirrel installer.
- Determining if annual recurrence is within sync range.
- Refresh dirty cache of master recurring event.
- Reset settings to default calendar if shared calendar not found.
- Convert recurring series end dates in 4500 AD to "no end" date.
- Catch error when checking for GitHub update.
- Adding omitted Squirrel DLL library files.

<sup>1</sup>The final final release. Ahem. :blush:

# v2.4.0.0 - Beta

:boom: This will be the final CodePlex release before we move home to GitHub.  
:spiral_notepad: .NET Framework 4.5 is now required.

**Enhancements Rolled in from Alphas**
- Timezone database moved to separate file that auto-updates itself  
- Workaround Alexa GMT timezone offset problem.
- Double-click "About" paths to open.
- Donors can hide the splash screen!
- Exclude cancelled Google Events from sync.
- If API quota exhausted, postpone next sync until new quota available (08:00 UTC)
- 3-way sync supported! (Eg, central Google cal syncing to 2 Outlook calendars).
- Now saves entryID, globalID and calendarID for watertight comparison.
- CSV exports now include these extra IDs.
- Added option to use own Google API keys (thanks ixe013).
- Allow sync range >365 days if using personal API keys.
- Handle "Start on startup" error if user doesn't have rights to update registry
- Double-click on tray icon during sync will no longer abort sync.
- Detect if OGCS and Outlook are running at different security elevations (1 as Administrator)
- Alexa timezone issue backported to Outlook 2003.
- Temporarily add forceSave attribute to items that must be saved (eg GUID attributes populated).
- Default to not sync attendees.
- Tooltip warning added for syncing attendees.
- Added ability to sync a shared calendar.
- Improved reporting of calendar items which could not be created/updated in Google.
- Allow single quotes at start/end of email address local-part
- Allow sync of calendar from any Outlook folder/store (instead of just by mail account).
- Progress bar added when retrieving Outlook calendars (can take a while).
- Get exact error message when failed to retrieve refresh token & improved advice.
----
## v2.3.6.0 - Alpha
:spiral_notepad: Final .NET Framework v3.5 release.  
**Bugfix**
- Could not access settings.xml file concurrently.

## v2.3.5.0 - Alpha
**Enhancements**
- Exclude cancelled Google Events from sync.
- If API quota exhausted, postpone next sync until new quota available (08:00 UTC)
- Handle "Start on startup" error if user doesn't have rights to update registry
- Double-click on tray icon during sync will no longer abort sync.
- Detect if OGCS and Outlook are running at different security elevations (1 as Administrator)
- Alexa timezone issue backported to Outlook 2003.

**Bugfix**
- Handle Outlook recipients with no AddressEntry.Type
- Hiding of splash screen.
- Recurrence for "last day of month" O->G
- GDI+ errors on Outlook setting tab.

## v2.3.4.0 - Alpha
**Enhancements**
- Temporarily add forceSave attribute to items that must be saved (eg GUID attributes populated).
- Default to not sync attendees.
- Tooltip warning added for syncing attendees.
- Framework for providing pool of API keys. Google are getting grumpy about quota.
- Added ability to sync a shared calendar.
- Removed EWS configuration.
- Disabled Alternate Mailbox and Shared Calendar options for Outlook2003.

## v2.3.3.0 - Alpha
**Enhancements**
- Workaround Alexa GMT timezone offset problem.
- Donors can hide the splash screen!
- 3-way sync supported! (Eg, central Google cal syncing to 2 Outlook calendars).
- Now saves entryID, globalID and calendarID for watertight comparison.
- CSV exports now include these extra IDs.
- Improved use of HRresult error codes and hex values.
- Improved reporting of calendar items which could not be created/updated in Google.

**Bugfix**
- Handle Outlook recipients with no valid email better.
- Handle regex against appointments with null Subjects
- Subscription expiry/rollover

## v2.3.2.0 - Alpha
**Enhancements**
- Only match "data" element of Outlook global ID
- Timezone database moved to separate file that auto-updates itself
- Allow single quotes at start/end of email address local-part
- Log exception type and code number for errors (need to remove reliance on English error message text)

**Bugfix**
- Skip corrupted items and inform user of the problematic item(s)
- Skip recurring iCalendar Events that have no RRULE - Outlook does not support them.
- Handle non-Exchange recipients with no AddressEntry.
- Ensure notification tray icon is initialised before using it for LogBox errors.
- Mark cached Event exceptions dirty when master is updated.
- Syncing to Outlook may cause some items to be deleted incorrectly.

## v2.3.1.0 - Alpha
**Enhancements**
- Allow sync of calendar from any Outlook folder/store (instead of just by mail account).
- Progress bar added when retrieving Outlook calendars (can take a while).
- Added option to use own Google API keys (thanks ixe013).
- Customise Google API error message back to user.
- Allow sync range >365 days if using personal API keys.
- Double-click "About" paths to open.
- Updated NodaTime.dll
- Get exact error message when failed to retrieve refresh token & improved advice.

**Bugfix**
- Sync recurring annual if month is within sync range
- Regression - don't always do dummy update for Appointments

# v2.3.0.0 - Beta
**Enhancements Rolled in from Alphas**
- Enable, disable, delay sync options added to right-click task-tray menu.
- Added DND time range for reminders in Google.
- Purge log files older than 30 days.
- Ability to filter Outlook categories.
- Right-click menu added to Outlook category filter checkboxlist.
- Close handle to Outlook if application not running in foreground.
- Improved instantiation of Outlook if not already running.
- Handle empty settings.xml file.
- Sync timezones properly.
- Proxy improvements.
- Capture error when G revokes token.
- Give hint the sync summary is NOT the logfile.
- Improve button and tab background colours.
- Catch COM / DLL exceptions.
- Improved handling of old IANA timezones (eg. Calcutta/Kolkata)
- Improved display at >100% display magnification
- Added "Auto Sync" menu to right-click. Includes option to delay next sync.
- Sync timer code reorganised into class.
- Improved exportToCSV functions.
- Shift-click Sync button to force a compare of all items (not just if recently modified).
- Treat Events with null or public visibility as "default".
----
## v2.2.5.0 - Alpha
**Enhancements**
- Don't keep Outlook open in background if application wasn't running in foreground.
- Improved instantiation of Outlook if not already running.
- Handle empty settings.xml file.

**Bugfix**
- Error updating recurrence exception in Google.

## v2.2.4.0 - Alpha
**Enhancements**
- Right-click menu added to Outlook category filter checkboxlist.
- Sync timezones properly.
- Proxy improvements.
- Rework of COM handling for Outlook objects.

**Bugfix**
- Sync reminders setting saves correctly.
- Sync "no end date" recurring attribute.
- Caching problems with Outlook items.

## v2.2.3.0 - Alpha
**Enhancements**
- Capture error when G revokes token.
- Ability to filter O categories.
- Give hint the sync summary is NOT the logfile.

**Bugfix**
- Remember "custom O date format" setting.
- Remember "description O->G only" setting.
- Properly release O recurrence COM references.
- Cater for difference in recurrence UNTIL values (ICS vs MS)

## v2.2.2.0 - Alpha
**Enhancements**
- Added "Auto Sync" menu to right-click. Includes options to delay next sync.
- Sync timer code reorganised into class.
- Don't delete reminders in O if not synced to G due to DND.
- Improve button and tab background colours.
- Catch COM / DLL exceptions.
- Purge log files older than 30 days.
- Improved handling of old IANA timezones (eg. Calcutta/Kolkata)
- Improved display at >100% display magnification

**Bugfix**
- Don't throw away Event after updating it
- Don't show program in alt-tab when minimised to task tray
- Handle pseudo all-day events (midnight to midnight times)

## v2.2.1.0 - Alpha
**Enhancements**
- Added DND time range for reminders in Google.
- Improved exportToCSV functions.
- Shift-click Sync button to force a compare of all items (not just if recently modified).

**Bugfix**
- Retrieve original yearly recurrence outside of sync range.
- Recalculate sync range on every sync, not just load/settings update.
- G->O for recurrence exception not syncing.
- Re-select the right Outlook calendar when re-connecting to Outlook.
- Store OGCSlastModified as culture invariant date in Google.

# v2.2.0.0 - Beta
**Enhancements Rolled in from Alphas**
- Now supports 64bit Outlook.
- Changed from using Startup menu shortcut to registry key.
- Improved error handling during sync.
- Mention OGCS in version update alert.
- Handle API daily limit being exhausted.
- Added option of subscribing for guaranteed API quota.
- Added option to use Google calendar default notification.
- "About" tab includes config file location.
- "About" tab now shows location the executable is running from.
- Show tray icon after MainForm initialised.
- Get on initialising during splash screen.
- Ignore Google events without basic attributes (eg start date)
- Ultra-Fine logging level added.
- Mask email addresses unless logging at Ultra-Fine level.
- Handle MeetingItems as well as AppointmentItems.
- Discard items without a date range.
- Required permissions now include Google user ID.
- Make proxy, push sync and start on startup take effect without saving.
- New abort method to kill background sync.
- Added links to TroubleShooting Help section.
- Minimum of 10 mins sync interval.
----
## v2.1.5.0 - Alpha
**Enhancements**
- Improved error handling during sync.
- Now supports 64bit Outlook.
- Changed from using Startup menu shortcut to registry key.
- Mention OGCS in version update alert.
- Added option to use Google calendar default notification.
- "About" tab includes config file location.

**Bugfix**
- Select correct Outlook calendar on reconnect (x-thread compliant).
- Ensure we have an Google username before subscribing.
- Make splash screen disappear no matter initialisation state.
- Remove cancelled recurrence in G not yet accepted by recipient in O.

## v2.1.4.0 - Alpha
**Enhancements**
- Show tray icon after MainForm initialised.
- Get on initialising during splash screen.
- Sync occurrences attribute of recurrence.
- Ignore Google events without basic attributes (eg start date)

**Bugfix**
- Don't keep recreating startup short cut.
- Failure to obtain recipient email address.
- Handle transient "401 Unauthorised" API errors.
- Push registration happening twice during app initialisation.
- G->O syncing to wrong calendar if Outlook's restarted.
- G->O update recurrence pattern before start/end dates.

## v2.1.3.0 - Alpha
**Enhancements**
- Ultra-Fine logging level added.
- Mask email addresses unless logging at Ultra-Fine level.
- Handle MeetingItems as well as AppointmentItems.
- Discard items without a date range.

**Bugfix**
- Allow recipient emails to start with underscore.
- Error when no subscribers.
- Remove single quote marks around an email address.
- Updating Outlook date range for non-recurring items.

## v2.1.2.0 - Alpha
**Enhancements**
- Adding AbortableBackgroundWorker.cs
- Handle API daily limit being exhausted.
- Added option of subscribing for guaranteed API quota.
- Required permissions now include Google user ID.
- "About" tab now shows location the executable is running from.

**Bugfix**
- Handle appointments with no end date.
- Interval of >1 hour not working properly.
- Merging items G->O not reliable.

## v2.1.1.0 - Alpha
**Enhancements**
- Use application name for shortcut link.
- Make proxy, push sync and start on startup take effect without saving.
- New abort method to kill background sync.
- Added links to TroubleShooting Help section.
- Minimum of 10 mins sync interval.

**Bugfix**
- Force sync of exceptions when creating recurring event.
- Google signature when event has no start/end time.
- Prevent concurrent syncs (push + manual).
- Error when updating an event just migrated from appointmentID to globalID.
- Not syncing annual recurrences with month falling within sync range.
- G->O switching to/from all-day events.
- rRule UNTIL value may not include time.
- G->O item comparison logic causing duplicates.

# v2.1.0.0 - Beta
**Enhancements Rolled in from Alphas**
- Added in **2-way synchronisation**.
- Major development to **properly sync recurring items**.
- Added option to obfuscate/mask calendar's subject.
- Back-off if hit Google API limit calls/user/second.
- Balloon click after sync shows sync log screen.
- Surfaced Outlook date formatting in UI for user configuration.
- Responding to Outlook invites no longer causes Google event to be recreated.
- Added "minimise instead of close" as configurable setting.
- Menu added to notification tray.
----
## v2.0.6.0 - Alpha
**Enhancements**
- Balloon click after sync shows sync log screen.
- Proper API backoff when limit hit.
- Simplified and more efficient mechanism to reprocess Events affected by attendee API limit.

**Bugfix**
- Outlook date format initialisation.
- Inadvertently making non-recurring events recurring.
- Handle empty timezone strings.
- Recurring weekly events on 1< day of the week.
- Repeated processing of affected Events while limit still in effect.

## v2.0.5.0 - Alpha
**Bugfix**
- Reliably get Appointment ID across all versions of Outlook
- Custom code to determine if recurring Outlook exception has been deleted (unreliable API)

## v2.0.4.0 - Alpha
**Enhancements**
- Surfaced Outlook date formatting in UI for user configuration.
- Explicitly attach to Outlook process if already running.
- Better exception handling when reconnecting to Outlook.

## v2.0.3.0 - Alpha
**Enhancements**
- Responding to Outlook invites no longer causes Google event to be recreated.
- Added better Google ExtendedProperty accessors.
- Retrieve specific Google event if already know Event ID (for recurring master) - 2way sync enhancement.

**Bugfix**
- Don't release Outlook items prematurely, if 2way sync.
- New way of specifying Outlook date range - better support for non-English regions.
- Handle release version numbers with parts >9 when checking for update.

## v2.0.2.0 - Alpha
**Enhancements**
- Explicitly dereference Outlook objects and GC on close.
- Improved error handling when user requests cancel of sync on error.
- Cache Google exceptions when retrieved outside of sync date range.
- Use modification time when comparing recurring Event exceptions.

**Bugfix**
- Cancel sync if requested after an error.
- G->O Handle null timezones events.
- Getting Outlook recurrence exception if moved and deleted.
- Logic to detect if new version available.
- Start in tray crashes application when shown.

## v2.0.1.0 - Alpha
**Enhancements**
- Notification tray icon:-
   - Show bubble when minimising to tray.
   - Click bubble to suppress future notifications.
   - Single click to show application.
   - Double click to start a sync.
   - Right click for menu.
   
**Bugfix**
- Syncing from notification tray icon when window minimised.
- Handle exception and reinitialisation of correct Outlook calendar if Outlook is restarted.
- Error on startup if offline and MAPI calendar folder unavailable.
- Recurring events
   - Syncing deleted exceptions from O->G.
   - Syncing exceptions with original date outside sync range.
- Flickering when restoring window.
----
## v1.2.6.6 - Alpha
**Enhancements**
- Major development to **properly sync recurring items**.
- Added "minimise instead of close" as configurable setting.
- Handle appointment timezone info in all Outlook versions.
- Compare iCal recurrence elements individually (not entire string).
- Retrieve "master" Google events starting before sync date range.
- Compare Google recurrence pattern to Outlook when updating.

**Bugfix**
- Obfuscation - fonts and saving settings.
- Duplication of recurring events.
----
## v1.2.6.5 - Alpha
**Enhancements**
- Menu added to notification tray icon.
- Closing window does not exit application if minimising to notification tray.

**Bugfix**
- Improve building of fake email address when none available.
- Update check now uses proxy

## v1.2.6.2 - Alpha
**Enhancements**
- Added option to obfuscate/mask calendar's subject.

## v1.2.6.1 - Alpha
**Enhancements**
- Added in **2-way synchronisation**.
- Protect against Outlook RTF descriptions being reformatted in Google and then trampling back into Outlook.
- Back-off if hit Google API limit calls/user/second.

**Bugfix**
- Continues on failure of sync.
----
# v1.2.6 - Beta
**Enhancements**
- Rework for exception handling - continue sync if single appointment fails.
- Better notification bubble messages.
- Automatically check for updates (including alphas for ZIP deployments).
- More logging around Google authentication.
- Improved logging during application startup.
- Added social features (G+, twitter) & anonymous stats.

**Bugfix**
- Skip unavailable calendar folders if Exchange disconnected. (Thanks azmodan2k)
- Truncation of Google descriptions to 8Kb
- Default values on settings.xml deserialisation.
- Don't close Outlook items when reclaiming orphans
- Handle default Settings values better + don't error on unknown xml attributes.
- Miscellenous errata.
