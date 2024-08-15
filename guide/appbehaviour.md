---
layout: page
title: User Guide | Settings | Application Behaviour
previous: Sync Options - What
previous-url: syncoptions-what
next: Help
next-url: help
---
{% include navigation-buttons.html %}

# Application Behaviour

![Application Behaviour Settings Screenshot](options-appbehaviour.png)

**Start on login:** Set OGCS to start up when you log in to Windows. 
<p style="margin-left:40px;margin-top:-20px"><b>Delay:</b> On system’s that take a little while to get started, OGCS can encounter problems if it starts too quickly - especially if Outlook is also set to start on startup. If OGCS only seems to have problems on system startup, or security pop-ups appear due to anti-virus definitions not yet updating, try setting a delay.</p>

<p style="margin-left:40px;margin-top:-10px"><b>For all users:</b> Configure OGCS to start, with optional delay, for an user that logs in to Windows.</p>

**Hide splash screen:** If you have donated £10 or more, you have the option of stopping the startup splash screen from displaying. Optionally, also to suppress any prompts to spread the word on social media.  

**Start in tray:** Instead of showing the application window on startup, set it to only display in the Windows notification tray are  a. From here you can control the syncs or show the application by right clicking the icon.  

**Minimise to tray:** By default, OGCS will have an application icon in the main taskbar area. This option will make it only display an icon in the notification area instead.  

<img src="options-appbehaviour-close.png" align="right" width="60%"/>
**Close button minimises:** Prevent the application from closing when clicking on the standard top right “close” button. Instead it will simply minimise.   

**Show system notifications in tray when syncing:** Get notifications when a sync begins and a summary of the changes on completion. Additionally, only show the notification **if changes are found**.  

**Make application portable:** Only available, but enabled by default, if the portable ZIP application is being used.  

### Logging

**Logging level:** Set the level of detail captured in the log file. From `OFF` through to `ALL`, with `DEBUG` being the default. Ensure it is set to `DEBUG` or greater if attaching it to a [GitHub Issue]({{ site.github-repo }}/issues).  

**Feedback errors to help improve OGCS:** If you are unfortunate enough to encounter an error, there is the option to automatically report this and future errors back to the developer - which helps them proactively identify problems and fix them!  

**Anonymise calendar subjects:** Logging includes the calendar subjects (which does make understanding the logs a lot easier!), but for the security concious these can be anonymised.

**Disable telemetry:** By default the software reports basic anonymised data, such as version of OGCS and Outlook being used.  

**Create CSV files of calendar entries:** This can be turned on to aid investigation of a possible problem, usually under instruction having raised a [GitHub Issue]({{ site.github-repo }}/issues).  

## Proxy Settings

**No Proxy:** When there is a direct connection to the internet, for example, using home broadband.  

**Inherit from Internet Explorer:** Use the same settings as configured in Internet Explorer. Normally the best option if connecting 
through a corporate proxy, or use a PAC file etc.  

**Custom Settings:** Manually enter the proxy server and port, plus your proxy credentials if required.

**Browser Agent:** How the application identifies itself to proxies. Some proxies may block certain agents, or older browser versions - in this situation, click `Check` and then `Copy agent text` with the detected browser information back into the OGCS settings.


<p>&nbsp;</p>
{% include navigation-buttons.html %}
<p>&nbsp;</p>
