---
layout: post
title:  "Download Blocked by Microsoft Defender SmartScreen?"
date:   2026-01-29
categories: blog
---

<style>
details summary {
    font-family: 'Architects Daughter';
    display: list-item;
    cursor: pointer;
}
</style>

After downloading OGCS, if it's been released recently, you may be faced with an alert below when attempting to run the installer or the portable application executable.

🔐 It can be **resolved safely**! See below for guidance 

It just takes time for each new release to [build reputation](https://learn.microsoft.com/en-us/windows/security/operating-system-security/virus-and-threat-protection/microsoft-defender-smartscreen/#:~:text=Checking%20downloaded%20files%20against%20a%20list%20of%20files%20that%20are%20well%20known%20and%20downloaded%20frequently.%20If%20the%20file%20isn%27t%20on%20that%20list%2C%20Microsoft%20Defender%20SmartScreen%20shows%20a%20warning%2C%20advising%20caution.){: target="_blank"} with Microsoft before the warning is removed.  

<h2>Browser Download Blocked</h2>

Your organisation may prevent you from being able to keep files that are "not commonly downloaded":-  
![](/images/posts/not-commonly-downloaded.png)

<details markdown="1"><summary>See workaround &nbsp;:eyes:</summary>

### Download via PowerShell
Until enough reputation has been built to convince your organisation, try:-
1. Click "Start" and "Run" (or `Win`+`R`) and type `powershell`
1. In the console, copy and paste the below:-
```powershell
(New-Object System.Net.WebClient).DownloadFile((read-host "Paste in the download link"), "$ENV:USERPROFILE\Downloads\")
```
1. Right-click the link on the webpage for the file you previously downloaded and "Copy link address". Paste it into the Powershell window and press `enter`  
<span style="font-size:small">Eg: https://github.com/phw198/OutlookGoogleCalendarSync/releases/download/v3.0.1-alpha/OGCS_Setup.exe</span>
1. Locate the file in your `Downloads` folder and run as normal.
</details>

## Unblocking SmartScreen 
{: style="margin-top: 2em"}

SmartScreen may still block a successfully downloaded file:-  
<img width="532" height="241" alt="image" src="https://github.com/user-attachments/assets/44c20804-f148-4c5f-83a0-a9bd387f010e" />

<details markdown="1"><summary>See workaround for the <u>installer</u> &nbsp;:eyes:</summary>

### Unblocking the downloaded installer:-
1. For `OGCS_Setup.exe` you may be able to simply click `More info` and then `Run Anyway`  
🔐 You can confirm the file is <u>authentic and safe</u> as it will show the publishers name  
<img width="400" alt="image" src="https://github.com/user-attachments/assets/ccd66aa1-56e4-4b45-a182-5369e283969e" />
<img width="400" alt="image" src="https://github.com/user-attachments/assets/a873026f-104c-40c7-b4f5-8e6843e35df4" />

2. If not, right-click the `OGCS_Setup.exe` file and select `Properties`
3. In the bottom right hand corner, click the `Unblock` checkbox and then `OK`.  
🔐 Again, you can confirm the file is <u>authentic and safe</u> as it will have been "signed" by the developer  
<img height="450" alt="image" src="https://github.com/user-attachments/assets/f4ed3069-0d4b-41cf-94f2-cba7047e1215" />
<img height="450" alt="image" src="https://github.com/user-attachments/assets/53308c2c-5e70-4887-9400-3c90b9c70734" />

4. Execute `OGCS_Setup.exe`
</details>

<details markdown="1"><summary>See workaround for the <u>portable zip</u> &nbsp;:eyes:</summary>

### Unblocking the downloaded portable zip:-
1. After extracting the zip file contents, right-click the `OutlookGoogleCalendarSync.exe` file, select `Properties` and then check `Unblock` as per the images in the above installer section
1. If there is still a problem, try unblocking _all_ the extracted files. The fastest way to achieve this is:-

    1. `Winkey`+`R` and type `powershell`
    1. Change directory to the extracted files  
`cd C:\my\extracted\files\directory`
    1. `Get-ChildItem | Unblock-File`
    1. Now run OGCS by double clicking the executable.
    {: class="indent-list" xstyle="margin-left: 30px; margin-top: -1em;" }
</details>

# Happy Syncing!
{: style="margin-top: 1em"}
