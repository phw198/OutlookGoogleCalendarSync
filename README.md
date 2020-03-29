# OGCS - Outlook VBA Code

The Visual Basic for Applications (VBA) code available here can be used for the following purposes:
* To start OGCS after Outlook has been started
* To prevent unexpected deletions when:
    * two-way sync is configured
    * OGCS is not running in the background
    * an already synced Outlook appointment is copied-and-pasted

The VBA project code for Outlook is accessed by pressing `alt`+`F11`. If this shows you already have code in the  `ThisOutlookSession` file, then you will most likely need to follow option #2 below.

As per [Microsoft documentation](https://support.microsoft.com/en-gb/help/290779/managing-and-distributing-outlook-visual-basic-for-vba), Microsoft Outlook only supports one VBA project at a time. There are two methods by which to use the code provided here:

1. Deploy the `VbaProject.otm` file
2. Manually create project files and paste in the code

## Enable Macros

Before any code is able to be run, macros settings need to be checked:

`File` > `Options` > `Trust Center` > `Trust Center Settings` > `Macro Settings` > Enable either
* Notifications for all macros; or
* Enable all macros

## Deploy the `VbaProject.otm` File

:warning: Only use this option if there is no code already in the VBA Outlook Project

This is covered in [Microsoft's document](https://support.microsoft.com/en-gb/help/290779/managing-and-distributing-outlook-visual-basic-for-vba) and entails the following steps.

1. Open Windows Explorer by pressing `Window key`+`e`
1. Enter `%APPDATA%\Microsoft\Outlook` in the address bar and press return
1. Having closed Outlook, rename `VbaProject.OTM` to something else, eg `VbaProject.OTM.bak`
1. Download the above `VbaProject.OTM` file from GitHub and save it to the same folder
1. Restart Outlook

Now when pressing `alt`+`F11`, you should see two class modules and code inside the `ThisOutlookSession` file.

## Manually Create Project Files

1. Download the above three `.cls` files from GitHub and save to your computer
1. Open VBA by pressing `alt`+`F11`.
1. In the left-hand pane, double click `ThisOutlookSession` to open it.
    1. If it has content in it already, you will need to work out how to merge it with OGCS code
    1. Open the downloaded `ThisOutlookSession.cls` in a text editor
    1. Copy and paste the code into the VBA editor window of the same name
1. For the other two downloaded files, click the `File` menu option and `Import File...` - select each file in turn

## Configure For Your OGCS Installation

1. Within the VBA editor, open the `ThisOutlookSession` file and navigate to the `Configure()` sub
1. Set the values of the following variables:
    * `ogcsDirectory` - to the directory where OGCS is installed/running from
    * `ogcsExecutable` - to the OGCS executable filename (this probably doesn't need changing)
    * `ogcsStartWithOutlook` - whether to have OGCS started whenever Outlook is started. This is instead of using the inbuilt options within the OGCS application (eg "start on system startup")

Whenever Outlook is now started, the code will automatically run in the background, triggered by the `Application_Startup` sub.

To start executing the code without a restart, click on a line of code within the `Application_Startup()` sub and then hit `F5`. You can also view debug output by going to `View` > `Immediate Window`.
