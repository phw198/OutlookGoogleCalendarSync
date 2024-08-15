---
layout: page
title: User Guide | Settings | Sync Options - How
previous: Sync Options
previous-url: syncoptions
next: Sync Options - When
next-url: syncoptions-when
---
{% include navigation-buttons.html %}

# Sync Options

## How

<img src="options-how-more.png" alt="More Options Screenshot" align="center" />

**Direction:** One-way sync from Outlook to Google, vice versa or two-way sync between both.

These next configuration items help you manage what should happen with items that OGCS wants to remove.

**Merge with existing entries:** It is recommended the target calendar is either empty or a new calendar created specifically for OGCS to sync with. This will make it easier for you to identify what has been synced etc. However, if you need to sync into a calendar that contains items not present in the source calendar and <font style="color:red">you want to keep those items</font>, check this option. 
<div class="tip" style="margin-bottom:7px">:bulb: Such a situation may arise if you wish to sync your work Outlook calendar in to the default Google calendar, in order that Alexa or Google Home can announce your schedule for the day.</div>

**Disable deletions:** This option is mostly a safeguard to be used by new users of OGCS. Once you have confidence the tool is working as expected, it should most likely be switched off.  

**Confirm deletions:** Once deletions have been permitted, if you turn this option on, each deletion will prompt you to confirm or deny if this should be allowed. Again, this is to help you gain confidence you have the right configuration and OGCS is doing what you expect. It is not intended as a setting to have enabled long term.  
If a sync causes every single item in the target calendar to be deleted before creating new items a-new, then there is always a "bulk deletion" warning, irrespective of the above settings. This is to provide a chance to confirm this is definitely expected and intended.  

Specific attributes can be overridden in the target calendar:
<ul style="margin-top:-20px; margin-left:20px">
  <li>Set privacy as private/public</li>
  <li>Set availability as free/busy. Also tentative/out of office if syncing to Outlook</li>
  <li>Set colour/category</li>
</ul>

For two-way sync, you can choose between Outlook and Google as to which is the target calendar, as well as whether to implement the override
<ul style="margin-top:-20px; margin-left:20px">
  <li>For all synced items; or</li>
  <li>Only for items newly created by the sync tool</li>
</ul>

This second option means that the overridden attribute(s) will not sync back to the source calendar - otherwise, for example, _everything_ would end up private which is probably not what is intended.

**Word obfuscation:** Through the use of [regular expressions](https://www.regular-expressions.info), certain letters, words or phrases in the calendar item’s subject can be altered. 

<img src="options-how-regex.png" alt="Word Obfuscation Screenshot" align="right" />
By clicking the `Rules` button, you will see a table with `Find`, `Replace` and `Target` columns; each row of regular expression rules would be applied in the order given using AND logic.

The example in the image shows the entire subject text (denoted by the expression `^.*$`) being replaced with the single word "Busy".

The target column indicates which attribute(s) the regular expression should be applied - it can be any combination of:-
<ul style="margin-top:-20px; margin-left:20px">  
<li><code>S</code>ubject</li>
<li><code>L</code>ocation</li>
<li><code>D</code>escription</li>
</ul>

To apply to all attributes, for example, enter `SLD`. The default is `S` to obfuscate the subject text only.

<div class="tip">:memo: If two-way sync is configured, the obfuscation can only work in one direction - choose which from the drop down menu.</div>
<br/>

<p>&nbsp;</p>
{% include navigation-buttons.html %}
<p>&nbsp;</p>

