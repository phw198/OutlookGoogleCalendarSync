---
layout: post
title:  "v2.11.0 Beta &amp; Google Permissions"
date:   2024-09-01
categories: blog
---

Since the v2.11.0 beta release of Outlook Google Calendar Sync (OGCS), Google have altered the way account permissions can be granted to the application.

Whilst you now have the ability to choose which permission(s) to grant, OGCS will not behave properly without them. 

Although [resolved in v2.11.0.3](https://github.com/phw198/OutlookGoogleCalendarSync/issues/1937#issuecomment-2323488980), it is also not possible to access the permission screen again unless manually deleting the file `Google.Apis.Auth.OAuth2.Responses.TokenResponse-user` - so please be sure to select the top checkbox when first prompted:-

![Google Permissions]({{ site.baseurl }}/images/posts/2-11-0_permissions.png)
