---
layout: page
title: Subscribe
---
# Subscribe for Dedicated API Quota

Have you received a message in the application that "Google's free calendar quota is being exceeded"?
{: style="margin-top:1em"}

If so, this message relates to the limited free quota offered by Google for programmatically accessing Google calendars. Users of OGCS all share a pooled quota, and as such, dependant on usage as a whole, it can be exhausted and temporarily prevents the application from working. 

For £12 per year, it is possible to be assigned into a different quota that won't be exhausted. Synchronisation features are not affected, it is simply about quota usage which developers outside of Google have limited control over. 

<span id="subscribeButton" align="center" style="display: block;">[![Subscribe Button](/images/subscribe-button.png)](#){: #subscribeButtonLink}</span>
<p id="cannotSubscribe" class="tip" align="center">:warning:The option to subscribe will only become available to those who have been affected by the quota being exhausted.<br/>You need to click the notification from within the app when it displays.</p>

:question:Already previously subscribed? Don't double up!<br/>
{: style="margin-bottom:0px; text-indent: -25px"}
- [Manage existing subscriptions](https://www.paypal.com/myaccount/autopay){: target="_blank"} in PayPal  
- &nbsp;Check for [previous payments for OGCS](https://www.paypal.com/myaccount/activities/?free_text_search=Paul+Woolcock&&type=PAYMENT_SENT){: target="_blank" #search-paypal-url}
{: style="padding-left:40px; margin-top:0px" }

<script>
  let theURL = document.getElementById("search-paypal-url");
  if (theURL) {  
    const today = new Date();
    const lastYear = new Date();
    lastYear.setFullYear(today.getFullYear() - 1);

    // Format dates as YYYY-MM-DD
    const formatDate = (date) => date.toISOString().split('T')[0];

    theURL.href = theURL.href.replace(`&&`, `&start_date=${formatDate(lastYear)}&end_date=${formatDate(today)}&`);
  }
  // window.alert(theURL.href);
</script>

## Doesn't OGCS advertise itself as free?

From its inception, keeping OGCS free (as in [libre](https://en.wikipedia.org/wiki/Gratis_versus_libre#Libre){: target="_blank"} or "liberty") and open source has been central to its development. 

Additionally, making it available for free (as in [gratis](https://en.wikipedia.org/wiki/Gratis_versus_libre#Gratis){: target="_blank"} or "free as in free beer") has also been important and the tool remains fully functional, without any restrictions to its functionality.

Subscription is specifically concerned with addressing quota limitations imposed by Google. Other options available are to wait some time before attempting to sync again, or configure your own *personal* API quota (see [Advanced/Developer Options](/guide/google#advanceddeveloper-options)).

## Quota Fair Usage
To prevent users consuming a disproportionate amount of quota, the maximum scheduled frequency of syncs are:
* Every 2 hours if Push Sync from Outlook is enabled; otherwise
* Every 15 minutes
* Restricted to 1 year past and future

It is only possible to remove these limitations through establishing your own personal API quota (see above).

## What If I Change My Computer?
Any subscription is associated with the email address of the Google account with which you are using OGCS to sync your calendar. Therefore, if you reinstall OGCS onto a new computer, as long as you sync to the **same Google account**, your subscription will retained.

### Privacy Policy
Your email address will be encrypted and stored, such that the OGCS application can automatically switch you to the guaranteed quota.

It is never used for any other purpose. Please see the full [privacy policy](/privacy-policy) for futher information.


<script language="JavaScript">
  function validateEmail(email) {
    // This pattern checks for:
    // 1. Local part (letters, numbers, and certain special characters)
    // 2. The '@' symbol
    // 3. Domain part (letters, numbers, hyphens)
    // 4. A dot followed by a TLD of at least 2 characters
    const emailRegex = /\s[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    return emailRegex.test(email);
  }
  try {
    let subscription = atob(new URL(window.location.href).searchParams.get("id"));
    const regex = /^OGCS\sPremium\s(for|renewal)\s/;
    if (subscription != null && regex.test(subscription) && validateEmail(subscription)) {
      document.getElementById("subscribeButton").style.display = "visible";
      document.getElementById("subscribeButtonLink").href = `https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=E595EQ7SNDBHA&item_name=${subscription}`;
      document.getElementById("cannotSubscribe").style.display = "none";
    } else {
      throw Error();
    }
  } catch (e){
    // window.alert(e);
    document.getElementById("subscribeButton").style.display = "none";
    document.getElementById("cannotSubscribe").style.display = "visible";
  }
</script>
