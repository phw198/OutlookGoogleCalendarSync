---
layout: page
title: Donate &amp; Support OGCS
---
# Donate &amp; Support OGCS

:blush: Thanks so much for thinking to donate - it's greatly appreciated and each one makes a huge difference
{: style="text-indent: -25px; padding-top: 1em"}

:point_right: £10 or more to enable splash screen hiding
{: style="text-indent: -25px"}

<span id="gAccountSection">
  Provide the Google account email you are using with OGCS to ensure splash screen hiding can be activated.<br/>
  This can be found in the app under `Settings` > `Google` > `Connected Account`.<br/><br/>
  <label for="gAccountTxt">Google Account email address:</label>
  <input type="email" id="gAccountTxt" name="gAccountTxt" size="50" required/>
  <span id="invalidEmailMsg" style="color: indianred; font-size: smaller;"></span>
</span>

## With PayPal
Please make donations via PayPal if possible.  
No account is needed and payments can be made by card in your local currency.

<p style="text-align: center">
  <a href="#" onClick="donate('paypal');">
    <img src="{{ site.baseurl }}/images/paypal_donate_button.png" alt="PayPal - The safer, easier way to pay online." border="0" style="vertical-align: bottom">
  </a>
</p>

## With Stripe
An alternative to PayPal if that's not available in your country (eg Japan).

After clicking the donate button, on the next screen please *manually* enter your Google account <b><span id="stripe-email"></span></b>into the field as depicted below. This is to ensure splash screen hiding can be activated for you: -<br/><img src="/images/stripe-donate-field.png"/>
<style>
  .stripe-donate-pill {
    background-color: #635bff; /* Stripe Purple */
    color: white;
    padding: 5px 20px;
    text-decoration: none;
    font-size: 14px;
    font-weight: bold;
    border-radius: 50px; 
    display: inline-block;
  }
  .stripe-donate-pill:hover {
    color: #ffffff;
  }
</style>
<p style="text-align: center">
  <a href="https://donate.stripe.com/8wM4h4e981DtdgceUU" class="stripe-donate-pill" onClick="handleClickEvent('outbound', 'Donate');" style="a:hover.color: white">Donate with Stripe</a>
</p>



<script language="JavaScript">
  function validateEmail(email) {
    // This pattern checks for:
    // 1. Local part (letters, numbers, and certain special characters)
    // 2. The '@' symbol
    // 3. Domain part (letters, numbers, hyphens)
    // 4. A dot followed by a TLD of at least 2 characters
    const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    return emailRegex.test(email);
  }

  function donate(platform) {
    {% if site.google_ad_testing == "off" %}
    handleClickEvent('outbound', 'Donate');
    {% endif %}
    
    if (platform == "paypal") {
      const paypalButtonId = "44DUQ7UT6WE2C";
      var donateItemName = "Outlook%20Google%20Calendar%20Sync";
      try {
        let gAccount = document.getElementById("gAccountTxt").value;
        
        if (gAccount != "#" && gAccount.length > 0) {
          if (!validateEmail(gAccount)) {
            let continueAnyway = confirm("The Google account email '"+ gAccount +"' provided is not valid.\r\n\r\nSplash screen hiding will not be available. OK to continue?")
            if (continueAnyway) {
              throw Error()
            }
          } else {
            donateItemName += "%20from%20"+ gAccount
            window.location = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id="+ paypalButtonId +"&item_name="+ donateItemName;
          }

        } else {
          let continueAnyway = confirm("Without providing your Google account, splash screen hiding cannot be automatically enabled for you.\r\n\r\nOK to continue anyway?")
          if (continueAnyway) {
            throw Error()
          }
        }

      } catch {
        donateItemName += "%20donation.%20For%20splash%20screen%20hiding,%20enter%20your%20Gmail%20address%20in%20comment%20section%20later"
        window.location = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id="+ paypalButtonId +"&item_name="+ donateItemName;
      }
    }
  }
  
  //Get Google account, if available, and hide the manual entry input section
  var gaccount = null;
  var gaccountCookie = null;
  const inputField = document.getElementById("gAccountTxt");
  const stripeText = document.getElementById("stripe-email");
  let visibility = "visible";
  try {
    gaccount = new URL(window.location.href).searchParams.get("id");
    if (gaccount != null && gaccount != "#" && gaccount.length > 0 && validateEmail(atob(gaccount)) ) {
      setCookie("googleAccount", gaccount, 30);
      visibility = "none";
    } else {
      gaccount = getCookie("googleAccount");
    }
    if (visibility == "none" || gaccount.length > 0) {
      inputField.value = atob(gaccount);
      stripeText.innerHTML = inputField.value +" ";
    }
    document.getElementById("gAccountSection").style.display = visibility;
  } catch { }

  const errorDisplay = document.getElementById('invalidEmailMsg');

  // Listen for the 'input' event for real-time updates
  inputField.addEventListener('input', (event) => {
    stripeText.innerHTML = event.target.value +" ";

    if (!validateEmail(event.target.value)) {
        errorDisplay.textContent = "The email address does not look valid.";
        errorDisplay.style.display = "block";
    } else {
        errorDisplay.style.display = "none";
    }
  });
</script>
