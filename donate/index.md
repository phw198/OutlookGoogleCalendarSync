---
layout: page
title: Donate &amp; Support OGCS
---
# Donate &amp; Support OGCS
<br/>
Thanks so much for thinking to donate :blush:  

<p style="margin-left: -22px">:bulb: £10 or more to enable splash screen hiding</p>

<span id="gAccountSection">
  Provide the Google account you are using with OGCS to ensure splash screen hiding can be activated.<br/>
  This can be found in the app under `Settings` > `Google` > `Connected Account`.<br/><br/>
  <label for="gAccountTxt">Google Account:</label>
  <input type="text" id="gAccountTxt" name="gAccountTxt" size="50"/>
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

After clicking the donate button, on the next screen please *manually* enter your Google account <b><span id="stripe-email"></span></b>into the field as depicted here: -<br/><img src="/images/stripe-donate-field.png"/>
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
          donateItemName += "%20from%20"+ document.getElementById("gAccountTxt").value
          window.location = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id="+ paypalButtonId +"&item_name="+ donateItemName;

        } else {
          let continueAnyway = confirm("Without providing your Google account, splash screen hiding cannot be automatically enabled for you.\r\n\r\nOK to continue anyway?")
          if (continueAnyway) {
            throw Error();
          }
        }

      } catch {
        donateItemName += "%20donation.%20For%20splash%20screen%20hiding,%20enter%20your%20Gmail%20address%20in%20comment%20section"
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
    if (gaccount != null && gaccount != "#" && gaccount.length > 0) {
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

  // Listen for the 'input' event for real-time updates
  inputField.addEventListener('input', (event) => {
    stripeText.innerHTML = event.target.value +" ";
  });
</script>
