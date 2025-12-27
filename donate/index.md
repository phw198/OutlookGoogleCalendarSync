---
layout: page
title: Donate &amp; Support OGCS
---
# Donate &amp; Support OGCS
<br/>
Thanks so much for thinking to donate :blush:  
Please make donations via PayPal - though if this is not available in your country, try Stripe.

:bulb: £10 or more to enable splash screen hiding

<span id="gAccountSection">
  Provide the Google account you are using with OGCS to ensure splash screen hiding can be activated.<br/>
  This can be found in `Settings` > `Google` > `Connected Account`.<br/><br/>
  <label for="gAccountTxt">Google Account:</label>
  <input type="text" id="gAccountTxt" name="gAccountTxt" size="50"/>
</span>

## With PayPal

<p style="text-align: center">
  <a href="#" onClick="donate('paypal');">
    <img src="{{ site.baseurl }}/images/paypal_donate_button.png" alt="PayPal - The safer, easier way to pay online." border="0" style="vertical-align: bottom">
  </a>
</p>

## With Stripe

After clicking the donate button, on the next screen please *manually* enter your Google account <b><span id="stripe-email"></span></b>into the field as depicted here: -<br/><img src="/images/stripe-donate-field.png"/>
<style>
  .stripe-donate-pill {
    background-color: #635bff; /* Stripe Purple */
    color: white;
    padding: 3px 20px;
    text-decoration: none;
    font-weight: bold;
    border-radius: 50px; 
    display: inline-block;
  }
  .stripe-donate-pill:hover {
    color: #ffffff;
  }
</style>
<p style="text-align: center">
  <a href="#" class="stripe-donate-pill" onClick="donate('stripe');" style="a:hover.color: white">Donate Now</a>
</p>



<script language="JavaScript">
  function donate(platform) {
    //handleClickEvent('outbound', 'Donate');
    
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

        // window.alert(currencyButtons[currencyCode]);

      } catch {
        donateItemName += "%20donation.%20For%20splash%20screen%20hiding,%20enter%20your%20Gmail%20address%20in%20comment%20section"
        window.location = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id="+ paypalButtonId +"&item_name="+ donateItemName;
      }
    } else if (platform == "stripe") {
      window.location = "https://donate.stripe.com/8wM4h4e981DtdgceUU";
    }
  }
  
  //Get Google account, if available, and hide the manual entry input section
  var gaccount = null;
  const inputField = document.getElementById("gAccountTxt");
  const stripeText = document.getElementById("stripe-email");
  try {
    gaccount = new URL(window.location.href).searchParams.get("id");
    if (gaccount != null && gaccount != "") {
      inputField.value = atob(gaccount);
      stripeText.innerHTML = atob(gaccount) +" ";
      document.getElementById("gAccountSection").style.display = "none";
    }
  } catch { }

  //Dynamically update the Stripe prompt text from the manual input
  const displayDiv = document.getElementById('displayArea');

  // Listen for the 'input' event for real-time updates
  inputField.addEventListener('input', (event) => {
      stripeText.innerHTML = event.target.value +" ";
  });
</script>




<!--table border="0">
  <tr><td align="center">
    <script async src="https://js.stripe.com/v3/buy-button.js"></script>

    <stripe-buy-button
      buy-button-id="buy_btn_1SXzwFRpmZ2dnHQ0d0z9Pixq"
      publishable-key="pk_live_51QBaC3RpmZ2dnHQ0iaXfoLwUv1dpwUktTb3KwfXpcID37dxXVMBkAd8V32w3tmOaRAFPqxIvRKk4BrSYP5BnrfLs00V1r3tdik"
      customer-email="foo@bar.com"
    >
    </stripe-buy-button>
  </td></tr>
</table>
<script>
      //data-payment-link="https://donate.stripe.com/prefilled_email=me%40awesome.com"
      </script-->

