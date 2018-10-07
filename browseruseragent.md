<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>

# Your Web Browser User Agent

<br/>

If you are using a custom proxy but OGCS is being blocked, try replacing the browser agent in the OGCS proxy settings with the text below.

<div class='highlighter-rouge'>
  <pre class='highlight'><code id='rawUa'>Retrieving your web browser's user agent...</code></pre>
</div>
<span id="copyButton"></span>

If your web browser's user agent is not shown above, you could also try these sites:-
* [http://www.whatsmyua.info](http://www.whatsmyua.info)
* [https://helloacm.com/api/user-agent](https://helloacm.com/api/user-agent)
* [https://www.whoishostingthis.com/tools/user-agent](https://www.whoishostingthis.com/tools/user-agent)
* [https://developers.whatismybrowser.com/useragents/parse/?analyse-my-user-agent=yes#parse-useragent](https://developers.whatismybrowser.com/useragents/parse/?analyse-my-user-agent=yes#parse-useragent)

<script>
  uaApi = "http://www.whatsmyua.info/api/v1/ua";
  uaApi = "https://helloacm.com/api/user-agent/";
  
  //$(document).ready(function(){
  $.get(uaApi, function(data){
  
    //$('#rawUa').html(data[0].ua.rawUa); //<-- For whatsmyua.info
    $('#rawUa').html(data);
    /*
    var json = jQuery.parseJSON(data);    
    $('#rawUa').html(json[0].ua.rawUa);
    */
    $('#copyButton').html('<button onclick="copyUa()">Copy agent text</button>');
  });
  
  function copyUa() {
    // Get the agent text
    var copyText = document.getElementById("rawUa").innerHTML;
    
    var el = document.createElement('textarea');
    el.value = copyText;

    // Set non-editable to avoid focus and move outside of view
    el.setAttribute('readonly', '');
    el.setAttribute('type', 'hidden');
    //el.style = {position: 'absolute', left: '-9999px'};
    document.body.appendChild(el);
    el.select();
    document.execCommand('copy');
    document.body.removeChild(el);

    alert("Copied the text: " + copyText);
  }
</script>
