---
layout: page
title: OGCS Version
onload: setVersion
---
<h1>OGCS Client Version: <span id="version"></span></h1>

<script>
  function setVersion() {
    var params = {};
    
    if (location.search) {
      var parts = location.search.substring(1).split('&');

      for (var i = 0; i < parts.length; i++) {
          var nv = parts[i].split('=');
          if (!nv[0]) continue;
          params[nv[0]] = nv[1] || true;
      }
    }
    var version = params.version;

    document.getElementById("version").innerHTML = version;
    gtag('event', 'version', {'event_category': "ogcs", 'event_label': version});
  }
</script>
