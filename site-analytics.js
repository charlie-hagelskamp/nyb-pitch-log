(function(){
  "use strict";

  const PRODUCTION_API = "https://script.google.com/macros/s/AKfycbwYw3hAF-O6qeiDkR9pb10OuwgXc8PrX48ZiRj2KCcrkAaPfFj9dJG4YXesRe4rsUc4/exec";
  const localHosts = ["localhost", "127.0.0.1"];
  const path = location.pathname.toLowerCase();

  if(localHosts.includes(location.hostname) || path.endsWith("/admin.html") || path.endsWith("/admin")){
    return;
  }

  function pageLabel(){
    if(path.endsWith("/drills.html")) return "Drills";
    if(path.endsWith("/fall-practice.html")) return "Fall Practice";
    if(path.endsWith("/index.html") || path.endsWith("/")) return "Pitch Log";
    return document.title || "Coach Page";
  }

  function sessionId(){
    const key = "nyb_analytics_session";
    try{
      let value = sessionStorage.getItem(key);
      if(!value){
        value = typeof crypto.randomUUID === "function"
          ? crypto.randomUUID()
          : "session-" + Date.now() + "-" + Math.random().toString(36).slice(2);
        sessionStorage.setItem(key, value);
      }
      return value;
    }catch(error){
      return "page-" + Date.now() + "-" + Math.random().toString(36).slice(2);
    }
  }

  function referrerLabel(){
    if(!document.referrer) return "Direct / unknown";
    try{
      const source = new URL(document.referrer);
      return source.hostname === location.hostname ? "Internal" : source.hostname;
    }catch(error){
      return "Direct / unknown";
    }
  }

  function deviceLabel(){
    const ua = navigator.userAgent || "";
    const touchMac = /Macintosh/.test(ua) && navigator.maxTouchPoints > 1;
    if(/iPad|Tablet|PlayBook|Silk/i.test(ua) || touchMac) return "Tablet";
    if(/Mobi|iPhone|iPod|Android/i.test(ua)) return "Mobile";
    return "Desktop";
  }

  function browserLabel(){
    const ua = navigator.userAgent || "";
    if(/Edg\//.test(ua)) return "Edge";
    if(/FxiOS|Firefox\//.test(ua)) return "Firefox";
    if(/CriOS|Chrome\//.test(ua)) return "Chrome";
    if(/Safari\//.test(ua)) return "Safari";
    return "Other";
  }

  function operatingSystemLabel(){
    const ua = navigator.userAgent || "";
    const touchMac = /Macintosh/.test(ua) && navigator.maxTouchPoints > 1;
    if(/iPhone|iPad|iPod/.test(ua) || touchMac) return "iOS";
    if(/Android/.test(ua)) return "Android";
    if(/Windows/.test(ua)) return "Windows";
    if(/CrOS/.test(ua)) return "ChromeOS";
    if(/Macintosh|Mac OS X/.test(ua)) return "macOS";
    if(/Linux/.test(ua)) return "Linux";
    return "Other";
  }

  function collect(){
    let timezone = "Unknown";
    try{
      timezone = Intl.DateTimeFormat().resolvedOptions().timeZone || "Unknown";
    }catch(error){
      // Keep the coarse fallback when timezone detection is unavailable.
    }

    return {
      eventType: "site_visit",
      page: pageLabel(),
      path: location.pathname,
      referrer: referrerLabel(),
      device: deviceLabel(),
      browser: browserLabel(),
      operatingSystem: operatingSystemLabel(),
      screen: screen.width + "x" + screen.height,
      viewport: window.innerWidth + "x" + window.innerHeight,
      language: navigator.language || "Unknown",
      timezone: timezone,
      sessionId: sessionId()
    };
  }

  function send(){
    const body = JSON.stringify(collect());
    if(navigator.sendBeacon){
      const queued = navigator.sendBeacon(
        PRODUCTION_API,
        new Blob([body], {type:"text/plain;charset=UTF-8"})
      );
      if(queued) return;
    }

    fetch(PRODUCTION_API, {
      method: "POST",
      mode: "no-cors",
      keepalive: true,
      headers: {"Content-Type":"text/plain;charset=UTF-8"},
      body: body
    }).catch(function(){
      // Analytics must never interrupt a coach's page.
    });
  }

  send();
})();
