(function(){

  // Early globals for editors (ONLYOFFICE + force local DOCX libs)
  try{
    if (typeof window.__DOCX_REQUIRE_LOCAL === 'undefined') window.__DOCX_REQUIRE_LOCAL = true;
    if (typeof window.ONLYOFFICE_SERVER_URL === 'undefined') window.ONLYOFFICE_SERVER_URL = 'http://127.0.0.1:8080';
    if (typeof window.ONLYOFFICE_CALLBACK_URL === 'undefined') window.ONLYOFFICE_CALLBACK_URL = 'https://dulnucqcglytdkiditbw.supabase.co/functions/v1/onlyoffice-callback';
  }catch(_){}
  
  // Local dev: relax CSP meta only on localhost to ease debugging (no inline allowed)
  try{
    var isLocal = /^(localhost|127\.0\.0\.1)$/i.test(location.hostname);
    if (isLocal){
      var m = document.getElementById('csp-meta');
      if (m){
        // Allow common CDNs and optional ONLYOFFICE origin during local dev
        var oo = (window.ONLYOFFICE_SERVER_URL || window.onlyOfficeServerUrl || '').trim();
        var ooOrigin = '';
        try{ if (oo){ var u = new URL(oo); ooOrigin = u.origin; } }catch(_){ ooOrigin=''; }
       var relaxed = "default-src 'self'; " +
  "script-src 'self' https://www.google.com https://www.gstatic.com https://cdn.jsdelivr.net https://unpkg.com https://cdnjs.cloudflare.com http://127.0.0.1:8080 http://localhost:8080 " + (ooOrigin? (ooOrigin+" ") : "") + "'unsafe-eval'; " +
  "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com https://www.gstatic.com https://cdn.jsdelivr.net; " +
  "img-src 'self' data: https: blob:; " +
  "font-src 'self' data: https://fonts.gstatic.com; " +
  "connect-src *; " +
  "frame-src 'self' blob: data: https://www.google.com https://www.gstatic.com http://127.0.0.1:8080 http://localhost:8080 " + (ooOrigin? (ooOrigin+" ") : "") + "; " +
  "child-src 'self' blob: data:; object-src 'self' blob: data:; base-uri 'self'; form-action 'self'";
        m.setAttribute('content', relaxed);
      }
    }
  }catch(_){}

  // Early: strip query parameters (never keep email/password in URL)
  try{ if (location.search) history.replaceState(null, '', location.pathname + location.hash); }catch(_){}

  
})();

