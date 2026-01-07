(function(){
  // Local dev: relax CSP meta only on localhost to ease debugging (no inline allowed)
  try{
    var isLocal = /^(localhost|127\.0\.0\.1)$/i.test(location.hostname);
    if (isLocal){
      var m = document.getElementById('csp-meta');
      if (m){
        // Remove 'unsafe-inline' to avoid inline warnings; keep 'unsafe-eval' if needed for dev tools
        var relaxed = "default-src 'self'; script-src 'self' https://www.google.com https://www.gstatic.com 'unsafe-eval'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com https://www.gstatic.com; img-src 'self' data: https: blob:; font-src 'self' data: https://fonts.gstatic.com; connect-src *; frame-src 'self' blob: data: https://www.google.com https://www.gstatic.com; child-src 'self' blob: data:; object-src 'self' blob: data:; base-uri 'self'; form-action 'self'";
        m.setAttribute('content', relaxed);
      }
    }
  }catch(_){}

  // Early: strip query parameters (never keep email/password in URL)
  try{ if (location.search) history.replaceState(null, '', location.pathname + location.hash); }catch(_){}
})();

