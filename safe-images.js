(function(){
  if (window.safeImageUrl) return;
  const __imageBlobCache = new Map();
  const __inflight = new Map();
  const __DBG = (()=>{ try{ return sessionStorage.getItem('safeimg_debug')==='1'; }catch{} return false; })();
  function log(){ if (__DBG) try{ console.log.apply(console, arguments); }catch{} }
  function shouldWrap(urlStr){
    try{
      const u = new URL(urlStr, location.origin);
      const host = u.host.toLowerCase();
      const sameOrigin = (u.origin === location.origin);
      // Skip-list: do not wrap certain hosts even if wrap-all is enabled
      try{
        if (Array.isArray(window.__SAFE_IMG_SKIP_HOSTS)){
          for (const h of window.__SAFE_IMG_SKIP_HOSTS){
            const pat = String(h||'').toLowerCase().trim(); if (!pat) continue;
            if (host === pat || host.endsWith('.'+pat)) return false;
          }
        }
      }catch{}
      // Default: Supabase only
      if (/\.supabase\.co$/i.test(host)) return true;
      // Opt-in: wrap any cross-origin if flag set
      if (window.__SAFE_IMG_WRAP_ALL && !sameOrigin) return true;
      // Opt-in: host allowlist
      if (Array.isArray(window.__SAFE_IMG_HOSTS)){
        for (const h of window.__SAFE_IMG_HOSTS){
          const pat = String(h||'').toLowerCase().trim(); if (!pat) continue;
          if (host === pat || host.endsWith('.'+pat)) return true;
        }
      }
      return false;
    }catch{ return false; }
  }
  async function safeImageUrl(rawUrl){
    try{
      const s = String(rawUrl||'').trim();
      if (!s) return '';
      // Strip optional crop meta appended after '|'
      const pipe = s.indexOf('|');
      const clean = pipe >= 0 ? s.slice(0, pipe) : s;
      // Skip local/inline URLs
      if (!/^https?:/i.test(clean)) return clean;
      // Check if we should wrap
      if (!shouldWrap(clean)) return clean;
      // Reuse cached object URL if present
      if (__imageBlobCache.has(clean)) return __imageBlobCache.get(clean);
      if (__inflight.has(clean)) return await __inflight.get(clean);
      // Fetch with CORS and without credentials so third‑party cookies are not accepted
      const runner = (async () => {
        let blob = null;
        // Try cache-first, then force reload if needed
        try{
          const resp1 = await fetch(clean, { mode: 'cors', credentials: 'omit', cache: 'force-cache' });
          if (resp1 && (resp1.ok || resp1.type === 'opaque')){
            try{ blob = await resp1.blob(); }catch{}
          }
        }catch{}
        if (!blob || !blob.size){
          try{
            const resp2 = await fetch(clean, { mode: 'cors', credentials: 'omit', cache: 'reload' });
            if (resp2 && (resp2.ok || resp2.type === 'opaque')){
              try{ blob = await resp2.blob(); }catch{}
            }
          }catch{}
        }
        if (!blob || !blob.size){ return clean; }
        const obj = URL.createObjectURL(blob);
        __imageBlobCache.set(clean, obj);
        log('[safeImageUrl] wrapped', clean, '->', obj, 'size:', blob && blob.size);
        return obj;
      })();
      __inflight.set(clean, runner);
      const result = await runner.finally(() => { try{ __inflight.delete(clean); }catch{} });
      if (result && typeof result === 'string') return result;
      return clean;
    }catch{ return String(rawUrl||''); }
  }
  try{ window.safeImageUrl = safeImageUrl; }catch{}
})();
