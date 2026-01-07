(function(){
  function isSupabaseUrl(u){
    try{
      const url = new URL(u, location.origin);
      const host = url.host.toLowerCase();
      if (/\.supabase\.co$/i.test(host)) return true;
      if (/\/storage\/v1\/object\//i.test(url.pathname)) return true;
      return false;
    }catch{ return false; }
  }
  function isHttpish(u){ return /^https?:/i.test(String(u||'')); }
  function isCrossOrigin(u){
    try{ const x = new URL(u, location.origin); return x.origin !== location.origin; }catch{ return false; }
  }
  function shouldIntercept(u){
    try{ return isSupabaseUrl(u) || (window.__SAFE_IMG_WRAP_ALL && isCrossOrigin(u)); }catch{ return isSupabaseUrl(u); }
  }

  // Intercept IMG.src assignments early to avoid initial network to Supabase
  try{
    const imgProto = window.HTMLImageElement && window.HTMLImageElement.prototype;
    if (imgProto){
      const desc = Object.getOwnPropertyDescriptor(imgProto, 'src');
      if (desc && desc.configurable && typeof desc.set === 'function' && typeof desc.get === 'function'){
        const ORIG_SET = desc.set; const ORIG_GET = desc.get;
        Object.defineProperty(imgProto, 'src', {
          get(){ try{ return ORIG_GET.call(this); }catch{ return ''; } },
          set(val){
            try{
              const s = String(val||'');
              if (!isHttpish(s) || !shouldIntercept(s) || typeof window.safeImageUrl !== 'function'){
                return ORIG_SET.call(this, s);
              }
              if (this.__safeImgLock) return ORIG_SET.call(this, s);
              const self = this;
              // Prevent immediate third‑party network; set blank first
              ORIG_SET.call(self, '');
              (async ()=>{
                try{
                  const su = await window.safeImageUrl(s);
                  self.__safeImgLock = true;
                  ORIG_SET.call(self, su || s);
                }catch{
                  try{ ORIG_SET.call(self, s); }catch{}
                }finally{
                  self.__safeImgLock = false;
                }
              })();
            }catch{ try{ ORIG_SET.call(this, val); }catch{} }
          },
          configurable: true
        });
      }
    }
  }catch{}

  // Helper: extract URL inside css url("...")
  function extractUrl(val){
    try{
      const s = String(val||'');
      const m = s.match(/url\(("|')?(.*?)\1\)/i);
      return m && m[2] ? m[2] : '';
    }catch{ return ''; }
  }

  // Intercept backgroundImage property assignment on style objects
  try{
    const styleProto = window.CSSStyleDeclaration && window.CSSStyleDeclaration.prototype;
    if (styleProto){
      // backgroundImage setter
      const descBg = Object.getOwnPropertyDescriptor(styleProto, 'backgroundImage');
      if (descBg && descBg.configurable && typeof descBg.set === 'function' && typeof descBg.get === 'function'){
        const ORIG_BG_SET = descBg.set; const ORIG_BG_GET = descBg.get;
        Object.defineProperty(styleProto, 'backgroundImage', {
          get(){ try{ return ORIG_BG_GET.call(this); }catch{ return ''; } },
          set(val){
            try{
              const v = String(val||'');
              const u = extractUrl(v);
              if (!isHttpish(u) || !shouldIntercept(u) || typeof window.safeImageUrl !== 'function' || this.__safeBgLock){
                return ORIG_BG_SET.call(this, v);
              }
              const self = this;
              // Prevent immediate third‑party network; set to none first
              ORIG_BG_SET.call(self, 'none');
              (async ()=>{
                try{
                  const su = await window.safeImageUrl(u);
                  self.__safeBgLock = true;
                  ORIG_BG_SET.call(self, su ? `url(${su})` : v);
                }catch{
                  try{ ORIG_BG_SET.call(self, v); }catch{}
                }finally{ self.__safeBgLock = false; }
              })();
            }catch{ try{ ORIG_BG_SET.call(this, val); }catch{} }
          },
          configurable: true
        });
      }
      // setProperty for background/background-image
      const ORIG_SET_PROP = styleProto.setProperty;
      if (ORIG_SET_PROP && typeof ORIG_SET_PROP === 'function'){
        styleProto.setProperty = function(name, value, priority){
          try{
            const prop = String(name||'').toLowerCase();
            const val = String(value||'');
            const wants = prop === 'background-image' || (prop === 'background' && /url\(/i.test(val));
            const u = wants ? extractUrl(val) : '';
            if (!wants || !isHttpish(u) || !shouldIntercept(u) || typeof window.safeImageUrl !== 'function' || this.__safeBgLock){
              return ORIG_SET_PROP.call(this, name, value, priority);
            }
            const self = this;
            // Prevent immediate load
            ORIG_SET_PROP.call(self, name, 'none', priority);
            (async ()=>{
              try{
                const su = await window.safeImageUrl(u);
                self.__safeBgLock = true;
                ORIG_SET_PROP.call(self, name, su ? `url(${su})` : val, priority);
              }catch{
                try{ ORIG_SET_PROP.call(self, name, val, priority); }catch{}
              }finally{ self.__safeBgLock = false; }
            })();
            return;
          }catch{ return ORIG_SET_PROP.call(this, name, value, priority); }
        };
      }
    }
  }catch{}

  // Intercept setting inline style attribute containing background-image
  try{
    const ORIG_SET_ATTR = Element.prototype.setAttribute;
    Element.prototype.setAttribute = function(name, value){
      try{
        const n = String(name||'').toLowerCase();
        const v = String(value||'');
        // Background-image in style attribute
        if (n === 'style' && /background/i.test(v) && /url\(/i.test(v) && !this.__safeBgAttrLock && typeof window.safeImageUrl === 'function'){
          const u = extractUrl(v);
          if (!isHttpish(u) || !shouldIntercept(u)){
            return ORIG_SET_ATTR.call(this, name, value);
          }
          // Set without background first to avoid load
          ORIG_SET_ATTR.call(this, name, v.replace(/background[^;]*;/ig, ''));
          const el = this;
          (async ()=>{
            try{
              const su = await window.safeImageUrl(u);
              el.__safeBgAttrLock = true;
              el.style.backgroundImage = su ? `url(${su})` : `url(${u})`;
            }catch{
              try{ ORIG_SET_ATTR.call(el, name, v); }catch{}
            }finally{ el.__safeBgAttrLock = false; }
          })();
          return;
        }
        // IMG/SOURCE src and srcset early interception
        if ((n === 'src' || n === 'srcset') && (this && this.tagName) && typeof window.safeImageUrl === 'function'){
          const tag = String(this.tagName).toUpperCase();
          if (tag === 'IMG' || tag === 'SOURCE'){
            if (n === 'src'){
              const u = v;
              if (isHttpish(u) && shouldIntercept(u) && !/^(?:blob:|data:)/i.test(u)){
                try{ ORIG_SET_ATTR.call(this, name, ''); }catch{}
                const el = this;
                (async ()=>{
                  try{ const su = await window.safeImageUrl(u); el.setAttribute('src', su || u); }catch{ try{ ORIG_SET_ATTR.call(el, name, u); }catch{} }
                })();
                return;
              }
            } else if (n === 'srcset'){
              const raw = v;
              if (/https?:/i.test(raw)){
                try{ ORIG_SET_ATTR.call(this, name, ''); }catch{}
                const el = this;
                (async ()=>{
                  try{ const rebuilt = await convertSrcsetString(raw); el.setAttribute('srcset', rebuilt || raw); }catch{ try{ ORIG_SET_ATTR.call(el, name, raw); }catch{} }
                })();
                return;
              }
            }
          }
        }
        return ORIG_SET_ATTR.call(this, name, value);
      }catch{ return ORIG_SET_ATTR.call(this, name, value); }
    };
  }catch{}

  // Helper: rebuild a srcset string with blob: URLs for Supabase items
  async function convertSrcsetString(srcset){
    try{
      const s = String(srcset||'');
      const parts = s.split(',').map(p=>p.trim()).filter(Boolean);
      const out = await Promise.all(parts.map(async part => {
        const tokens = part.split(/\s+/);
        const url = tokens.shift();
        const desc = tokens.join(' ');
        if (!url || !isHttpish(url) || !shouldIntercept(url) || /^(?:blob:|data:)/i.test(url) || typeof window.safeImageUrl !== 'function') return part;
        try{ const su = await window.safeImageUrl(url); return [su||url, desc].filter(Boolean).join(' '); }catch{ return part; }
      }));
      return out.join(', ');
    }catch{ return String(srcset||''); }
  }

  // Intercept IMG.srcset early similar to IMG.src
  try{
    const imgProto2 = window.HTMLImageElement && window.HTMLImageElement.prototype;
    if (imgProto2){
      const desc = Object.getOwnPropertyDescriptor(imgProto2, 'srcset');
      if (desc && desc.configurable && typeof desc.set === 'function' && typeof desc.get === 'function'){
        const ORIG_SET = desc.set; const ORIG_GET = desc.get;
        Object.defineProperty(imgProto2, 'srcset', {
          get(){ try{ return ORIG_GET.call(this); }catch{ return ''; } },
          set(val){
            try{
              const s = String(val||'');
              if (!/https?:/i.test(s) || typeof window.safeImageUrl !== 'function' || this.__safeImgSetLock){
                return ORIG_SET.call(this, s);
              }
              const self = this; ORIG_SET.call(self, '');
              (async ()=>{
                try{ const rebuilt = await convertSrcsetString(s); self.__safeImgSetLock = true; ORIG_SET.call(self, rebuilt || s); }
                catch{ try{ ORIG_SET.call(self, s); }catch{} }
                finally{ self.__safeImgSetLock = false; }
              })();
            }catch{ try{ ORIG_SET.call(this, val); }catch{} }
          },
          configurable: true
        });
      }
    }
  }catch{}

  async function convertImg(img){
    try{
      if (!img || !img.getAttribute) return;
      try{ if (img && img.closest && (img.closest('#pageBody.law-4688') || img.closest('[data-skip-safe-images="1"]'))) return; }catch{}
      // Handle src
      let src = img.getAttribute('src') || '';
      const fallbackSrc = img.getAttribute('data-safe-src') || '';
      if (!src && fallbackSrc) src = fallbackSrc;
      if (src && isHttpish(src) && shouldIntercept(src) && !/^(?:blob:|data:)/i.test(src) && img.dataset.blobbed !== '1' && window.safeImageUrl){
        const su = await window.safeImageUrl(src);
        if (su && typeof su === 'string'){
          img.src = su;
          try{ img.removeAttribute('data-safe-src'); }catch{}
          img.dataset.blobbed = '1';
        }
      }
      // Handle srcset
      let ss = img.getAttribute('srcset') || '';
      const fallbackSet = img.getAttribute('data-safe-srcset') || '';
      if (!ss && fallbackSet) ss = fallbackSet;
      if (ss && /https?:/i.test(ss) && !img.dataset.blobbedSrcset && typeof window.safeImageUrl === 'function'){
        try{ const rebuilt = await convertSrcsetString(ss); if (rebuilt){ img.setAttribute('srcset', rebuilt); try{ img.removeAttribute('data-safe-srcset'); }catch{} img.dataset.blobbedSrcset = '1'; } }catch{}
      }
    }catch{}
  }

  async function convertBg(el){
    try{
      if (!el || !el.style) return;
      try{ if (el && el.closest && (el.closest('#pageBody.law-4688') || el.closest('[data-skip-safe-images="1"]'))) return; }catch{}
      const bg = el.style.backgroundImage || '';
      const m = bg && bg.match(/url\(("|')?(.*?)\1\)/i);
      const url = m && m[2] ? m[2] : '';
      if (!url || !isHttpish(url)) return;
      if (!shouldIntercept(url)) return;
      if (/^(?:blob:|data:)/i.test(url)) return;
      if (!window.safeImageUrl) return;
      const su = await window.safeImageUrl(url);
      if (su && typeof su === 'string'){
        el.style.backgroundImage = `url(${su})`;
      }
    }catch{}
  }

  // Convert <source> elements (src and srcset)
  async function convertSource(el){
    try{
      if (!el || !el.getAttribute) return;
      try{ if (el && el.closest && (el.closest('#pageBody.law-4688') || el.closest('[data-skip-safe-images="1"]'))) return; }catch{}
      let src = el.getAttribute('src') || '';
      const fallbackSrc = el.getAttribute('data-safe-src') || '';
      if (!src && fallbackSrc) src = fallbackSrc;
      if (src && isHttpish(src) && shouldIntercept(src) && !/^(?:blob:|data:)/i.test(src) && window.safeImageUrl){
        try{ const su = await window.safeImageUrl(src); if (su){ el.setAttribute('src', su); try{ el.removeAttribute('data-safe-src'); }catch{} } }catch{}
      }
      let ss = el.getAttribute('srcset') || '';
      const fallbackSet = el.getAttribute('data-safe-srcset') || '';
      if (!ss && fallbackSet) ss = fallbackSet;
      if (ss && /https?:/i.test(ss) && window.safeImageUrl){
        try{ const rebuilt = await convertSrcsetString(ss); if (rebuilt){ el.setAttribute('srcset', rebuilt); try{ el.removeAttribute('data-safe-srcset'); }catch{} } }catch{}
      }
    }catch{}
  }

  function scan(root){
    try{
      try{ if (root && root.nodeType === 1 && root.closest && (root.closest('#pageBody.law-4688') || root.closest('[data-skip-safe-images="1"]'))) return; }catch{}
      const imgs = (root || document).querySelectorAll ? (root || document).querySelectorAll('img') : [];
      imgs && imgs.forEach(img => { convertImg(img); });
      const sources = (root || document).querySelectorAll ? (root || document).querySelectorAll('source') : [];
      sources && sources.forEach(s => { convertSource(s); });
      const styled = (root || document).querySelectorAll ? (root || document).querySelectorAll('[style*="background"]') : [];
      styled && styled.forEach(el => { convertBg(el); });
    }catch{}
  }

  function setupObserver(){
    try{
      const obs = new MutationObserver((mutations)=>{
        for (const m of mutations){
          if (m.type === 'childList'){
            m.addedNodes && m.addedNodes.forEach(n => {
              if (n.nodeType === 1){ // ELEMENT_NODE
                try{ if (n && n.closest && (n.closest('#pageBody.law-4688') || n.closest('[data-skip-safe-images="1"]'))) return; }catch{}
                if (n.tagName === 'IMG') convertImg(n);
                if (n.tagName === 'SOURCE') convertSource(n);
                scan(n);
              }
            });
          } else if (m.type === 'attributes'){
            const t = m.target;
            try{ if (!(t && t.closest && (t.closest('#pageBody.law-4688') || t.closest('[data-skip-safe-images="1"]')))){
              if (m.attributeName === 'src' && t && (t.tagName === 'IMG' || t.tagName === 'SOURCE')){ t.tagName === 'IMG' ? convertImg(t) : convertSource(t); }
              if (m.attributeName === 'srcset' && t && (t.tagName === 'IMG' || t.tagName === 'SOURCE')){ t.tagName === 'IMG' ? convertImg(t) : convertSource(t); }
              if (m.attributeName === 'style' && t){ convertBg(t); }
            }}catch{}
          }
        }
      });
      obs.observe(document.documentElement || document.body, {
        subtree: true,
        childList: true,
        attributes: true,
        attributeFilter: ['src','srcset','style']
      });
      // Also respond to custom slider update
      try{ document.addEventListener('slider:updated', ()=> scan(document), { once:false }); }catch{}
    }catch{}
  }

  // Pre-process nodes before insertion to avoid initial network on innerHTML/DOM inserts
  try{
    function collectCandidates(root){
      const found = [];
      try{
        try{ if (root && root.nodeType === 1 && root.closest && (root.closest('#pageBody.law-4688') || root.closest('[data-skip-safe-images="1"]'))) return found; }catch{}
        if (!root) return found;
        if (root.querySelectorAll){
          root.querySelectorAll('img[src], img[srcset], source[src], source[srcset], [style*="background"]').forEach(el => found.push(el));
        }
        const tag = (root.tagName||'').toUpperCase();
        if (tag === 'IMG' || tag === 'SOURCE' || (root.getAttribute && /background/i.test(root.getAttribute('style')||''))) found.push(root);
      }catch{}
      return found;
    }
    function preSanitize(root){
      const tasks = [];
      try{ if (root && root.nodeType === 1 && root.closest && (root.closest('#pageBody.law-4688') || root.closest('[data-skip-safe-images="1"]'))) return tasks; }catch{}
      collectCandidates(root).forEach(el => {
        try{
          const tag = (el.tagName||'').toUpperCase();
          if (tag === 'IMG' || tag === 'SOURCE'){
            const a = el.getAttribute('src');
            const ss = el.getAttribute('srcset');
            if (a && isHttpish(a) && isSupabaseUrl(a) && !/^(?:blob:|data:)/i.test(a)){
              el.setAttribute('data-safe-src', a);
              el.setAttribute('src','');
              tasks.push(()=> (tag==='IMG' ? convertImg(el) : convertSource(el)));
            }
            if (ss && /https?:/i.test(ss)){
              el.setAttribute('data-safe-srcset', ss);
              el.setAttribute('srcset','');
              tasks.push(()=> (tag==='IMG' ? convertImg(el) : convertSource(el)));
            }
          } else if (el.getAttribute){
            const st = el.getAttribute('style')||'';
            if (/background/i.test(st) && /url\(/i.test(st)){
              const u = extractUrl(st);
              if (isHttpish(u) && shouldIntercept(u)){
                // Do NOT strip existing background here to avoid racing with property interceptors.
                // Just schedule a conversion to blob: after insertion.
                tasks.push(()=> convertBg(el));
              }
            }
          }
        }catch{}
      });
      return tasks;
    }
    const ORIG_APPEND = Node.prototype.appendChild;
    const ORIG_INSERT = Node.prototype.insertBefore;
    Node.prototype.appendChild = function(node){
      try{
        try{ if (this && this.nodeType === 1 && this.closest && (this.closest('#pageBody.law-4688') || this.closest('[data-skip-safe-images="1"]'))) return ORIG_APPEND.call(this, node); }catch{}
        const tasks = preSanitize(node);
        const r = ORIG_APPEND.call(this, node);
        tasks.forEach(fn=>{ try{ fn(); }catch{} });
        return r;
      }catch{ return ORIG_APPEND.call(this, node); }
    };
    Node.prototype.insertBefore = function(node, ref){
      try{
        try{ if (this && this.nodeType === 1 && this.closest && (this.closest('#pageBody.law-4688') || this.closest('[data-skip-safe-images="1"]'))) return ORIG_INSERT.call(this, node, ref); }catch{}
        const tasks = preSanitize(node);
        const r = ORIG_INSERT.call(this, node, ref);
        tasks.forEach(fn=>{ try{ fn(); }catch{} });
        return r;
      }catch{ return ORIG_INSERT.call(this, node, ref); }
    };

    // Intercept innerHTML to sanitize string before insertion
    const elProto = Element.prototype;
    const descInner = Object.getOwnPropertyDescriptor(elProto, 'innerHTML');
    if (descInner && descInner.configurable && typeof descInner.set === 'function' && typeof descInner.get === 'function'){
      const ORIG_GET = descInner.get; const ORIG_SET = descInner.set;
      Object.defineProperty(elProto, 'innerHTML', {
        get(){ try{ return ORIG_GET.call(this); }catch{ return ''; } },
        set(val){
          try{
            try{ if (this && this.nodeType === 1 && this.closest && (this.closest('#pageBody.law-4688') || this.closest('[data-skip-safe-images="1"]'))) { ORIG_SET.call(this, val); return; } }catch{}
            const tag = (this.tagName||'').toUpperCase();
            const skipPre = /^(TABLE|THEAD|TBODY|TFOOT|TR|COLGROUP|COL|SELECT|OPTGROUP|OPTION|UL|OL|DL)$/.test(tag);
            if (skipPre){
              // Avoid parsing with DIV context which breaks table markup; rely on observer after insertion
              ORIG_SET.call(this, val);
              scan(this);
            } else {
              const tmp = document.createElement('div');
              try{ tmp.innerHTML = String(val||''); }catch{ tmp.innerHTML = val; }
              preSanitize(tmp); // clears eager-loading attributes
              ORIG_SET.call(this, tmp.innerHTML);
              // Convert live nodes inside this element after insertion
              scan(this);
            }
          }catch{ try{ ORIG_SET.call(this, val); }catch{} }
        },
        configurable: true
      });
    }
    // Intercept insertAdjacentHTML similarly
    const ORIG_IAH = Element.prototype.insertAdjacentHTML;
    if (ORIG_IAH && typeof ORIG_IAH === 'function'){
      Element.prototype.insertAdjacentHTML = function(position, text){
        try{
          try{ if (this && this.nodeType === 1 && this.closest && (this.closest('#pageBody.law-4688') || this.closest('[data-skip-safe-images="1"]'))) return ORIG_IAH.call(this, position, text); }catch{}
          const tag = (this && this.tagName||'').toUpperCase();
          const skipPre = /^(TABLE|THEAD|TBODY|TFOOT|TR|COLGROUP|COL|SELECT|OPTGROUP|OPTION|UL|OL|DL)$/.test(tag);
          if (skipPre){
            ORIG_IAH.call(this, position, text);
            scan(this);
          } else {
            const tmp = document.createElement('div');
            try{ tmp.innerHTML = String(text||''); }catch{ tmp.innerHTML = text; }
            preSanitize(tmp);
            ORIG_IAH.call(this, position, tmp.innerHTML);
            scan(this);
          }
        }catch{ return ORIG_IAH.call(this, position, text); }
      };
    }
  }catch{}

  function init(){
    scan(document);
    setupObserver();
  }

  if (document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', init, { once:true });
  } else {
    init();
  }
})();
