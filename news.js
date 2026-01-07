(function(){
  const qs = (s, r=document) => r.querySelector(s);
  const params = new URLSearchParams(window.location.search);
  const id = params.get('id');

  async function loadNews(){
    try{
      if (!window.ilkeSupabase){
        qs('#newsTitle').textContent = 'Yapılandırma bulunamadı';
        return;
      }
      if (!id){
        qs('#newsTitle').textContent = 'Haber bulunamadı';
        return;
      }
      const { data, error } = await window.ilkeSupabase
        .from('news')
        .select('*')
        .eq('id', id)
        .limit(1)
        .maybeSingle();
      if (error) throw error;
      if (!data){
        qs('#newsTitle').textContent = 'Haber bulunamadı';
        return;
      }
      qs('#newsTitle').textContent = data.title || 'Haber';
      qs('#newsMeta').textContent = data.published_at ? new Date(data.published_at).toLocaleString('tr-TR') : '';
      // Render gallery: main viewer + thumbnails
      const galRoot = qs('#newsGallery');
      const mainImg = qs('#newsMain');
      const thumbs = qs('#newsThumbs');
      const prevBtn = qs('#prevImage');
      const nextBtn = qs('#nextImage');
      const lb = qs('#imgLightbox');
      const lbImg = qs('#lbImage');
      const lbPrev = qs('#lbPrev');
      const lbNext = qs('#lbNext');
      const lbClose = qs('#lbClose');
      // helpers: strip crop meta and cache-bust url (skip blob:/data:)
      const stripMeta = (u)=>{ if(!u) return ''; const s=String(u); const i=s.indexOf('|'); return i>=0 ? s.slice(0,i) : s; };
      const bust = (u)=>{ try{ if (!u) return ''; if (/^(?:blob:|data:)/i.test(u)) return u; const x=new URL(stripMeta(u), location.origin); x.searchParams.set('v', Date.now()); return x.toString(); }catch{ return stripMeta(u); } };
      const setImgSrc = (img, url)=>{
        try{
          const base = stripMeta(url||'');
          if (!img) return;
          if (window.safeImageUrl && /^https?:/i.test(base)){
            window.safeImageUrl(base).then(su => { img.src = bust(su || base); }).catch(()=>{ img.src = bust(base); });
          } else {
            img.src = bust(base);
          }
        }catch{ if (img) img.src = stripMeta(url||''); }
      };
      if (galRoot && mainImg && thumbs){
        let gallery = [];
        try{
          if (data.gallery_urls) gallery = Array.isArray(data.gallery_urls) ? data.gallery_urls : JSON.parse(data.gallery_urls||'[]');
        }catch{ gallery = []; }
        // Prefer the actual image for detail, then cover, then first gallery
        const image = stripMeta(data.image_url||'');
        const cover = stripMeta(data.cover_image_url||'');
        const gal = (gallery||[]).map(stripMeta);
        const seen = new Set();
        const ordered = [];
        if (image && !seen.has(image)) { ordered.push(image); seen.add(image); }
        if (cover && !seen.has(cover)) { ordered.push(cover); seen.add(cover); }
        (gal||[]).forEach(u => { if (u && !seen.has(u)) { ordered.push(u); seen.add(u); } });
        if (!ordered.length && image) ordered.push(image);

        let idx = 0;
        let autoplayTimer = null;
        function stopAutoplay(){ if (autoplayTimer){ clearInterval(autoplayTimer); autoplayTimer = null; } }
        function startAutoplay(){ stopAutoplay(); if (ordered.length > 1){ autoplayTimer = setInterval(()=> show(idx + 1), 5000); } }
        function show(i){
          if (!ordered.length) { mainImg.hidden = true; thumbs.innerHTML=''; return; }
          idx = (i + ordered.length) % ordered.length;
          setImgSrc(mainImg, ordered[idx]);
          mainImg.alt = data.title || 'Haber görseli';
          mainImg.hidden = false;
          Array.from(thumbs.querySelectorAll('img')).forEach((t, ti)=>{
            if (ti === idx) t.classList.add('active'); else t.classList.remove('active');
          });
          if (prevBtn && nextBtn){
            const one = ordered.length <= 1;
            prevBtn.style.display = one ? 'none' : '';
            nextBtn.style.display = one ? 'none' : '';
          }
          // Keep lightbox image in sync if open
          if (lb && lb.style.display === 'flex' && lbImg){ setImgSrc(lbImg, ordered[idx]); lbImg.alt = mainImg.alt; }
        }
        function renderThumbs(){
          thumbs.innerHTML = '';
          ordered.forEach((u, ti)=>{
            const t = document.createElement('img');
            setImgSrc(t, u); t.alt = '';
            if (ti === idx) t.classList.add('active');
            t.addEventListener('click', ()=> show(ti));
            thumbs.appendChild(t);
          });
        }
        renderThumbs();
        show(0);
        if (prevBtn) prevBtn.onclick = ()=> show(idx - 1);
        if (nextBtn) nextBtn.onclick = ()=> show(idx + 1);
        // Autoplay: rotate images automatically
        startAutoplay();
        // Pause on hover over main image; resume on leave
        mainImg.addEventListener('mouseenter', stopAutoplay);
        mainImg.addEventListener('mouseleave', startAutoplay);
        // Pause when page hidden
        document.addEventListener('visibilitychange', ()=>{ if (document.hidden) stopAutoplay(); else startAutoplay(); });

        // Lightbox open on main image click
        function openLightbox(){ if (!lb || !lbImg) return; stopAutoplay(); lb.style.display = 'flex'; setImgSrc(lbImg, ordered[idx]); lbImg.alt = mainImg.alt; }
        function closeLightbox(){ if (!lb) return; lb.style.display = 'none'; startAutoplay(); }
        mainImg.addEventListener('click', openLightbox);
        if (lbPrev) lbPrev.onclick = ()=> show(idx - 1);
        if (lbNext) lbNext.onclick = ()=> show(idx + 1);
        if (lbClose) lbClose.onclick = closeLightbox;
        if (lb) lb.addEventListener('click', (e)=>{ if (e.target === lb) closeLightbox(); });

        // Swipe gestures for mobile (main + lightbox)
        function addSwipe(el){
          let startX = 0, startY = 0, swiping = false;
          el.addEventListener('touchstart', (e)=>{ const t = e.touches[0]; startX = t.clientX; startY = t.clientY; swiping = true; }, {passive:true});
          el.addEventListener('touchend', (e)=>{
            if (!swiping) return; swiping = false;
            const t = e.changedTouches[0]; const dx = t.clientX - startX; const dy = t.clientY - startY;
            if (Math.abs(dx) > 40 && Math.abs(dx) > Math.abs(dy)){
              if (dx < 0) show(idx + 1); else show(idx - 1);
            }
          }, {passive:true});
        }
        addSwipe(mainImg);
        if (lbImg) addSwipe(lbImg);
      }
      // Body: render full body (Markdown -> HTML) if present, otherwise fallback to summary
      const raw = (data.body && String(data.body).trim()) ? String(data.body) : String(data.summary||'');
      qs('#newsBody').innerHTML = sanitizeHtml(mdToHtml(raw));
    } catch(e){
      qs('#newsTitle').textContent = 'Haber yüklenemedi';
      console.error(e);
    }
  }

  document.addEventListener('DOMContentLoaded', loadNews);
})();

// Minimal Markdown to HTML and sanitizer (same rules as admin preview)
function mdToHtml(md){
  let html = '\n' + String(md||'') + '\n';
  html = html.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  html = html.replace(/^######\s+(.*)$/gm, '<h6>$1</h6>')
             .replace(/^#####\s+(.*)$/gm, '<h5>$1</h5>')
             .replace(/^####\s+(.*)$/gm, '<h4>$1</h4>')
             .replace(/^###\s+(.*)$/gm, '<h3>$1</h3>')
             .replace(/^##\s+(.*)$/gm, '<h2>$1</h2>')
             .replace(/^#\s+(.*)$/gm, '<h1>$1</h1>');
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1<\/strong>')
             .replace(/\*(.+?)\*/g, '<em>$1<\/em>')
             .replace(/`([^`]+)`/g, '<code>$1<\/code>');
  html = html.replace(/\[([^\]]+)\]\((https?:[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noopener noreferrer">$1<\/a>');
  html = html.replace(/^(?:\s*[-\*]\s+.+\n)+/gm, (block)=>{
    const items = block.trim().split(/\n/).map(l=> l.replace(/^\s*[-\*]\s+/, '')).map(t=>`<li>${t}<\/li>`).join('');
    return `<ul>${items}<\/ul>`;
  });
  html = html.replace(/^(?!<h\d|<ul|<li|<p|<code|<blockquote|<img|<a)(.+)$/gm, '<p>$1<\/p>');
  return html;
}

function sanitizeHtml(dirty){
  const allowedTags = new Set(['A','B','STRONG','I','EM','U','P','UL','OL','LI','H1','H2','H3','H4','H5','H6','CODE','BR','SPAN','DIV']);
  const allowedAttrs = {
    'A': new Set(['href','target','rel','style','class']),
    'P': new Set(['class','style']),
    'SPAN': new Set(['class','style']),
    'DIV': new Set(['class','style']),
    'H1': new Set(['class','style']),
    'H2': new Set(['class','style']),
    'H3': new Set(['class','style']),
    'H4': new Set(['class','style']),
    'H5': new Set(['class','style']),
    'H6': new Set(['class','style'])
  };
  const allowedStyles = new Set(['text-align','font-size','font-family','font-weight','font-style','text-decoration']);
  function cleanStyle(el){
    const style = el.getAttribute('style')||'';
    if (!style) return;
    const parts = style.split(';').map(s=>s.trim()).filter(Boolean);
    const kept = [];
    for (const part of parts){
      const [prop, valRaw] = part.split(':'); if (!prop || !valRaw) continue;
      const p = prop.trim().toLowerCase(); const v = valRaw.trim();
      if (!allowedStyles.has(p)) continue;
      if (p === 'font-size' && !/^\d{1,3}px$/.test(v)) continue;
      if (p === 'text-align' && !/^(left|right|center|justify)$/.test(v)) continue;
      if (p === 'font-family' && /[;{}]/.test(v)) continue;
      kept.push(`${p}:${v}`);
    }
    if (kept.length) el.setAttribute('style', kept.join('; ')); else el.removeAttribute('style');
  }
  const tmp = document.createElement('div');
  tmp.innerHTML = dirty || '';
  const walker = document.createTreeWalker(tmp, NodeFilter.SHOW_ELEMENT, null);
  const toRemove = [];
  while (walker.nextNode()){
    const el = walker.currentNode;
    if (!allowedTags.has(el.tagName)) { toRemove.push(el); continue; }
    Array.from(el.attributes).forEach(attr => {
      if (!allowedAttrs[el.tagName] || !allowedAttrs[el.tagName].has(attr.name)) {
        el.removeAttribute(attr.name);
      }
    });
    cleanStyle(el);
    if (el.tagName === 'A'){
      const href = el.getAttribute('href')||'';
      if (!/^https?:\/\//i.test(href)) el.removeAttribute('href');
      el.setAttribute('rel','noopener noreferrer');
      el.setAttribute('target','_blank');
    }
  }
  toRemove.forEach(n => n.replaceWith(document.createTextNode(n.textContent||'')));
  return tmp.innerHTML;
}
