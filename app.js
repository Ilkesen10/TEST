'use strict';

window.IlkeSendika = (function(){
  const api = {};
  // reCAPTCHA state
  let recaptcha = {
    siteKey: null,
    widgetId: null,
    token: null
  };

  // Turkish National ID (TCKN) validation
  api.isValidTCKN = isValidTCKN;
  function isValidTCKN(t){
    try{
      const s = String(t||'').trim();
      if (!/^\d{11}$/.test(s)) return false;
      if (s[0] === '0') return false;
      const digits = s.split('').map(n=>parseInt(n,10));
      const d10 = digits[9];
      const d11 = digits[10];
      const oddSum = digits[0]+digits[2]+digits[4]+digits[6]+digits[8];
      const evenSum = digits[1]+digits[3]+digits[5]+digits[7];
      const calc10 = ((oddSum*7) - evenSum) % 10;
      if (calc10 !== d10) return false;
      const sumFirst10 = digits.slice(0,10).reduce((a,b)=>a+b,0) % 10;
      if (sumFirst10 !== d11) return false;
      return true;
    }catch{ return false; }
  }

  // Public: dynamic İl / İlçe selects on membership application page
  async function initPublicGeoSelectors(){
    try{
      if (!document.body || document.body.getAttribute('data-page') !== 'uyelik-basvurusu') return;
      const provSel = document.querySelector('select[name="work_province"]');
      const distSel = document.querySelector('select[name="work_district"]');
      if (!provSel || !distSel) return;

      const GEO_CACHE_KEY = 'ilkesen_geo_cache_v1';
      const GEO_TTL_MS = 24 * 60 * 60 * 1000; // 24h
      window.__ilkesenGeoCache = window.__ilkesenGeoCache || { provinces:null, districtsByProv:{} , ts:0 };
      async function loadCache(){
        try{
          if (window.__ilkesenGeoCache?.provinces) return window.__ilkesenGeoCache;
          const raw = localStorage.getItem(GEO_CACHE_KEY);
          if (!raw) return window.__ilkesenGeoCache;
          const obj = JSON.parse(raw);
          if (obj && Date.now() - (obj.ts||0) < GEO_TTL_MS){ window.__ilkesenGeoCache = obj; }
        }catch{}
        return window.__ilkesenGeoCache;
      }

      async function saveCache(){ try{ localStorage.setItem(GEO_CACHE_KEY, JSON.stringify(window.__ilkesenGeoCache)); }catch{} }
      async function fetchProvinces(){
        const upper = (s)=>{ try{ return String(s||'').toLocaleUpperCase('tr-TR'); }catch{ return String(s||'').toUpperCase(); } };
        const cache = await loadCache();
        if (cache.provinces){
          try{
            const out = (cache.provinces||[]).map(p => Object.assign({}, p, { name: upper(p.name) }));
            window.__ilkesenGeoCache.provinces = out; window.__ilkesenGeoCache.ts = Date.now(); await saveCache();
            return out;
          }catch{ return cache.provinces; }
        }
        // Try DB first
        try{
          if (window.ilkeSupabase){
            const { data: provs, error } = await window.ilkeSupabase.from('provinces').select('id,name,plate_code').order('name');
            if (error) throw error;
            const out = (provs||[]).map(p => ({ id: String(p.id), name: upper(p.name), plate_code: p.plate_code }));
            window.__ilkesenGeoCache.provinces = out;
            window.__ilkesenGeoCache.ts = Date.now();
            await saveCache();
            return window.__ilkesenGeoCache.provinces;
          }
        }catch{}
        // Fallback: static JSON
        try{
          const res = await fetch('data/districts_tr.json', { cache: 'no-store' });
          const obj = await res.json();
          const names = Object.keys(obj || {});
          const provs = names.map((n, idx)=>({ id: String(idx+1), name: upper(n), plate_code: null }));
          window.__ilkesenGeoCache.provinces = provs;
          window.__ilkesenGeoCache.ts = Date.now();
          await saveCache();
          return window.__ilkesenGeoCache.provinces;
        }catch{ return []; }
      }
      async function fetchDistrictsSmart({ provinceId, plateCode, provinceName }){
        const cache = await loadCache();
        const cacheKey = provinceId ? `pid:${provinceId}` : (plateCode ? `pc:${plateCode}` : (provinceName ? `pn:${provinceName}` : ''));
        if (cacheKey && cache.districtsByProv[cacheKey]) return cache.districtsByProv[cacheKey];
        // Try DB by province id
        if (provinceId){
          try{
            if (window.ilkeSupabase){
              const { data: d1, error: e1 } = await window.ilkeSupabase.from('districts').select('id,name,province_id').eq('province_id', provinceId).order('name');
              if (!e1 && d1 && d1.length){ window.__ilkesenGeoCache.districtsByProv[cacheKey] = d1; window.__ilkesenGeoCache.ts = Date.now(); await saveCache(); return d1; }
            }
          }catch{}
        }
        // Fallback by province name via static JSON (case-insensitive)
        try{
          const res = await fetch('data/districts_tr.json', { cache: 'no-store' });
          const obj = await res.json();
          const up = (s)=>{ try{ return String(s||'').toLocaleUpperCase('tr-TR'); }catch{ return String(s||'').toUpperCase(); } };
          const pk = Object.keys(obj || {}).find(k => up(k) === up(provinceName));
          const list = pk ? obj[pk] : [];
          const dists = (list||[]).map((n, idx)=>({ id: String(idx+1), name: n, province_id: provinceId || plateCode || provinceName || '' }));
          const key = cacheKey || (provinceName ? `pn:${provinceName}` : '');
          if (key){ window.__ilkesenGeoCache.districtsByProv[key] = dists; window.__ilkesenGeoCache.ts = Date.now(); await saveCache(); }
          return window.__ilkesenGeoCache.districtsByProv[key];
        }catch{}
        return [];
      }

      // Populate provinces
      const provs = await fetchProvinces();
      provSel.innerHTML = '';
      { const optNone = document.createElement('option'); optNone.value=''; optNone.textContent='İl Seçiniz'; provSel.appendChild(optNone); }
      (provs||[]).forEach(p=>{ const o=document.createElement('option'); o.value=String(p.id); o.textContent=p.name; o.dataset.plate = String(p.plate_code||p.plate||''); provSel.appendChild(o); });

      async function loadDistrictsForSelected(){
        distSel.innerHTML='';
        const o0=document.createElement('option'); o0.value=''; o0.textContent='İlçe Seçiniz'; distSel.appendChild(o0);
        const selOpt = provSel.selectedOptions && provSel.selectedOptions[0];
        const pid = provSel.value;
        const plate = selOpt ? (selOpt.dataset.plate || '') : '';
        const pname = selOpt ? (selOpt.textContent || '').trim() : '';
        if (!pid && !plate && !pname){ return; }
        const dists = await fetchDistrictsSmart({ provinceId: pid, plateCode: plate, provinceName: pname });
        (dists||[]).forEach(d=>{ const o=document.createElement('option'); o.value=String(d.id); o.textContent=d.name; distSel.appendChild(o); });
      }
      provSel.addEventListener('change', loadDistrictsForSelected);
      await loadDistrictsForSelected();
    }catch(e){ console.warn('initPublicGeoSelectors error:', e); }
  }

  // Public: enforce input filters (letters-only + TR uppercase) and digits-only for membership form
  function initMembershipFormInputFilters(){
    try{
      if (!document.body || document.body.getAttribute('data-page') !== 'uyelik-basvurusu') return;
      const form = document.getElementById('membershipForm');
      if (!form) return;
      const toUpperTR = (s)=>{ try{ return String(s||'').toLocaleUpperCase('tr-TR'); }catch{ return String(s||'').toUpperCase(); } };
      const lettersOnly = (s)=> String(s||'').replace(/[^A-Za-zÇĞİÖŞÜçğıöşü\s]/g,'').replace(/\s+/g,' ');
      const upperNames = ['first_name','last_name','father_name','mother_name','birth_place','institution_name','work_unit','title'];
      upperNames.forEach(n=>{
        const el = form.querySelector(`input[name="${n}"]`);
        if (!el) return;
        el.addEventListener('input', ()=>{ el.value = toUpperTR(lettersOnly(el.value)); });
        el.addEventListener('blur', ()=>{ el.value = toUpperTR(lettersOnly(el.value)).trim(); });
      });
      function setDigitsOnly(name, maxLen){
        const el = form.querySelector(`input[name="${name}"]`);
        if (!el) return;
        try{ el.setAttribute('inputmode','numeric'); }catch{}
        if (maxLen) { try{ el.maxLength = maxLen; }catch{} }
        el.addEventListener('input', ()=>{
          let v = (el.value||'').replace(/\D/g,'');
          if (maxLen) v = v.slice(0, maxLen);
          el.value = v;
        });
      }
      setDigitsOnly('national_id', 11);
      setDigitsOnly('corp_reg_no');
      setDigitsOnly('retirement_no');
      setDigitsOnly('ssk_no');
      const phone = form.querySelector('input[name="phone"]');
      if (phone){
        try{ phone.setAttribute('inputmode','numeric'); }catch{}
        phone.maxLength = 10;
        phone.addEventListener('input', ()=>{ phone.value = (phone.value||'').replace(/\D/g,'').slice(0,10); });
      }
    }catch(e){ console.warn('initMembershipFormInputFilters error:', e); }
  }

  // Membership Application form handler
  api.handleMembershipApplySubmit = async function(e){
    e.preventDefault();
    const form = e.target;
    const fd = new FormData(form);
    // Read fields
    let first_name = String(fd.get('first_name')||'').trim();
    let last_name = String(fd.get('last_name')||'').trim();
    let national_id = String(fd.get('national_id')||'').trim();
    let father_name = String(fd.get('father_name')||'').trim();
    let mother_name = String(fd.get('mother_name')||'').trim();
    let birth_place = String(fd.get('birth_place')||'').trim();
    const birth_date = String(fd.get('birth_date')||'').trim();
    const gender = String(fd.get('gender')||'').trim();
    const education = String(fd.get('education')||'').trim();
    // Province/District as selects
    const provSel = form.querySelector('select[name="work_province"]');
    const distSel = form.querySelector('select[name="work_district"]');
    const work_province_val = String(fd.get('work_province')||'').trim();
    const work_district_val = String(fd.get('work_district')||'').trim();
    let institution_name = String(fd.get('institution_name')||'').trim();
    let work_unit = String(fd.get('work_unit')||'').trim();
    let work_unit_address = String(fd.get('work_unit_address')||'').trim();
    let corp_reg_no = String(fd.get('corp_reg_no')||'').trim();
    let title = String(fd.get('title')||'').trim();
    const blood_type = String(fd.get('blood_type')||'').trim();
    let retirement_no = String(fd.get('retirement_no')||'').trim();
    let ssk_no = String(fd.get('ssk_no')||'').trim();
    let email = String(fd.get('email')||'').trim();
    let phone = String(fd.get('phone')||'').trim();
    const kvkk = form.querySelector('input[name="kvkk"]');

    // Normalize per requirements (uppercase letters-only for specified fields; digits-only for others except email)
    {
      const upperTR = (s)=>{ try{ return String(s||'').toLocaleUpperCase('tr-TR'); }catch{ return String(s||'').toUpperCase(); } };
      const letters = (s)=> String(s||'').replace(/[^A-Za-zÇĞİÖŞÜçğıöşü\s]/g,' ').replace(/\s+/g,' ').trim();
      first_name = upperTR(letters(first_name));
      last_name = upperTR(letters(last_name));
      father_name = upperTR(letters(father_name));
      mother_name = upperTR(letters(mother_name));
      birth_place = upperTR(letters(birth_place));
      institution_name = upperTR(letters(institution_name));
      work_unit = upperTR(letters(work_unit));
      title = upperTR(letters(title));
      const digits = (s, maxLen)=>{ let v=String(s||'').replace(/\D/g,''); return (maxLen? v.slice(0,maxLen): v); };
      national_id = digits(national_id, 11);
      corp_reg_no = digits(corp_reg_no);
      retirement_no = digits(retirement_no);
      ssk_no = digits(ssk_no);
      phone = digits(phone, 10);
    }

    // Required validations
    const req = [first_name,last_name,national_id,father_name,mother_name,birth_place,birth_date,gender,education,work_province_val,work_district_val,institution_name,work_unit,corp_reg_no,title,blood_type,email,phone];
    if (req.some(v => !v)) { alert('Lütfen tüm zorunlu alanları doldurunuz.'); return false; }
    if (!api.isValidTCKN(national_id)) { alert('TC Kimlik No geçersizdir. Lütfen kontrol ediniz.'); return false; }
    // One of retirement_no or ssk_no must be filled
    if (!retirement_no && !ssk_no){ alert('Emekli Sandığı Sicil No veya SSK No alanlarından en az biri dolu olmalıdır.'); return false; }
    // Email normalize and validate
    try{ email = email.normalize('NFKC').replace(/[\s\u00A0\u200B-\u200D\uFEFF]/g,'').toLowerCase(); }catch{}
    if (!/^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(email)){ alert('Geçerli bir e‑posta giriniz.'); return false; }
    // Phone normalize to +90XXXXXXXXXX (10 digits)
    {
      let d = String(phone||'').replace(/\D/g,'');
      if (d.startsWith('90')) d = d.slice(2);
      if (d.startsWith('0')) d = d.slice(1);
      d = d.slice(0,10);
      if (d.length !== 10){ alert('Telefon 10 haneli olmalıdır (5XX XXX XX XX).'); return false; }
      phone = '+90' + d;
    }
    if (kvkk && !kvkk.checked){ alert('Lütfen KVKK Aydınlatma Metni\'ni onaylayınız.'); return false; }

    // reCAPTCHA validation
    {
      const hasSiteKey = !!(window.IlkeSendikaRecaptcha && window.IlkeSendikaRecaptcha.siteKey);
      if (hasSiteKey && window.grecaptcha){
        try{
          const wid = (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.widgetId);
          const resp = (wid != null) ? window.grecaptcha.getResponse(wid) : window.grecaptcha.getResponse();
          (window.IlkeSendika || {}).recaptcha && (window.IlkeSendika.recaptcha.token = resp);
        }catch{}
        const token = (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.token) || '';
        if (!token){ alert('Lütfen reCAPTCHA doğrulamasını tamamlayınız.'); return false; }
      } else {
        const fallback = document.getElementById('captchaAgree');
        if (!fallback || !fallback.checked){ alert('Lütfen reCAPTCHA doğrulamasını tamamlayınız.'); return false; }
      }
    }

    // Optional uploads
    let photo_url = null;
    let documents_urls = [];
    try{
      if (!window.ilkeSupabase) throw new Error('Veritabanı bağlantısı yok');
      const photoInput = form.querySelector('input[name="photo"]');
      const docsInput = form.querySelector('input[name="documents"]');
      if (photoInput && photoInput.files && photoInput.files[0]){
        const f = photoInput.files[0];
        const ext = (f.name.split('.').pop()||'jpg').toLowerCase();
        const key = `apply/${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
        const { error: upErr } = await window.ilkeSupabase.storage.from('member-photos').upload(key, f, { cacheControl:'3600', upsert:false, contentType: f.type||'application/octet-stream' });
        if (!upErr){ const r = window.ilkeSupabase.storage.from('member-photos').getPublicUrl(key); photo_url = r?.data?.publicUrl || null; }
      }
      if (docsInput && docsInput.files && docsInput.files.length){
        for (let i=0; i<docsInput.files.length; i++){
          const f = docsInput.files[i];
          const ext = (f.name.split('.').pop()||'dat').toLowerCase();
          const key = `apply/${Date.now()}_${i}_${Math.random().toString(36).slice(2)}.${ext}`;
          const { error: upErr } = await window.ilkeSupabase.storage.from('member-docs').upload(key, f, { cacheControl:'3600', upsert:false, contentType: f.type||'application/octet-stream' });
          if (!upErr){ const r = window.ilkeSupabase.storage.from('member-docs').getPublicUrl(key); const url = r?.data?.publicUrl || null; if (url) documents_urls.push(url); }
        }
      }
    }catch(uploadErr){ console.warn('Optional uploads failed:', uploadErr); }

    // Resolve province/district NAMES from selected options for storage
    let work_province = null, work_district = null;
    try{
      work_province = provSel && provSel.value ? (provSel.selectedOptions[0]?.textContent || null) : null;
      work_district = distSel && distSel.value ? (distSel.selectedOptions[0]?.textContent || null) : null;
    }catch{}

    // Insert member via secure Edge Function (pending: member_no null)
    try{
      if (!window.ilkeSupabase){ throw new Error('Veritabanı bağlantısı yok'); }
      const payload = {
        first_name, last_name, national_id, father_name, mother_name,
        birth_place, birth_date, gender, education,
        work_province, work_district, institution_name, work_unit,
        work_unit_address,
        corp_reg_no, title, blood_type, retirement_no, ssk_no,
        email, phone,
        status: 'pending',
        join_date: null,
        leave_date: null,
        photo_url: photo_url || null,
        documents_urls: documents_urls.length ? JSON.stringify(documents_urls) : JSON.stringify([])
      };
      // Call Edge Function apply_membership
      const token = (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.token) || null;
      const { data: fxData, error: fxError } = await window.ilkeSupabase.functions.invoke('apply_membership', {
        body: { payload, captcha_token: token }
      });
      if (fxError) {
        let ctxText = '';
        let parsed;
        try{
          const ctx = fxError && fxError.context;
          if (ctx && typeof ctx.text === 'function'){
            ctxText = await ctx.text();
          } else if (typeof ctx === 'string'){
            ctxText = ctx;
          } else if (ctx && typeof ctx.body === 'string'){
            ctxText = ctx.body;
          } else {
            ctxText = fxError.message || '';
          }
        }catch{}
        try{ parsed = ctxText ? JSON.parse(ctxText) : null; }catch{}
        // Friendly messages
        let userMsg = 'Başvurunuz kaydedilemedi. Lütfen doğrulamayı kontrol edip tekrar deneyiniz.';
        const errCode = parsed && parsed.error;
        if (errCode === 'missing_captcha_token') userMsg = 'Doğrulama bulunamadı. Lütfen sayfayı yenileyip tekrar deneyiniz.';
        else if (errCode === 'captcha_failed') userMsg = 'Doğrulama başarısız. Lütfen reCAPTCHA adımını tekrar yapınız.';
        else if (errCode === 'server_misconfigured') userMsg = 'Sunucu yapılandırması eksik. Lütfen daha sonra tekrar deneyiniz.';
        else if (errCode === 'insert_failed') userMsg = 'Başvurunuz şu anda alınamadı. Lütfen daha sonra tekrar deneyiniz.';
        console.error('apply_membership failed:', fxError, 'contextText:', ctxText, 'parsed:', parsed);
        alert(userMsg);
        return false;
      }
      if (fxData && fxData.error){
        console.error('apply_membership error payload:', fxData);
        alert('Başvurunuz kaydedilemedi. Lütfen doğrulamayı kontrol edip tekrar deneyiniz.');
        return false;
      }

      // Also create a message entry for admin "Mesajlar > Üyelik Başvurusu"
      try{
        const subj = `[Üyelik Başvurusu] ${first_name} ${last_name} - ${national_id}`;
        const body = [
          `Ad: ${first_name}`,
          `Soyad: ${last_name}`,
          `TC: ${national_id}`,
          `Telefon: ${phone}`,
          `E-posta: ${email}`,
          `İl: ${work_province || '-'}`,
          `İlçe: ${work_district || '-'}`,
          `Kurum: ${institution_name || '-'}`,
          `Birim: ${work_unit || '-'}`,
          `Kurum Sicil No: ${corp_reg_no || '-'}`,
          `Unvan: ${title || '-'}`,
          `Kan Grubu: ${blood_type || '-'}`,
          (retirement_no ? `Emekli Sicil No: ${retirement_no}` : null),
          (ssk_no ? `SSK No: ${ssk_no}` : null),
          '',
          'Açıklama:',
          'Site üzerinden yeni üyelik başvurusu oluşturuldu.'
        ].filter(Boolean).join('\n');
        const msgPayload = {
          name: `${first_name} ${last_name}`.trim(),
          email,
          subject: subj,
          body,
          category: 'membership',
          captcha_token: (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.token) || null,
          created_at: new Date().toISOString()
        };
        await window.ilkeSupabase.from('messages').insert(msgPayload);
      }catch{ /* best-effort */ }

      alert('Üyelik başvuru talebiniz alınmıştır. En kısa sürede tarafınıza dönüş yapılacaktır.');
      form.reset();
      try{ if (window.grecaptcha) window.grecaptcha.reset(); }catch{}
      return false;
    }catch(err){
      console.error('Başvuru kaydı başarısız:', err);
      alert('Başvurunuz kaydedilemedi. Lütfen daha sonra tekrar deneyiniz.');
      return false;
    }
  };

  // Meeting Request form handler
  api.handleMeetingRequestSubmit = async function(e){
    e.preventDefault();
    const form = e.target;
    const fd = new FormData(form);
    const fullName = (fd.get('name') || '').toString().trim();
    const phone = (fd.get('phone') || '').toString().trim();
    const email = (fd.get('email') || '').toString().trim();
    const company = (fd.get('company') || '').toString().trim();
    const subject = (fd.get('subject') || '').toString().trim();
    const message = (fd.get('message') || '').toString().trim();

    // KVKK checkbox must be checked
    const kvkk = form.querySelector('input[name="kvkk"]');

    // Basic validation
    if (!fullName || !phone || !email || !company || !subject || !message){
      alert('Lütfen formdaki zorunlu alanları doldurunuz.');
      return false;
    }
    if (kvkk && !kvkk.checked){
      alert('Lütfen KVKK Aydınlatma Metni\'ni onaylayınız.');
      return false;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){
      alert('Lütfen geçerli bir e-posta adresi giriniz.');
      return false;
    }

    // reCAPTCHA validation (same container ids as contact form)
    {
      const hasSiteKey = !!(window.IlkeSendikaRecaptcha && window.IlkeSendikaRecaptcha.siteKey);
      if (hasSiteKey && window.grecaptcha){
        try{
          const wid = (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.widgetId);
          const resp = (wid != null) ? window.grecaptcha.getResponse(wid) : window.grecaptcha.getResponse();
          (window.IlkeSendika || {}).recaptcha && (window.IlkeSendika.recaptcha.token = resp);
        }catch{}
        const token = (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.token) || '';
        if (!token){ alert('Lütfen doğrulamayı tamamlayınız.'); return false; }
      } else {
        const fallback = document.getElementById('captchaAgree');
        if (!fallback || !fallback.checked){
          alert('Lütfen robot olmadığınızı onaylayınız.');
          return false;
        }
      }
    }

    // Persist to Supabase messages table
    try{
      if (!window.ilkeSupabase){ throw new Error('Veritabanı bağlantısı yok'); }
      const payload = {
        name: fullName,
        email,
        subject: subject,
        body: [
          `Ad Soyad: ${fullName}`,
          `Telefon: ${phone}`,
          `Kurum: ${company || '-'}`,
          '',
          'Açıklama:',
          message
        ].filter(Boolean).join('\n'),
        category: 'meeting',
        captcha_token: (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.token) || null,
        created_at: new Date().toISOString()
      };
      const { error } = await window.ilkeSupabase.from('messages').insert(payload);
      if (error) throw error;

      // Optional: trigger email to admins via existing function (best-effort)
      try{
        await window.ilkeSupabase.functions.invoke('send_contact', {
          body: {
            name: fullName,
            email,
            subject: subject,
            message: payload.body,
            captcha_token: payload.captcha_token
          }
        });
      }catch{}

      alert('Görüşme talebiniz başarıyla oluşturulmuştur. En kısa sürede tarafınıza dönüş sağlanacaktır.');
      form.reset();
      try{ if (window.grecaptcha) window.grecaptcha.reset(); }catch{}
      return false;
    }catch(err){
      console.error('Talep kaydı başarısız:', err);
      alert('Talebiniz kaydedilemedi. Lütfen daha sonra tekrar deneyiniz.');
      return false;
    }
  };

  // Expose shared reCAPTCHA state so global initRecaptcha() can write token/widgetId
  api.recaptcha = recaptcha;

  // Lawyer Request form handler
  api.handleLawyerRequestSubmit = async function(e){
    e.preventDefault();
    const form = e.target;
    const fd = new FormData(form);
    const name = (fd.get('name') || '').toString().trim();
    const surname = (fd.get('surname') || '').toString().trim();
    const memberNo = (fd.get('member_no') || '').toString().trim();
    const phone = (fd.get('phone') || '').toString().trim();
    const email = (fd.get('email') || '').toString().trim();
    const company = (fd.get('company') || '').toString().trim();
    const subject = (fd.get('subject') || '').toString().trim();
    const message = (fd.get('message') || '').toString().trim();

    // Basic validation
    if (!name || !surname || !memberNo || !phone || !email || !subject || !message){
      alert('Lütfen formdaki zorunlu alanları doldurunuz.');
      return false;
    }
    // Üye No sayısal olmalıdır
    if (!/^\d{1,20}$/.test(memberNo)){
      alert('Üye No yalnızca rakamlardan oluşmalıdır.');
      return false;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){
      alert('Lütfen geçerli bir e-posta adresi giriniz.');
      return false;
    }

    // reCAPTCHA validation (same container ids as contact form)
    {
      const hasSiteKey = !!(window.IlkeSendikaRecaptcha && window.IlkeSendikaRecaptcha.siteKey);
      if (hasSiteKey && window.grecaptcha){
        try{
          const wid = (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.widgetId);
          const resp = (wid != null) ? window.grecaptcha.getResponse(wid) : window.grecaptcha.getResponse();
          (window.IlkeSendika || {}).recaptcha && (window.IlkeSendika.recaptcha.token = resp);
        }catch{}
        const token = (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.token) || '';
        if (!token){ alert('Lütfen doğrulamayı tamamlayınız.'); return false; }
      } else {
        const fallback = document.getElementById('captchaAgree');
        if (!fallback || !fallback.checked){
          alert('Lütfen robot olmadığınızı onaylayınız.');
          return false;
        }
      }
    }

    // Persist to Supabase messages table
    try{
      if (!window.ilkeSupabase){ throw new Error('Veritabanı bağlantısı yok'); }
      const payload = {
        name: `${name} ${surname}`.trim(),
        email,
        subject: subject,
        body: [
          `Ad: ${name}`,
          `Soyad: ${surname}`,
          (memberNo ? `Üye No: ${memberNo}` : null),
          `Telefon: ${phone}`,
          `Kurum: ${company || '-'}`,
          '',
          'Açıklama:',
          message
        ].filter(Boolean).join('\n'),
        category: 'lawyer',
        captcha_token: (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.token) || null,
        created_at: new Date().toISOString()
      };
      const { error } = await window.ilkeSupabase.from('messages').insert(payload);
      if (error) throw error;

      // Optional: trigger email to admins via existing function (best-effort)
      try{
        await window.ilkeSupabase.functions.invoke('send_contact', {
          body: {
            name: `${name} ${surname}`.trim(),
            email,
            subject: subject,
            message: payload.body,
            captcha_token: payload.captcha_token
          }
        });
      }catch{}

      alert('Avukatlık talebiniz başarıyla oluşturulmuştur. Sendika avukatımız tarafınan tarafınıza dönüş sağlanacaktır.');
      form.reset();
      try{ if (window.grecaptcha) window.grecaptcha.reset(); }catch{}
      return false;
    }catch(err){
      console.error('Talep kaydı başarısız:', err);
      alert('Talebiniz kaydedilemedi. Lütfen daha sonra tekrar deneyiniz.');
      return false;
    }
  };

  function qs(sel, root=document){ return root.querySelector(sel); }
  function qsa(sel, root=document){ return Array.from(root.querySelectorAll(sel)); }

  function initNav() {
  const navToggle = document.querySelector('.nav-toggle');
  const navMenu = document.getElementById('navMenu');
  if (!navToggle || !navMenu) return;
  // Ana menüyü (hamburger) aç/kapat
  navToggle.addEventListener('click', () => {
    const isOpen = navMenu.classList.toggle('open');
    navToggle.setAttribute('aria-expanded', String(isOpen));
  });
  // Menü içindeki tüm link tıklamalarını yönet
  document.querySelectorAll('#navMenu a').forEach(anchor => {
    anchor.addEventListener('click', (event) => {
      // Bu mantık sadece mobil ekranlar için geçerlidir
      if (window.innerWidth > 768) {
        return;
      }
      const parentLi = anchor.closest('li');
      if (!parentLi) return;
      // Eğer tıklanan linkin ebeveyni bir alt menü kabı ise ("Kurumsal" gibi)
      if (parentLi.classList.contains('has-sub')) {
        // Linkin normalde gideceği yere gitmesini engelle
        event.preventDefault();
        // ve alt menüyü aç/kapat.
        const isOpen = parentLi.classList.toggle('open');
        anchor.setAttribute('aria-expanded', String(isOpen));
      } else {
        // Diğer tüm normal linklere ("Anasayfa", "Vizyon" vb.) tıklandığında
        // ana menüyü kapat. Link normal şekilde çalışmaya devam edecek.
        navMenu.classList.remove('open');
        navToggle.setAttribute('aria-expanded', 'false');
      }
    });
  });
}
 
  // Ensure menu items exist and correct ordering across all pages
  function ensureMenuStructure(){
    try{
      const nav = document.getElementById('navMenu');
      if (!nav) return;
      const items = Array.from(nav.children).filter(el => el.matches && el.matches('li.has-sub'));
      // Helper to match top-level by anchor text
      const findTop = (title) => items.find(li => {
        const a = li.firstElementChild && li.firstElementChild.tagName === 'A' ? li.firstElementChild : li.querySelector('a');
        return a && (a.textContent||'').trim().toLowerCase() === title;
      });

      // 1) Anasayfa submenu: add "Amaç" after "Misyon"
      const homeLi = findTop('anasayfa');
      if (homeLi){
        const submenu = (homeLi.querySelector('.submenu') || Array.from(homeLi.children).find(c=>c.classList && c.classList.contains('submenu')));
        if (submenu){
          const lis = Array.from(submenu.children).filter(n => n.tagName === 'LI');
          const anchors = Array.from(submenu.querySelectorAll('a'));
          const hasAmac = anchors.some(a=>{
            const h=(a.getAttribute('href')||'').toLowerCase();
            return /slug=amac/.test(h) || /\/(?:page\/)?amac(?:[\/#?]|$)/.test(h);
          });
          const misyonLi = lis.find(li => {
            const a = li.querySelector('a'); const h=(a && a.getAttribute('href')||'').toLowerCase();
            return /slug=misyon/.test(h) || /\/(?:page\/)?misyon(?:[\/#?]|$)/.test(h);
          });
          if (!hasAmac){
            // Create Amaç item
            const li = document.createElement('li');
            const a = document.createElement('a'); a.href = '/amac'; a.textContent = 'Amaç';
            li.appendChild(a);
            if (misyonLi && misyonLi.nextSibling){
              submenu.insertBefore(li, misyonLi.nextSibling);
            } else if (misyonLi){
              submenu.appendChild(li);
            } else {
              // Fallback: append at end
              submenu.appendChild(li);
            }
          }
        }
      }

      // 2) Kurumsal submenu: move Genel Başkan to top
      const corpLi = findTop('kurumsal');
      if (corpLi){
        const submenu = (corpLi.querySelector('.submenu') || Array.from(corpLi.children).find(c=>c.classList && c.classList.contains('submenu')));
        if (submenu){
          const gbLi = Array.from(submenu.children).find(li => {
            const a = li.querySelector('a'); const h=(a && a.getAttribute('href')||'').toLowerCase();
            return /slug=genel-baskan/.test(h) || /\/(?:page\/)?genel-baskan(?:[\/#?]|$)/.test(h);
          });
          if (gbLi && submenu.firstElementChild !== gbLi){
            submenu.insertBefore(gbLi, submenu.firstElementChild);
          }
        }
      }

      // 3) Normalize top-level "İletişim" link to contact.html on all pages (handle Turkish casing/diacritics)
      try{
        const links = Array.from(nav.querySelectorAll('a'));
        const norm = (s)=> String(s||'')
          .toLowerCase()
          .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
          .replace(/ı/g,'i').replace(/İ/g,'i')
          .trim();
        links.forEach(a=>{
          const t = norm(a.textContent);
          if (t === 'iletisim' || t.includes('iletisim')){
            const href = (a.getAttribute('href') || '').trim();
            if (!/contact\.html(?:$|[?#])/i.test(href)){
              a.setAttribute('href','contact.html');
            }
          }
        });
      }catch{}

      // 4) Ensure "İletişim" has submenu items (İletişim, Avukatlık Talebi, Görüşme Talebi, Üyelik Başvurusu)
      try{
        const norm = (s)=> String(s||'')
          .toLowerCase()
          .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
          .replace(/ı/g,'i').replace(/İ/g,'i')
          .trim();
        const contactAnchor = Array.from(nav.querySelectorAll('a')).find(a => {
          const t = norm(a.textContent);
          return t === 'iletisim' || t.includes('iletisim');
        });
        if (contactAnchor){
          const li = contactAnchor.closest('li');
          if (li){
            li.classList.add('has-sub');
            // Turn top-level into a toggle on mobile; real link lives in submenu as first item
            contactAnchor.setAttribute('href','#');
            let sub = li.querySelector('ul.submenu');
            if (!sub){ sub = document.createElement('ul'); sub.className='submenu'; li.appendChild(sub); }
            const ensureSub = (title, href) => {
              const exists = Array.from(sub.querySelectorAll('a')).some(a => norm(a.textContent) === norm(title));
              if (!exists){ const li2=document.createElement('li'); const a=document.createElement('a'); a.textContent=title; a.href=href; li2.appendChild(a); sub.appendChild(li2); }
            };
            ensureSub('İletişim', 'contact.html');
            ensureSub('Avukatlık Talebi', 'avukatlik-talebi.html');
            ensureSub('Görüşme Talebi', 'gorusme-talebi.html');
            ensureSub('Üyelik Başvurusu', 'uyelik-basvurusu.html');
          }
        }
      }catch{}
    }catch(e){ /* no-op */ }
  }
  // Expose for other modules/IIFEs
  window.sanitizeInlineHtml = sanitizeInlineHtml;

  // For homepage slider band and news cards: allow only tiny inline HTML
  function sanitizeInlineHtml(dirty){
    try{
      const allowedTags = new Set(['B','STRONG','I','EM','U','SPAN','A','BR','SMALL']);
      const allowedAttrs = {
        'A': new Set(['href','target','rel']),
        'SPAN': new Set(['style','class']),
        'SMALL': new Set([])
      };
      const allowedStyles = new Set(['font-weight','font-style','text-decoration','font-size','font-family']);
      const tmp = document.createElement('div'); tmp.innerHTML = dirty || '';
      const walker = document.createTreeWalker(tmp, NodeFilter.SHOW_ELEMENT, null);
      const toRemove = [];
      function cleanStyle(el){
        const style = el.getAttribute('style')||''; if (!style) return;
        const parts = style.split(';').map(s=>s.trim()).filter(Boolean);
        const kept = [];
        for (const part of parts){
          const [prop, valRaw] = part.split(':'); if (!prop || !valRaw) continue;
          const p = prop.trim().toLowerCase(); const v = valRaw.trim();
          if (!allowedStyles.has(p)) continue;
          if (p === 'font-size' && !/^\d{1,3}px$/.test(v)) continue;
          if (p === 'font-family' && /[;{}]/.test(v)) continue;
          kept.push(`${p}:${v}`);
        }
        if (kept.length) el.setAttribute('style', kept.join('; ')); else el.removeAttribute('style');
      }
      while (walker.nextNode()){
        const el = walker.currentNode;
        if (!allowedTags.has(el.tagName)) { toRemove.push(el); continue; }
        Array.from(el.attributes).forEach(attr=>{
          if (!allowedAttrs[el.tagName] || !allowedAttrs[el.tagName].has(attr.name)) el.removeAttribute(attr.name);
        });
        cleanStyle(el);
        if (el.tagName === 'A'){
          const href = el.getAttribute('href')||'';
          if (!/^https?:\/\//i.test(href)) el.removeAttribute('href');
          el.setAttribute('rel','noopener noreferrer'); el.setAttribute('target','_blank');
        }
      }
      toRemove.forEach(n => n.replaceWith(document.createTextNode(n.textContent||'')));
      return tmp.innerHTML;
    }catch{ return String(dirty||''); }
  }

  // Ensure "Konu" field exists on contact page even if cached HTML lacks it
  function ensureContactSubjectField(){
    try{
      const form = document.querySelector('.contact-form');
      if (!form) return;
      const hasSubject = !!form.querySelector('input[name="subject"]');
      if (hasSubject) return;
      const emailInput = form.querySelector('input[name="email"]');
      const emailLabel = emailInput && emailInput.closest ? emailInput.closest('label') : null;
      const messageTextarea = form.querySelector('textarea[name="message"]');
      const messageLabel = messageTextarea && messageTextarea.closest ? messageTextarea.closest('label') : null;
      const lbl = document.createElement('label');
      lbl.innerHTML = '<span>Konu</span>';
      const input = document.createElement('input'); input.type='text'; input.name='subject'; input.required = true; input.id = 'contactSubject';
      lbl.appendChild(input);
      if (messageLabel && messageLabel.parentNode){ messageLabel.parentNode.insertBefore(lbl, messageLabel); }
      else if (emailLabel && emailLabel.parentNode){ emailLabel.parentNode.insertBefore(lbl, emailLabel.nextSibling); }
    }catch(e){ console.warn('ensureContactSubjectField error:', e); }
  }
  // expose for console/debug
  window.ensureContactSubjectField = ensureContactSubjectField;

  // Minimal sanitizer for rendering trusted admin content safely
  function sanitizePolicyHtml(dirty){
    try{
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
        const style = el.getAttribute('style')||''; if (!style) return;
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
      const tmp = document.createElement('div'); tmp.innerHTML = dirty || '';
      const walker = document.createTreeWalker(tmp, NodeFilter.SHOW_ELEMENT, null);
      const toRemove = [];
      while (walker.nextNode()){
        const el = walker.currentNode;
        if (!allowedTags.has(el.tagName)) { toRemove.push(el); continue; }
        Array.from(el.attributes).forEach(attr => {
          if (!allowedAttrs[el.tagName] || !allowedAttrs[el.tagName].has(attr.name)) el.removeAttribute(attr.name);
        });
        cleanStyle(el);
        if (el.tagName === 'A'){
          const href = el.getAttribute('href')||''; if (!/^https?:\/\//i.test(href)) el.removeAttribute('href');
          el.setAttribute('rel','noopener noreferrer'); el.setAttribute('target','_blank');
        }
      }
      toRemove.forEach(n => n.replaceWith(document.createTextNode(n.textContent||'')));
      return tmp.innerHTML;
    }catch{ return String(dirty||''); }
  }

  function initPolicyModal(){
    const modal = document.getElementById('policyModal');
    if (!modal) return; // only on pages that have the modal
    const titleEl = document.getElementById('policyTitle');
    const contentEl = document.getElementById('policyContent');
    const closeBtn = document.getElementById('policyClose');
    const linkPrivacy = document.getElementById('privacyLink');
    const linkDisclosure = document.getElementById('disclosureLink');
    const linkDisclosure2 = document.getElementById('kvkkDisclosureLink');

    function hide(){ modal.style.display = 'none'; document.removeEventListener('keydown', onEsc); }
    function onEsc(e){ if (e.key === 'Escape') hide(); }
    closeBtn && closeBtn.addEventListener('click', hide);

    async function openPolicy(key, title){
      try{
        titleEl && (titleEl.textContent = title || 'Bilgi');
        contentEl && (contentEl.innerHTML = '<div class="muted">Yükleniyor...</div>');
        if (window.IlkeSendikaReload && typeof window.IlkeSendikaReload.getPolicy === 'function'){
          const txt = await window.IlkeSendikaReload.getPolicy(key);
          contentEl.innerHTML = sanitizePolicyHtml(txt||'');
          // Trim trailing blank nodes (<br>, empty <p>, &nbsp;) to remove extra bottom spacing
          (function cleanup(){
            if (!contentEl) return;
            const isEmpty = (node) => {
              if (!node) return true;
              if (node.nodeType === Node.TEXT_NODE) return !node.textContent.replace(/\u00a0/g,' ').trim();
              if (node.nodeType === Node.ELEMENT_NODE){
                if (node.tagName === 'BR') return true;
                const html = (node.innerHTML || '').replace(/&nbsp;/g,'').replace(/<br\s*\/?>/ig,'').trim();
                const text = (node.textContent || '').replace(/\u00a0/g,' ').trim();
                return !html || !text;
              }
              return false;
            };
            let n = contentEl.lastChild;
            while (n && isEmpty(n)) { const prev = n.previousSibling; n.remove(); n = prev; }
            if (contentEl.lastElementChild){ contentEl.lastElementChild.style.marginBottom='0'; contentEl.lastElementChild.style.paddingBottom='0'; }
          })();
        }
        modal.style.display = 'block';
        document.addEventListener('keydown', onEsc);
      } catch(e){
        console.warn('policy modal error:', e);
      }
    }

    function maybeBind(linkEl, key, title){
      if (!linkEl) return;
      // If link navigates to page.html, do not intercept; modal is only for in-place display legacy
      const href = (linkEl.getAttribute('href')||'').toLowerCase();
      if (href.includes('page.html')) return;
      linkEl.addEventListener('click', (e)=>{ e.preventDefault(); openPolicy(key, title); });
    }
    maybeBind(linkPrivacy, 'policy_privacy', 'Gizlilik Politikası');
    maybeBind(linkDisclosure, 'policy_disclosure', 'Aydınlatma Metni');
    maybeBind(linkDisclosure2, 'policy_disclosure', 'Aydınlatma Metni');
  }

  // (no data-access helpers here; defined in Supabase IIFE below)

  // (duplicate removed)

    function initSlider(){
    const slider = qs('[data-slider]');
    if (!slider) return;
    const track = qs('[data-slides]', slider);
    const prev = qs('[data-prev]', slider);
    const next = qs('[data-next]', slider);
    if (!track) return;

    let index = 0;

    function slides(){ return qsa('.slide', track); }

    function ensureIndex(){
      const len = slides().length;
      if (len === 0) return 0;
      if (index >= len) index = len - 1;
      if (index < 0) index = 0;
      return len;
    }

    function update(){
      const len = ensureIndex();
      if (len === 0) return;
      const offset = -index * 100;
      track.style.transform = `translateX(${offset}%)`;
      slides().forEach((s, i) => s.classList.toggle('current', i === index));
    }

    function go(delta){
      const len = slides().length;
      if (len === 0) return;
      index = (index + delta + len) % len;
      update();
    }

    prev && prev.addEventListener('click', () => go(-1));
    next && next.addEventListener('click', () => go(1));
    // Delegate clicks in case buttons are re-rendered
    slider.addEventListener('click', (ev) => {
      const t = ev.target;
      if (!t) return;
      if (t.closest && t.closest('[data-prev]')) { ev.preventDefault(); go(-1); }
      else if (t.closest && t.closest('[data-next]')) { ev.preventDefault(); go(1); }
    });

    // Auto-rotate every 6s
    let timer = setInterval(() => go(1), 6000);
    slider.addEventListener('mouseenter', () => clearInterval(timer));
    slider.addEventListener('mouseleave', () => { timer = setInterval(() => go(1), 6000); });

    // If slides load asynchronously, observe and update when children change
    const obs = new MutationObserver(() => { index = 0; update(); });
    obs.observe(track, { childList: true });
    // Also respond to a custom event if loaders dispatch it
    document.addEventListener('slider:updated', () => { index = 0; update(); }, { once:false });

    // Initial update (may be a no-op if no slides yet)
    update();
  }

  api.handleContactSubmit = async function(e){
    e.preventDefault();
    const form = e.target;
    const fd = new FormData(form);
    const name = (fd.get('name') || '').toString().trim();
    const email = (fd.get('email') || '').toString().trim();
    const subject = (fd.get('subject') || '').toString().trim();
    const message = (fd.get('message') || '').toString().trim();

    // Basic validation
    if (!name || !email || !subject || !message){
      alert('Lütfen formdaki tüm alanları doldurunuz.');
      return false;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){
      alert('Lütfen geçerli bir e-posta adresi giriniz.');
      return false;
    }

    // reCAPTCHA validation
    const isContactPage = document.body && document.body.getAttribute('data-page') === 'contact';
    if (isContactPage){
      const hasSiteKey = !!(window.IlkeSendikaRecaptcha && window.IlkeSendikaRecaptcha.siteKey);
      if (hasSiteKey && window.grecaptcha){
        // Ensure token exists (use specific widgetId when present)
        try{
          const wid = (window.IlkeSendika && window.IlkeSendika.recaptcha && window.IlkeSendika.recaptcha.widgetId);
          const resp = (wid != null) ? window.grecaptcha.getResponse(wid) : window.grecaptcha.getResponse();
          recaptcha.token = resp;
        }catch{}
        if (!recaptcha.token){
          alert('Lütfen doğrulamayı tamamlayınız.');
          return false;
        }
      } else {
        const fallback = document.getElementById('captchaAgree');
        if (!fallback || !fallback.checked){
          alert('Lütfen robot olmadığınızı onaylayınız.');
          return false;
        }
      }
    }

// Persist to Supabase messages table
try{
  if (!window.ilkeSupabase){ throw new Error('Veritabanı bağlantısı yok'); }
  const payload = {
    name,
    email,
    subject,
    body: message,
    category: 'contact',
    captcha_token: recaptcha.token || null,
    created_at: new Date().toISOString()
  };
  const { error } = await window.ilkeSupabase.from('messages').insert(payload);
  if (error) throw error;

  // Trigger email via Edge Function (non-blocking)
  try {
    await window.ilkeSupabase.functions.invoke('send_contact', {
      body: { name, email, subject, message, captcha_token: recaptcha.token || null }
    });
  } catch {}

  alert('Teşekkürler! Mesajınız alınmıştır. En kısa sürede dönüş yapacağız.');
  form.reset();
  try{ if (window.grecaptcha) window.grecaptcha.reset(); }catch{}
  return false;
}catch(err){
  console.error('Mesaj kaydı başarısız:', err);
  alert('Mesajınız kaydedilemedi. Lütfen daha sonra tekrar deneyiniz.');
  return false;
}
  };

  function initYear(){
    const el = qs('#year');
    if (el) el.textContent = new Date().getFullYear();
  }

  function initHeaderScrollShadow(){
    const header = document.querySelector('.site-header');
    if (!header) return;
    const onScroll = () => {
      if ((window.scrollY || document.documentElement.scrollTop || 0) > 8) header.classList.add('scrolled');
      else header.classList.remove('scrolled');
    };
    onScroll();
    window.addEventListener('scroll', onScroll, { passive: true });
  }

  function devPatchPrettyLinks(){
    try{
      const isLocal = /^(localhost|127\.0\.0\.1)$/i.test(location.hostname);
      if (!isLocal) return;
      const re = /^\/(?:page\/)?([a-z0-9-]+)(?:\/?|[?#].*)$/i;
      document.addEventListener('click', (ev)=>{
        const t = ev.target && ev.target.closest ? ev.target.closest('a') : null;
        if (!t) return;
        const href = t.getAttribute('href') || '';
        if (!href || href.startsWith('http') || href.startsWith('mailto:') || href.startsWith('#')) return;
        const url = new URL(href, location.origin);
        if (url.origin !== location.origin) return;
        const m = url.pathname.match(re);
        if (!m) return;
        ev.preventDefault();
        const slug = m[1].toLowerCase();
        const v = '20260101-02';
        location.href = `page.html?v=${v}&slug=${encodeURIComponent(slug)}`;
      }, true);
    }catch{}
  }

  function init(){
    initNav();
    ensureMenuStructure();
    devPatchPrettyLinks();
    initSlider();
    initYear();
    initPolicyModal();
    initHeaderScrollShadow();
    if (window.initReveal) window.initReveal();
    if (window.initAnnouncementsTicker) window.initAnnouncementsTicker();

    // Enforce subject field and reCAPTCHA on contact page
    if (document.body && document.body.getAttribute('data-page') === 'contact'){
      ensureContactSubjectField();
      // Watch the form in case it is re-rendered later
      try{
        const form = document.querySelector('.contact-form');
        if (form){
          const obs = new MutationObserver(()=> ensureContactSubjectField());
          obs.observe(form, { childList:true, subtree:true });
        }
      }catch{}
    }

    // Dynamic content from Supabase
    const callLoaders = () => {
      try{
        if (!window.ilkeSupabase) return false;
        if (!window.IlkeSendikaReload) return false;
        window.IlkeSendikaReload.loadSliderNews && window.IlkeSendikaReload.loadSliderNews();
        window.IlkeSendikaReload.loadNewsCards && window.IlkeSendikaReload.loadNewsCards();
        window.IlkeSendikaReload.loadAnnouncements && window.IlkeSendikaReload.loadAnnouncements();
        window.IlkeSendikaReload.loadFooterSettings && window.IlkeSendikaReload.loadFooterSettings();
        // Home posters area (if exists) — call without await (this function is not async)
        try{
          const hp = document.getElementById('homePosters');
          if (hp) { window.IlkeSendikaReload.loadHomePosters().catch(()=>{}); }
        }catch{}
        // Contact page extras
        if (document.body && document.body.getAttribute('data-page') === 'contact'){
          ensureContactSubjectField();
          try{
            const p = window.IlkeSendikaReload.loadContactSettings ? window.IlkeSendikaReload.loadContactSettings() : null;
            if (p && typeof p.then === 'function'){
              p.then(()=>{ initRecaptcha(); }).catch((e)=>{ console.warn('loadContactSettings failed:', e); initRecaptcha(); });
            } else {
              initRecaptcha();
            }
          }catch(e){ console.warn('contact init error:', e); initRecaptcha(); }
        } else if (document.body && (
          document.body.getAttribute('data-page') === 'lawyer-request' ||
          document.body.getAttribute('data-page') === 'avukatlik-talebi' ||
          document.body.getAttribute('data-page') === 'gorusme-talebi' ||
          document.body.getAttribute('data-page') === 'uyelik-basvurusu'
        )){
          // Lawyer/Membership request pages: init geo selects (if membership) then recaptcha
          try{ if (document.body.getAttribute('data-page') === 'uyelik-basvurusu'){ initPublicGeoSelectors(); initMembershipFormInputFilters(); } }catch{}
          try{
            const p = window.IlkeSendikaReload && window.IlkeSendikaReload.loadContactSettings ? window.IlkeSendikaReload.loadContactSettings() : null;
            if (p && typeof p.then === 'function'){
              p.then(()=>{ initRecaptcha(); }).catch((e)=>{ console.warn('loadContactSettings failed:', e); initRecaptcha(); });
            } else {
              initRecaptcha();
            }
          }catch(e){ console.warn('lawyer-request init error:', e); initRecaptcha(); }
        }
        return true;
      } catch(e){ console.error('Dynamic load error:', e); return false; }
    };
    if (!callLoaders()){
      let attempts = 0;
      const timer = setInterval(() => {
        attempts++;
        if (callLoaders() || attempts >= 50) clearInterval(timer);
      }, 200);
    }
  }

  // Re-apply menu normalization at strategic times to defeat stale markup
  window.addEventListener('load', () => {
    try { ensureMenuStructure(); } catch {}
    try {
      const nav = document.getElementById('navMenu');
      if (nav){
        // Re-run when user hovers the nav (useful if markup changed late)
        nav.addEventListener('mouseenter', () => { try { ensureMenuStructure(); } catch {} }, { once:false });
        // Observe DOM changes within nav and re-normalize
        const obs = new MutationObserver(() => { try { ensureMenuStructure(); } catch {} });
        obs.observe(nav, { childList: true, subtree: true });
      }
    } catch {}
  }, { once: true });

  document.addEventListener('DOMContentLoaded', init);
  return api;
})();

// ============ Contact Helpers: reCAPTCHA enable/disable ============
function initRecaptcha(){
  try{
    const container = document.getElementById('recaptcha-container');
    const submitBtn = document.getElementById('contactSubmit');
    const setSubmitEnabled = (v)=>{ if (submitBtn){ submitBtn.disabled = false; submitBtn.removeAttribute('aria-disabled'); } };
    // Default: disabled until verified
    setSubmitEnabled(true);
    if (!container){ setSubmitEnabled(true); return; }

    const key = (window.IlkeSendikaRecaptcha && window.IlkeSendikaRecaptcha.siteKey) || window.IlkeSendikaRecaptchaDefaultSiteKey || null;
    if (!key){
      // No site key: show fallback checkbox
      const fb = document.getElementById('captchaFallback'); if (fb) fb.hidden = false;
      const cb = document.getElementById('captchaAgree'); if (cb){ cb.addEventListener('change', ()=>{}); }
    }

    // We have a site key: ensure fallback stays hidden unless script fails
    { const fb = document.getElementById('captchaFallback'); if (fb) fb.hidden = true; }

    // Load Google reCAPTCHA script once
    const RENDER_CB = 'onRecaptchaReady';
    const scriptId = 'recaptcha-api';
    if (!document.getElementById(scriptId)){
      const s = document.createElement('script');
      s.id = scriptId;
      s.src = `https://www.google.com/recaptcha/api.js?onload=${RENDER_CB}&render=explicit`;
      s.async = true; s.defer = true;
      document.head.appendChild(s);
    }
    // If script fails to load within 6s, present fallback checkbox
    setTimeout(()=>{
      try{
        if (!window.grecaptcha || !window.grecaptcha.render){
          const fb = document.getElementById('captchaFallback'); if (fb) fb.hidden = false;
          const cb = document.getElementById('captchaAgree'); if (cb){ cb.addEventListener('change', ()=> setSubmitEnabled(!!cb.checked)); }
        }
      }catch{}
    }, 6000);
    // Define onload callback (idempotent)
    window[RENDER_CB] = function(){
      try{
        if (!window.grecaptcha) return;
        // Hide fallback if it was shown by timeout and proceed with widget
        { const fb = document.getElementById('captchaFallback'); if (fb) fb.hidden = true; }
        const widgetId = window.grecaptcha.render('recaptcha-container', {
          sitekey: key,
          callback: (token) => {
            try { (window.IlkeSendika || {}).recaptcha && (window.IlkeSendika.recaptcha.token = token); } catch {}
            // Token received -> ensure fallback stays hidden and enable submit
            { const fb = document.getElementById('captchaFallback'); if (fb) fb.hidden = true; }
            setSubmitEnabled(true);
          },
          'expired-callback': () => {
            try { (window.IlkeSendika || {}).recaptcha && (window.IlkeSendika.recaptcha.token = null); } catch {}
            setSubmitEnabled(true);
          }
        });
        try { (window.IlkeSendika || {}).recaptcha && ((window.IlkeSendika.recaptcha.widgetId = widgetId)); } catch {}
        // Do NOT auto-show fallback just because user hasn't clicked yet. Only show it when script/render is unavailable.
      } catch(e){ console.warn('recaptcha render error:', e); }
    };
  } catch(e){ console.warn('initRecaptcha error:', e); }
}

// ============ Supabase Dynamic Content ============
(function(){
  const db = () => window.ilkeSupabase;

  async function loadSliderNews(){
    if (!db()) return;
    const slidesRoot = document.getElementById('sliderSlides');
    if (!slidesRoot) return;
    try {
      // Prefer featured news, then latest
      const { data, error } = await db()
        .from('news')
        .select('id, title, summary, image_url, cover_image_url, published_at, unpublish_at, is_featured, status')
        .eq('status', 'published')
        .order('is_featured', { ascending: false })
        .order('published_at', { ascending: false })
        .limit(6);
      if (error) throw error;

      const now = Date.now();
      const items = (data || []).filter(n => {
        const u = n.unpublish_at ? new Date(n.unpublish_at).getTime() : null;
        return !u || u > now;
      });

      slidesRoot.innerHTML = '';
      // helper to split meta (key|z=...,ox=...,oy=...)
      function splitMetaUrl(u){
  const s = String(u||''); const i = s.indexOf('|');
  if (i < 0) return { url: s, meta: null };
  const base = s.slice(0, i); const tail = s.slice(i+1);
  const meta = { z:1, ox:0, oy:0, oxp:null, oyp:null, zm:null, oxm:null, oym:null, oxpm:null, oypm:null };
  tail.split(',').forEach(p=>{
    const [k,v] = p.split('='); const n = Number(v);
    if (k==='z' && isFinite(n)) meta.z=n;
    if (k==='ox' && isFinite(n)) meta.ox=Math.round(n);
    if (k==='oy' && isFinite(n)) meta.oy=Math.round(n);
    if (k==='oxp' && isFinite(n)) meta.oxp=n;
    if (k==='oyp' && isFinite(n)) meta.oyp=n;
    if (k==='zm' && isFinite(n)) meta.zm=n;
    if (k==='oxm' && isFinite(n)) meta.oxm=Math.round(n);
    if (k==='oym' && isFinite(n)) meta.oym=Math.round(n);
    if (k==='oxpm' && isFinite(n)) meta.oxpm=n;
    if (k==='oypm' && isFinite(n)) meta.oypm=n;
  });
  return { url: base, meta };
}
      items.forEach((n, idx) => {
         const art = document.createElement('article');
         art.className = 'slide' + (idx === 0 ? ' current' : '');
         art.setAttribute('role', 'listitem');
         // Image underlay (prefer cover with crop meta)
         const raw = n.cover_image_url || n.image_url || n.cover_image_url; // prefer cover
         if (raw) {
           const { url, meta } = splitMetaUrl(raw);
           const holder = document.createElement('div');
           holder.style.position = 'absolute'; holder.style.inset = '0'; holder.style.overflow = 'hidden'; holder.style.background = '#fff';
           const bg = document.createElement('div');
           bg.style.position = 'absolute'; bg.style.inset = '0'; bg.style.backgroundColor = '#fff';
           bg.style.backgroundImage = `url(${url})`;
           bg.style.backgroundRepeat = 'no-repeat';
           // Clamp to at least 100% to avoid letterboxing/black areas
           if (meta){
  const isMobile = window.matchMedia('(max-width: 960px)').matches;
  const z = Math.max(1, Number((isMobile && isFinite(meta.zm)) ? meta.zm : (meta.z||1)));
  bg.style.backgroundSize = `${Math.round(100 * z)}%`;
  const hasPercent = isMobile
    ? ((meta.oxpm != null && isFinite(meta.oxpm)) || (meta.oypm != null && isFinite(meta.oypm)))
    : ((meta.oxp  != null && isFinite(meta.oxp )) || (meta.oyp  != null && isFinite(meta.oyp )));
  if (hasPercent){
    const x = isMobile
      ? `calc(50% + ${(meta.oxpm != null && isFinite(meta.oxpm)) ? meta.oxpm : 0}%)`
      : `calc(50% + ${(meta.oxp  != null && isFinite(meta.oxp )) ? meta.oxp  : 0}%)`;
    const y = isMobile
      ? `calc(50% + ${(meta.oypm != null && isFinite(meta.oypm)) ? meta.oypm : 0}%)`
      : `calc(50% + ${(meta.oyp  != null && isFinite(meta.oyp )) ? meta.oyp  : 0}%)`;
    bg.style.backgroundPosition = `${x} ${y}`;
  } else {
    const ox = isMobile ? (meta.oxm || 0) : (meta.ox || 0);
    const oy = isMobile ? (meta.oym || 0) : (meta.oy || 0);
    bg.style.backgroundPosition = `calc(50% + ${ox}px) calc(50% + ${oy}px)`;
  }
} else {
  bg.style.backgroundSize = 'cover';
  bg.style.backgroundPosition = 'center center';
}
holder.appendChild(bg);
           art.appendChild(holder);
         }
         // Overlay content: only title
         const wrap = document.createElement('div');
         wrap.className = 'slide-content';
         const title = document.createElement('h3'); title.className = 'slide-title';
         const link = document.createElement('a'); link.href = `news.html?id=${encodeURIComponent(n.id)}`; link.textContent = n.title || 'Haber';
         title.appendChild(link); wrap.appendChild(title);
         art.appendChild(wrap);
         slidesRoot.appendChild(art);
       });
      // Notify slider that content changed
      document.dispatchEvent(new CustomEvent('slider:updated'));
    } catch (e) {
      console.error('loadSliderNews error:', e);
    }
  }

  async function loadNewsCards(){
    if (!db()) return;
    const root = document.getElementById('newsCards');
    if (!root) return;
    try {
      const { data, error } = await db()
        .from('news')
        .select('id, title, summary, published_at, unpublish_at, image_url, cover_image_url, status')
        .eq('status', 'published')
        .order('published_at', { ascending: false })
        .limit(9);
      if (error) throw error;

      const now = Date.now();
      const items = (data || []).filter(n => {
        const u = n.unpublish_at ? new Date(n.unpublish_at).getTime() : null;
        return !u || u > now;
      });

      root.innerHTML = '';
      items.forEach(n => {
        const card = document.createElement('article');
        card.className = 'card';
        card.setAttribute('role', 'listitem');
        // Show cover/image + title only
        const imgUrl2 = n.image_url || n.cover_image_url;
        if (imgUrl2) {
          const img = document.createElement('img');
          img.src = imgUrl2; img.alt = n.title || 'Haber görseli'; img.style.borderRadius = '10px';
          card.appendChild(img);
        }
        const h3 = document.createElement('h3'); h3.textContent = n.title || 'Haber';
        const a = document.createElement('a'); a.className = 'card-link'; a.href = `news.html?id=${encodeURIComponent(n.id)}`; a.textContent = 'Haberi Oku →';
        card.appendChild(h3);
        card.appendChild(a);
        root.appendChild(card);
      });
    } catch (e) {
      console.error('loadNewsCards error:', e);
    }
  }

  async function loadAnnouncements(){
    if (!db()) return;
    const ul = document.getElementById('announcementsList');
    if (!ul) return;
    try {
      const { data, error } = await db()
        .from('announcements')
        .select('id, title, body, image_url, published_at, unpublish_at, status')
        .eq('status', 'published')
        .order('published_at', { ascending: false })
        .limit(6);
      if (error) throw error;

      const now = Date.now();
      const items = (data || []).filter(a => {
        const u = a.unpublish_at ? new Date(a.unpublish_at).getTime() : null;
        return !u || u > now;
      });

      ul.innerHTML = '';
      items.forEach(a => {
        const li = document.createElement('li');
        // thumbnail
        if (a.image_url) {
          const img = document.createElement('img');
          img.src = a.image_url;
          img.alt = a.title || 'Duyuru görseli';
          img.style.width = '48px';
          img.style.height = '48px';
          img.style.objectFit = 'cover';
          img.style.borderRadius = '8px';
          img.style.marginRight = '8px';
          img.style.verticalAlign = 'middle';
          li.appendChild(img);
        }
        const strong = document.createElement('strong'); strong.textContent = (a.title || 'Duyuru') + ': ';
        const span = document.createElement('span'); span.textContent = a.body || '';
        li.appendChild(strong);
        li.appendChild(span);
        ul.appendChild(li);
      });
    } catch (e) {
      console.error('loadAnnouncements error:', e);
    }
  }

  async function loadFooterSettings(){
    if (!db()) return;
    try {
      const { data, error } = await db()
        .from('settings')
        .select('key, value')
        .in('key', ['footer_logo_url','header_logo_url','footer_address','footer_email','footer_phone','social_twitter','social_facebook','social_instagram','social_youtube']);
      if (error) throw error;
      const map = new Map((data||[]).map(r => [r.key, r.value]));
      const logoUrl = map.get('footer_logo_url') || 'ilke-logo.png';
      const headerLogoUrl = map.get('header_logo_url') || null;
      const addr = map.get('footer_address') || null;
      const email = map.get('footer_email') || null;
      const phone = map.get('footer_phone') || null;
      const sTwitter = map.get('social_twitter') || '';
      const sFacebook = map.get('social_facebook') || '';
      const sInstagram = map.get('social_instagram') || '';
      const sYouTube = map.get('social_youtube') || '';

      // Cache-bust images to avoid stale assets
      const buster = (url) => {
        try { const u = new URL(url, window.location.origin); u.searchParams.set('v', Date.now().toString()); return u.toString(); }
        catch { return url; }
      };
       // Use only CSP-allowed origins for images (self or *.supabase.co)
       const isAllowedImageUrl = (u) => {
        try {
          if (!u) return false;
          // Allow relative/local paths
          if (!/^https?:/i.test(u)) return true;
          const host = new URL(u, window.location.origin).host.toLowerCase();
          if (host === window.location.host.toLowerCase()) return true;
          if (/\\.supabase\\.co$/i.test(host)) return true;
          if (host === 'www.google.com' || host === 'www.gstatic.com') return true;
          return false;
        } catch { return false; }
      };
      const effectiveLogoUrl = isAllowedImageUrl(logoUrl) ? logoUrl : 'ilke-logo.png';
      const effectiveHeaderLogoUrl = isAllowedImageUrl(headerLogoUrl) ? headerLogoUrl : effectiveLogoUrl;
      const footerLogo = document.getElementById('footerLogo'); if (footerLogo && logoUrl) footerLogo.src = buster(effectiveLogoUrl);
      const headerLogo = document.getElementById('headerLogo'); if (headerLogo && (headerLogoUrl || logoUrl)) headerLogo.src = buster(effectiveHeaderLogoUrl);
      const addrEl = document.getElementById('footerAddress');
      if (addrEl){
        const full = String(addr || addrEl.textContent || '').trim();
        let short = full;
        try{
          // Tercih: "İlçe/İl" ya da "İlçe, İl"
          const slashIdx = full.lastIndexOf('/');
          if (slashIdx > -1){
            const left = full.slice(0, slashIdx).trim();
            const right = full.slice(slashIdx+1).trim();
            const leftTokens = left.split(/\s+/).filter(Boolean);
            const district = leftTokens.length ? leftTokens[leftTokens.length-1].replace(/[.,;:]$/,'') : left;
            short = `${district}, ${right}`;
          } else if (full.includes(',')){
            const parts = full.split(',').map(s=>s.trim()).filter(Boolean);
            if (parts.length >= 2){
              const left = parts[parts.length-2];
              const right = parts[parts.length-1];
              const lt = left.split(/\s+/).filter(Boolean);
              const district = lt.length ? lt[lt.length-1].replace(/[.,;:]$/,'') : left;
              short = `${district}, ${right}`;
            } else {
              short = parts.join(', ');
            }
          } else {
            // Yedek: sondan iki kelime
            const toks = full.split(/\s+/).filter(Boolean);
            short = toks.length >= 2 ? `${toks[toks.length-2]} ${toks[toks.length-1]}` : full;
          }
        }catch{}
        addrEl.textContent = short;
      }
      const emailEl = document.getElementById('footerEmail'); if (emailEl && email) emailEl.textContent = email;
      const phoneEl = document.getElementById('footerPhone'); if (phoneEl && phone) phoneEl.textContent = phone;

      function setLink(id, url){
        const a = document.getElementById(id);
        if (!a) return;
        if (url && /^https?:\/\//i.test(url)) { a.href = url; a.target = '_blank'; a.rel = 'noopener noreferrer'; a.style.display=''; }
        else { a.style.display='none'; }
      }
      setLink('socialTwitter', sTwitter);
      setLink('socialFacebook', sFacebook);
      setLink('socialInstagram', sInstagram);
      setLink('socialYouTube', sYouTube);
    } catch(e){
      console.warn('loadFooterSettings error:', e);
    }
    
  }

  async function loadContactSettings(){
    if (!db()) return;
    try{
      const { data, error } = await db()
        .from('settings')
        .select('key, value')
        .in('key', ['footer_address','footer_email','footer_phone','recaptcha_site_key']);
      if (error) throw error;
      const map = new Map((data||[]).map(r => [r.key, r.value]));
      const setText = (id, val) => { const el = document.getElementById(id); if (el && val) el.textContent = val; };
      setText('contactAddress', map.get('footer_address') || '');
      setText('contactEmail', map.get('footer_email') || '');
      setText('contactPhone', map.get('footer_phone') || '');
      const siteKey = map.get('recaptcha_site_key');
      if (siteKey){
        window.IlkeSendikaRecaptcha = window.IlkeSendikaRecaptcha || {};
        window.IlkeSendikaRecaptcha.siteKey = siteKey;
      }
    }catch(e){
      console.warn('loadContactSettings error:', e);
    }
  }

  // Home "Önemli Günler" carousel (3-up, auto-rotate, keyboard + buttons)
  async function loadHomePosters(){
    if (!db()) return;
    const wrap = document.getElementById('homePosters'); if (!wrap) return;
    const track = wrap.querySelector('[data-hp-track]'); if (!track) return;
    try{
      // Inject styles once
      if (!document.getElementById('homePostersCss')){
        const st=document.createElement('style'); st.id='homePostersCss'; st.textContent=`
          .home-posters{ position:relative; overflow:hidden; padding:0 36px; min-height:160px; }
          .home-posters .hp-track{ display:flex; gap:12px; will-change:transform; transition:transform .4s ease; }
          .home-posters .hp-item{ flex:0 0 calc(33.333% - 8px); box-sizing:border-box; }
          .home-posters .hp-card{ background:#fff; border-radius:10px; overflow:hidden; box-shadow:0 2px 10px rgba(0,0,0,.06); }
          .home-posters .hp-card img{ width:100%; height:140px; object-fit:contain; display:block; background:#fff; }
          .home-posters .hp-nav{ position:absolute; top:50%; transform:translateY(-50%); border:none; background:rgba(0,0,0,.35); color:#fff; width:28px; height:36px; border-radius:10px; cursor:pointer; }
          .home-posters .hp-prev{ left:4px; }
          .home-posters .hp-next{ right:4px; }
          @media (max-width: 900px){ .home-posters .hp-item{ flex-basis: calc(50% - 8px); } }
          @media (max-width: 560px){ .home-posters .hp-item{ flex-basis: 100%; } }
        `; document.head.appendChild(st);
      }

      const { data, error } = await db()
        .from('posters')
        .select('id, title, image_url, status, published_at')
        .eq('status','published')
        .order('published_at', { ascending:false, nullsFirst:true })
        .limit(12);
      if (error) throw error;

      const items = (data||[]).filter(p=>p.image_url);
      track.innerHTML='';
      items.forEach(p=>{
        const li=document.createElement('div'); li.className='hp-item';
        const card=document.createElement('article'); card.className='hp-card';
        const a=document.createElement('a'); a.href='page.html?slug=afis'; a.setAttribute('aria-label', p.title||'Afiş');
        const img=document.createElement('img'); img.src=p.image_url; img.alt=p.title||'Afiş';
        a.appendChild(img); card.appendChild(a); li.appendChild(card); track.appendChild(li);
      });

      // Carousel state
      let idx = 0; let timer = null;
      const prevBtn = wrap.querySelector('[data-hp-prev]');
      const nextBtn = wrap.querySelector('[data-hp-next]');

      function visible(){ const w = wrap.clientWidth; if (w <= 560) return 1; if (w <= 900) return 2; return 3; }
      function clampIndex(){ const vis = visible(); const max = Math.max(0, items.length - vis); if (idx < 0) idx = 0; if (idx > max) idx = max; return { vis, max }; }
      function update(){ const { vis } = clampIndex(); const item = track.querySelector('.hp-item'); if (!item) return; const itemWidth = item.getBoundingClientRect().width + 12; const off = -(idx * itemWidth); track.style.transform = `translate3d(${off}px,0,0)`; }
      function go(d){ idx += d; update(); }
      function start(){ stop(); timer = setInterval(()=>{ const { max } = clampIndex(); if (idx >= max) idx = 0; else idx++; update(); }, 5000); }
      function stop(){ if (timer){ clearInterval(timer); timer = null; } }

      prevBtn && prevBtn.addEventListener('click', ()=> go(-1));
      nextBtn && nextBtn.addEventListener('click', ()=> go(1));
      wrap.addEventListener('mouseenter', stop); wrap.addEventListener('mouseleave', start);
      // Keyboard arrows
      const onKey = (e)=>{ if (e.key==='ArrowLeft') go(-1); else if (e.key==='ArrowRight') go(1); };
      document.addEventListener('keydown', onKey);
      // Keep layout responsive
      window.addEventListener('resize', update);
      // Initial
      update(); start();
    }catch(e){ console.warn('loadHomePosters error:', e); }
  }

  // Expose for manual refresh from console if needed
  // Provide policy reader for modal consumers
  async function getPolicy(key){
    if (!db()) return '';
    try {
      const { data, error } = await db()
        .from('settings')
        .select('value')
        .eq('key', key)
        .maybeSingle();
      if (error) throw error;
      const val = data?.value || '';
      if (val && String(val).trim()) return val;
      // Fallback to pages table for known policy slugs
      const slug = key === 'policy_privacy' ? 'gizlilik' : (key === 'policy_disclosure' ? 'aydinlatma' : '');
      if (!slug) return '';
      try{
        const { data: pageRow } = await db()
          .from('pages')
          .select('body, status')
          .eq('slug', slug)
          .eq('status', 'published')
          .maybeSingle();
        return String(pageRow?.body || '');
      }catch{ return ''; }
    } catch(e){
      console.warn('getPolicy error:', e);
      return '';
    }
  }

  window.IlkeSendikaReload = {
    loadSliderNews,
    loadNewsCards,
    loadAnnouncements,
    loadFooterSettings,
    getPolicy,
    loadContactSettings,
    loadHomePosters
  };
})();

// ============ Motion & Interactions (Reveal + Ticker) ============
(function(){
  function qsa(sel, root=document){ return Array.from(root.querySelectorAll(sel)); }

  // Scroll-Reveal: applies .revealed when in viewport
  window.initReveal = function(){
    const candidates = qsa('.section, .card, .slide');
    if (!('IntersectionObserver' in window) || candidates.length === 0) return;
    candidates.forEach(el => el.classList.add('reveal'));
    const obs = new IntersectionObserver((entries)=>{
      entries.forEach(entry => {
        if (entry.isIntersecting){
          entry.target.classList.add('revealed');
          obs.unobserve(entry.target);
        }
      });
    }, { root: null, rootMargin: '0px 0px -5% 0px', threshold: 0.12 });
    candidates.forEach(el => obs.observe(el));
  };

  // Announcements auto-scroll (vertical ticker)
  window.initAnnouncementsTicker = function(){
    const list = document.getElementById('announcementsList');
    if (!list) return;
    const items = qsa('li', list);
    if (items.length <= 1) return;

    // Ensure a fixed-height viewport to clip the scrolling list within the card
    function ensureViewport(){
      try{
        const parent = list.parentElement;
        if (!parent) return;
        const lis = Array.from(list.querySelectorAll('li'));
        if (!lis.length) return;
        let maxH = 0;
        lis.forEach(li => {
          const rect = li.getBoundingClientRect();
          const cs = window.getComputedStyle(li);
          const mt = parseFloat(cs.marginTop || '0') || 0;
          const mb = parseFloat(cs.marginBottom || '0') || 0;
          const h = rect.height + mt + mb;
          if (h > maxH) maxH = h;
        });
        const viewH = Math.ceil(maxH) + 2; // small buffer
        const loading = anyImgLoading();
        if (parent.classList && parent.classList.contains('ann-viewport')){
          parent.style.overflow = 'hidden';
          parent.style.height = loading ? '' : (viewH + 'px');
          parent.style.width = '100%';
        }else{
          const vp = document.createElement('div');
          vp.className = 'ann-viewport';
          vp.style.overflow = 'hidden';
          vp.style.height = loading ? '' : (viewH + 'px');
          vp.style.width = '100%';
          parent.replaceChild(vp, list);
          vp.appendChild(list);
        }
      }catch(e){}
    }
    function anyImgLoading(){
      try { return Array.from(list.querySelectorAll('img')).some(img => !img.complete); }
      catch { return false; }
    }
    ensureViewport();
    // Also adjust on image loads and list DOM mutations (data loads, etc.)
    function attachImageLoadHandlers(){
      try{
        const imgs = list.querySelectorAll('img');
        imgs.forEach(img => { if (!img.complete) img.addEventListener('load', ensureViewport, { once: true }); });
      }catch{}
    }
    attachImageLoadHandlers();
    const mo = new MutationObserver(() => { ensureViewport(); attachImageLoadHandlers(); maybeStart(); });
    try{ mo.observe(list, { childList: true, subtree: true }); }catch{}
    window.addEventListener('resize', () => { ensureViewport(); maybeStart(); });
    window.addEventListener('load', () => { ensureViewport(); maybeStart(); });
    [0,50,150,400,800].forEach(t => setTimeout(() => { ensureViewport(); maybeStart(); }, t));

    let timer = null; let started = false;
    function step(){
      const first = list.querySelector('li');
      if (!first) return;
      const h = first.getBoundingClientRect().height;
      // Keep viewport height in sync with current item height (handles dynamic content/load)
      ensureViewport();
      list.style.transition = 'transform 500ms ease';
      list.style.transform = `translateY(-${h}px)`;
      const onEnd = ()=>{
        list.style.transition = 'none';
        list.style.transform = 'translateY(0)';
        list.appendChild(first);
        list.removeEventListener('transitionend', onEnd);
        // allow next tick to re-enable transition
        setTimeout(()=>{ list.style.transition = ''; }, 20);
      };
      list.addEventListener('transitionend', onEnd);
    }
    function start(){ if (!timer) timer = setInterval(step, 3500); }
    function stop(){ if (timer){ clearInterval(timer); timer = null; } }
    function maybeStart(){ if (started) return; if (anyImgLoading()) return; if (list.querySelectorAll('li').length <= 1) return; start(); started = true; }
    list.addEventListener('mouseenter', stop); list.addEventListener('mouseleave', start);
    maybeStart();
  };
})();

