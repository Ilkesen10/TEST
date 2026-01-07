(function(){
  'use strict';
  const $ = (s, r=document)=> r.querySelector(s);

  function getQuery(key){
    const u = new URL(location.href);
    return u.searchParams.get(key);
  }

  function getSupabase(){
    if (window.ilkeSupabase) return window.ilkeSupabase;
    try{
      const url = document.querySelector('meta[name="supabase-url"]')?.content || '';
      const anon = document.querySelector('meta[name="supabase-anon"]')?.content || '';
      if (url && anon){
        const { createClient } = window.supabase || supabase;
        return createClient(url, anon);
      }
    }catch{}
    return null;
  }

  function qs(sel, root=document){ return root.querySelector(sel); }
  const statusEl = qs('#status') || qs('#verifyStatus');
  const box = qs('#result') || qs('#verifyResult');

  async function callVerify(token){
    try{
      const sb = window.ilkeSupabase;
      if (!sb) throw new Error('Supabase istemcisi yok');
      const { data, error } = await sb.functions.invoke('verify_membership', { body: { t: String(token||'').trim() } });
      if (error) throw error;
      return data || {};
    }catch(e){
      return { ok:false, error: e?.message || String(e) };
    }
  }

  function renderVerified(member){
    if (statusEl) statusEl.textContent = 'Üyelik doğrulandı';
    if (!box) return;
    box.style.display = '';
    const full = member?.full_name || '';
    const no = member?.membership_no || '';
    const st = member?.status || '';
    const meta = [no?`Üye No: ${no}`:'', st?`Durum: ${st}`:''].filter(Boolean).join(' • ');
    // Build DOM safely without innerHTML
    box.innerHTML = '';
    const row = document.createElement('div'); row.className = 'row';
    const wrap = document.createElement('div');
    const nameEl = document.createElement('div'); nameEl.id = 'fullName'; nameEl.className = 'name'; nameEl.textContent = full;
    const metaEl = document.createElement('div'); metaEl.id = 'meta'; metaEl.className = 'muted'; metaEl.textContent = meta;
    const badgeWrap = document.createElement('div'); badgeWrap.id = 'badge'; badgeWrap.style.marginTop = '8px';
    const badgeSpan = document.createElement('span'); badgeSpan.className = 'badge ok'; badgeSpan.textContent = '✔ Geçerli Üyelik';
    badgeWrap.appendChild(badgeSpan);
    wrap.appendChild(nameEl); wrap.appendChild(metaEl); wrap.appendChild(badgeWrap);
    row.appendChild(wrap);
    box.appendChild(row);
  }

  function renderInvalid(reason){
    if (statusEl) statusEl.textContent = 'Doğrulama başarısız';
    if (!box) return;
    const why = reason==='inactive' ? 'Üyelik pasif' : 'Geçersiz ya da eski QR';
    box.style.display='';
    box.innerHTML = '';
    const el = document.createElement('div'); el.className = 'badge bad'; el.textContent = `✖ ${why}`;
    box.appendChild(el);
  }

  async function run(token){
    if (!token){ if (statusEl) statusEl.textContent = 'Doğrulama kodu bulunamadı'; return; }
    if (statusEl) statusEl.textContent = 'Sorgulanıyor…';
    const res = await callVerify(token);
    if (res && res.ok && res.verified){ return renderVerified(res.member||{}); }
    if (res && res.ok && res.verified===false){ return renderInvalid(res.reason); }
    renderInvalid('error');
  }

  async function main(){
    const id = getQuery('id');
    if (!id){ statusEl.textContent = 'Geçersiz istek: id parametresi eksik.'; return; }

    const sb = getSupabase();
    if (!sb){ statusEl.textContent = 'Sunucu yapılandırması eksik (Supabase).'; return; }

    try{
      statusEl.textContent = 'Sorgulanıyor…';
      const { data, error } = await sb.from('members').select('*').eq('id', id).maybeSingle();
      if (error) throw error;
      if (!data){ statusEl.textContent = 'Kayıt bulunamadı.'; return; }
      const m = data;
      $('#fullName').textContent = `${m.first_name||''} ${m.last_name||''}`.trim();
      $('#meta').textContent = `Üye No: ${m.member_no||'-'} • TC: ${m.national_id||'-'}`;
      $('#v_member_no').textContent = m.member_no || '-';
      $('#v_tc').textContent = m.national_id || '-';
      $('#v_status').textContent = (m.status === 'active') ? 'Aktif' : (m.status || '-');
      $('#v_join').textContent = m.join_date ? new Date(m.join_date).toLocaleDateString('tr-TR') : '-';
      const ph = $('#photo');
      if (m.photo_url){ ph.src = m.photo_url; ph.style.display='block'; } else { ph.style.display='none'; }
      const badge = document.createElement('span');
      badge.className = 'badge ' + (m.status === 'active' ? 'ok' : 'bad');
      badge.textContent = m.status === 'active' ? 'Doğrulandı' : 'Pasif';
      $('#badge').innerHTML = ''; $('#badge').appendChild(badge);
      statusEl.style.display = 'none';
      box.style.display = 'block';
    }catch(e){
      statusEl.textContent = 'Doğrulama sırasında hata: ' + (e?.message||String(e));
    }
  }

  document.addEventListener('DOMContentLoaded', ()=>{
    try{
      const p = new URLSearchParams(location.search);
      const t = p.get('t') || p.get('token') || '';
      if (t) return run(t);
      main();
    }catch{}
  });
})();
