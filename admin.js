'use strict';

(function(){
  const sb = () => window.ilkeSupabase;
  const qs = (s, r=document) => r.querySelector(s);
  const qsa = (s, r=document) => Array.from(r.querySelectorAll(s));

  const $loginSection = () => qs('#loginSection');
  const $adminSection = () => qs('#adminSection');
  const $authStatus = () => qs('#authStatus');
  const $newsTBody = () => qs('#newsTableBody');
  const $annTBody = () => qs('#annTableBody');
  const $msgsTBody = () => qs('#msgsTableBody');
  const $membersTBody = () => qs('#membersTableBody');
  const $usersTBody = () => qs('#usersTableBody');
  const $foundersTBody = () => qs('#foundersTableBody');
  const $chairmanTBody = () => qs('#chairmanTableBody');
  const $modal = () => qs('#modal');
  const $modalTitle = () => qs('#modalTitle');
  const $modalForm = () => qs('#modalForm');

  // Modal helpers
  function openModal(){
    const m = $modal();
    if (!m) return;
    m.style.display = 'block';
    m.style.maxHeight = '90vh';
    m.style.overflow = 'auto';
  }

  async function generateNextOutgoingRecordNo(){
    try{
      const year = new Date().getFullYear();
      const prefix = `${year}/`;
      const { data } = await sb().from('outgoing_docs')
        .select('record_no')
        .ilike('record_no', `${prefix}%`)
        .order('record_no', { ascending:false })
        .limit(1000);
      let maxSeq = 0;
      (data||[]).forEach(r => {
        const s = String(r.record_no||'');
        if (s.startsWith(prefix)){
          const tail = s.slice(prefix.length).replace(/[^0-9]/g,'');
          const n = parseInt(tail||'0',10);
          if (!isNaN(n) && n > maxSeq) maxSeq = n;
        }
      });
      const next = maxSeq + 1;
      const pad = String(next).padStart(3,'0');
      return `${prefix}${pad}`;
    }catch{ return `${new Date().getFullYear()}/001`; }
  }

  async function openOutgoingDocModal(row){
    try{
      row = row || null;
      const isEdit = !!(row && row.id);
      $modalTitle().textContent = isEdit ? 'Giden Evrak Düzenle' : 'Yeni Giden Evrak';
      const form = $modalForm(); form.innerHTML='';
      let recordNo = isEdit && row.record_no ? String(row.record_no) : await generateNextOutgoingRecordNo();
      const today = new Date();
      const dateIsoDefault = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate())).toISOString().slice(0,10);

      const recLbl = document.createElement('label'); recLbl.style.display='grid'; recLbl.style.gap='6px'; recLbl.innerHTML = '<span>Kayıt Sıra No</span>';
      const recIn = document.createElement('input'); recIn.readOnly = true; recIn.disabled = true; recIn.value = recordNo; recLbl.appendChild(recIn); form.appendChild(recLbl);

      const dateLbl = document.createElement('label'); dateLbl.style.display='grid'; dateLbl.style.gap='6px'; dateLbl.innerHTML = '<span>Tarihi</span>';
      const dateIn = document.createElement('input'); dateIn.type='date'; dateIn.value = row?.date ? String(row.date).slice(0,10) : dateIsoDefault; dateLbl.appendChild(dateIn); form.appendChild(dateLbl);

      const toLbl = document.createElement('label'); toLbl.style.display='grid'; toLbl.style.gap='6px'; toLbl.innerHTML = '<span>Gönderildiği Kişi/Kurum</span>';
      const toIn = document.createElement('textarea'); toIn.rows = 2; toIn.value = row?.to_org || ''; toLbl.appendChild(toIn); form.appendChild(toLbl);

      const attLbl = document.createElement('label'); attLbl.style.display='grid'; attLbl.style.gap='6px'; attLbl.innerHTML = '<span>Eki</span>';
      const attIn = document.createElement('input'); attIn.value = row?.attachment || ''; attLbl.appendChild(attIn); form.appendChild(attLbl);

      const subLbl = document.createElement('label'); subLbl.style.display='grid'; subLbl.style.gap='6px'; subLbl.innerHTML = '<span>Konusu</span>';
      const subIn = document.createElement('input'); subIn.value = row?.subject || ''; subLbl.appendChild(subIn); form.appendChild(subLbl);

      const fileNoLbl = document.createElement('label'); fileNoLbl.style.display='grid'; fileNoLbl.style.gap='6px'; fileNoLbl.innerHTML = '<span>Muhafaza Dosya No</span>';
      const fileNoIn = document.createElement('input'); fileNoIn.value = row?.file_no || ''; fileNoLbl.appendChild(fileNoIn); form.appendChild(fileNoLbl);

      const actions = document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
      const genBtn = document.createElement('button'); genBtn.className='btn btn-primary'; genBtn.textContent='Belge Oluştur';
      const cancelBtn = document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
      actions.appendChild(genBtn); actions.appendChild(cancelBtn); form.appendChild(actions);
      cancelBtn.addEventListener('click', (e)=>{ e.preventDefault(); closeModal(); });
      genBtn.addEventListener('click', async (e)=>{
        e.preventDefault();
        try{
          genBtn.disabled = true; genBtn.textContent = 'Yükleniyor...';
          recordNo = String(recIn.value||recordNo);
          const dSel = dateIn.value ? new Date(dateIn.value+'T12:00:00Z') : today;
          const dateIso = dateIn.value ? String(dateIn.value) : dateIsoDefault;
          const toOrg = String(toIn.value||'');
          const att = String(attIn.value||'');
          const subject = String(subIn.value||'');
          const fileNo = String(fileNoIn.value||'');
          const resp = await fetch('assets/sablon.docx'); if (!resp.ok) throw new Error('Şablon alınamadı');
          const tmplAb = await resp.arrayBuffer();
          const zip = new (window.PizZip||window.pizzip||window.PIZZip||PizZip)(tmplAb);
          const DocxClass = window.docxtemplater || window.Docxtemplater || window.DOCXTEMPLATER; if (!DocxClass) throw new Error('Docxtemplater yüklenemedi');
          const doc = new DocxClass(zip, { paragraphLoop:true, linebreaks:true });
          const trDate = dSel.toLocaleDateString('tr-TR');
          doc.setData({ KAYITSIRANO: recordNo, KAYIT_SIRA_NO: recordNo, TARIH: trDate, KISI_KURUM: toOrg, EK: att, EKI: att, KONUSU: subject, DOSYA_NO: fileNo });
          try{ doc.render(); }catch{ throw new Error('Şablon doldurulamadı'); }
          const filledAb = doc.getZip().generate({ type:'arraybuffer' });
          const res = await window.mammoth.convertToHtml({ arrayBuffer: filledAb });
          const html = res && res.value ? res.value : '<div></div>';
          form.innerHTML = '';
          const ed = document.createElement('div'); ed.contentEditable='true'; ed.style.minHeight='900px'; ed.style.padding='24px'; ed.style.background='#fff'; ed.style.border='1px solid #e5e7eb'; ed.style.borderRadius='10px'; ed.style.overflow='auto'; ed.innerHTML = html; form.appendChild(ed);
          const act = document.createElement('div'); act.style.display='flex'; act.style.gap='8px'; act.style.marginTop='8px';
          const saveBtn = document.createElement('button'); saveBtn.className='btn btn-success'; saveBtn.textContent='Kaydet';
          const cancelBtn2 = document.createElement('button'); cancelBtn2.className='btn btn-danger'; cancelBtn2.textContent='İptal';
          act.appendChild(saveBtn); act.appendChild(cancelBtn2); form.appendChild(act);
          cancelBtn2.addEventListener('click', (e2)=>{ e2.preventDefault(); closeModal(); });
          saveBtn.addEventListener('click', async (e2)=>{
            e2.preventDefault();
            try{
              saveBtn.disabled = true; saveBtn.textContent='Kaydediliyor...';
              const canvas = await html2canvas(ed, { scale: 2, backgroundColor:'#ffffff' });
              const img = canvas.toDataURL('image/png');
              const jsPDF = window.jspdf && window.jspdf.jsPDF ? window.jspdf.jsPDF : null; if (!jsPDF) throw new Error('PDF oluşturma hazır değil');
              const pdf = new jsPDF({ unit:'pt', format:'a4', orientation:'portrait' });
              const pw = pdf.internal.pageSize.getWidth(); const ph = pdf.internal.pageSize.getHeight();
              const iw = canvas.width; const ih = canvas.height; const r = Math.min(pw/iw, ph/ih);
              const w = iw*r; const h = ih*r; const x = (pw - w)/2; const y = (ph - h)/2;
              pdf.addImage(img, 'PNG', x, y, w, h, undefined, 'FAST');
              const blob = pdf.output('blob');
              const safeKey = recordNo.replace(/\//g,'_');
              const key = `outgoing/${new Date().getFullYear()}/${safeKey}_${Date.now()}.pdf`;
              const fileUrl = await uploadToBucketGeneric(blob, 'docs', key);
              const payload = { record_no: recordNo, date: dateIso, to_org: toOrg, attachment: att, subject: subject, file_no: fileNo, file_url: fileUrl };
              let q; if (isEdit && row.id){ q = sb().from('outgoing_docs').update(payload).eq('id', row.id); } else { q = sb().from('outgoing_docs').insert(payload).select('id').single(); }
              const { error } = await q; if (error) throw error;
              closeModal(); await loadAdminOutgoingDocs();
            }catch(err){ alert('Kaydedilemedi: ' + (err?.message||String(err))); saveBtn.textContent='Kaydet'; saveBtn.disabled=false; }
          });
        }catch(e){ alert('Belge açılmadı: ' + (e?.message||String(e))); genBtn.textContent='Belge Oluştur'; genBtn.disabled=false; }
      });
      (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
    }catch(e){ alert('Evrak formu açılamadı: ' + (e?.message||String(e))); }
  }

  // ========== BENEFITS (Üyelere Özel) ==========
  async function loadAdminBenefits(){
    const tbody = qs('#benefitsTableBody'); if (!tbody) return; tbody.innerHTML = '';
    try{
      const table = tbody.closest('table');
      const headRow = table && table.tHead && table.tHead.rows && table.tHead.rows[0];
      if (headRow){
        headRow.innerHTML = `
          <th>Logo</th>
          <th>Firma</th>
          <th>İndirim</th>
          <th>İletişim</th>
          <th>Durum</th>
          <th>Sıra</th>
          <th class="col-actions">İşlem</th>
        `;
      }
    }catch{}
    try{
      const { data, error } = await sb()
        .from('benefits')
        .select('id, name, discount_text, contact, website, phone, address, logo_url, status, sort, published_at')
        .order('sort', { ascending:true, nullsFirst:true })
        .order('name', { ascending:true });
      if (error) throw error;
      (data||[]).forEach(row=>{
        const tr = document.createElement('tr');
        const statusMap = { published:'Yayımlandı', draft:'Taslak', unpublished:'Yayından Kaldırıldı' };
        const statusTr = statusMap[String(row.status||'').toLowerCase()] || (row.status||'');
        const contact = row.phone ? ('Tel: ' + row.phone) : (row.website || '');
        tr.innerHTML = `
          <td>${row.logo_url ? `<img src="${escapeHtml(row.logo_url)}" alt="logo" style="max-height:34px">` : '-'}</td>
          <td>${escapeHtml(row.name||'')}</td>
          <td>${escapeHtml(row.discount_text||'')}</td>
          <td>${escapeHtml(contact||'')}</td>
          <td>${escapeHtml(statusTr)}</td>
          <td>${row.sort != null ? String(row.sort) : '-'}</td>
          <td class="actions">
            <button class="btn btn-warning" data-edit-benefit="${row.id}">Düzenle</button>
            ${row.status==='published' ? `<button class="btn btn-danger" data-unpub-benefit="${row.id}">Yayından Kaldır</button>` : `<button class="btn btn-success" data-pub-benefit="${row.id}">Yayınla</button>`}
          </td>`;
        tbody.appendChild(tr);
      });
      const addBtn = qs('#newBenefitBtn'); if (addBtn && !addBtn.dataset.wired){ addBtn.dataset.wired='1'; addBtn.onclick = ()=> openBenefitModal({ id:null, name:'', discount_text:'', website:'', phone:'', address:'', logo_url:'', status:'draft', sort:null, published_at:null }); }
      wireBenefitsRowActions();
    }catch(e){ alert('Üyelere Özel listesi yüklenemedi: ' + (e?.message||String(e))); }
  }

  function wireBenefitsRowActions(){
    try{
      const tbody = qs('#benefitsTableBody'); if (!tbody) return;
      tbody.querySelectorAll('button[data-edit-benefit]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-edit-benefit');
          try{ const { data } = await sb().from('benefits').select('*').eq('id', id).maybeSingle(); openBenefitModal(data||{ id, name:'', discount_text:'', website:'', phone:'', address:'', logo_url:'', status:'draft', sort:null, published_at:null }); }
          catch(err){ alert('Kayıt yüklenemedi: ' + (err?.message||String(err))); }
        });
      });
      tbody.querySelectorAll('button[data-pub-benefit]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-pub-benefit');
          try{ const { error } = await sb().from('benefits').update({ status:'published', published_at: new Date().toISOString() }).eq('id', id); if (error) throw error; await loadAdminBenefits(); }
          catch(err){ alert('Yayınlanamadı: ' + (err?.message||String(err))); }
        });
      });
      tbody.querySelectorAll('button[data-unpub-benefit]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-unpub-benefit');
          try{ const { error } = await sb().from('benefits').update({ status:'unpublished' }).eq('id', id); if (error) throw error; await loadAdminBenefits(); }
          catch(err){ alert('Yayından kaldırılamadı: ' + (err?.message||String(err))); }
        });
      });
    }catch{}
  }

  function openBenefitModal(row){
    try{
      row = row || { id:null, name:'', discount_text:'', website:'', phone:'', address:'', logo_url:'', status:'draft', sort:null, published_at:null };
      $modalTitle().textContent = row.id ? 'İş Yerini Düzenle' : 'Yeni İş Yeri';
      const form = $modalForm(); form.innerHTML='';
      const g = document.createElement('div'); g.style.display='grid'; g.style.gap='8px'; form.appendChild(g);
      function field(label, input){ const L=document.createElement('label'); L.style.display='grid'; L.style.gap='6px'; L.innerHTML = `<span>${label}</span>`; L.appendChild(input); g.appendChild(L); return input; }
      const iName=field('Firma Adı', Object.assign(document.createElement('input'), { value: row.name||'' }));
      const iDiscount=field('İndirim / Kampanya', Object.assign(document.createElement('input'), { value: row.discount_text||'' }));
      const iWebsite=field('Web Sitesi (https://)', Object.assign(document.createElement('input'), { value: row.website||'' }));
      const iPhone=field('Telefon', Object.assign(document.createElement('input'), { value: row.phone||'' }));
      const iAddress=field('Adres', Object.assign(document.createElement('textarea'), { value: row.address||'' }));
      const iLogo=field('Logo URL', Object.assign(document.createElement('input'), { value: row.logo_url||'' }));
      const iFile=document.createElement('input'); iFile.type='file'; iFile.accept='image/*'; field('Logo Yükle', iFile);
      const stSel=document.createElement('select');
      [
        { v:'draft', t:'Taslak' },
        { v:'published', t:'Yayımlandı' },
        { v:'unpublished', t:'Yayından Kaldırıldı' },
      ].forEach(({v,t})=>{ const o=document.createElement('option'); o.value=v; o.textContent=t; if ((row.status||'draft')===v) o.selected=true; stSel.appendChild(o); });
      field('Durum', stSel);
      const iSort=field('Sıra', Object.assign(document.createElement('input'), { type:'number', value: row.sort==null?'':String(row.sort) }));
      if (!row.id && (row.sort == null || row.sort === undefined)) {
        (async ()=>{
          try{
            const { data: maxRow } = await sb().from('benefits').select('sort').order('sort', { ascending:false, nullsLast:true }).limit(1).maybeSingle();
            const next = (Number(maxRow?.sort || 0) || 0) + 1;
            if (!iSort.value) iSort.value = String(next);
          }catch{}
        })();
      }
      const actions=document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
      const saveBtn=document.createElement('button'); saveBtn.type='button'; saveBtn.className='btn btn-success'; saveBtn.textContent='Kaydet'; actions.appendChild(saveBtn);
      const cancelBtn=document.createElement('button'); cancelBtn.type='button'; cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal'; actions.appendChild(cancelBtn);
      if (row.id){
        const delBtn=document.createElement('button'); delBtn.type='button'; delBtn.className='btn btn-danger'; delBtn.textContent='Sil'; actions.appendChild(delBtn);
        delBtn.addEventListener('click', async (e)=>{
          e.preventDefault();
          try{
            if (!confirm('Bu kaydı silmek istediğinize emin misiniz?')) return;
            const { error } = await sb().from('benefits').delete().eq('id', row.id);
            if (error) throw error;
            closeModal(); await loadAdminBenefits();
          }catch(err){ alert('Silinemedi: ' + (err?.message||String(err))); }
        });
      }
      g.appendChild(actions);
      const origSaveText = saveBtn.textContent;
      let saving = false;
      cancelBtn.addEventListener('click', (e)=>{ e.preventDefault(); closeModal(); });
      saveBtn.addEventListener('click', async (e)=>{
        e.preventDefault();
        if (saving) return;
        saving = true;
        saveBtn.disabled = true;
        saveBtn.textContent = 'Kaydediliyor...';
        try{
          const payload = {
            name: String(iName.value||'').trim(),
            discount_text: String(iDiscount.value||'').trim(),
            website: String(iWebsite.value||'').trim(),
            phone: String(iPhone.value||'').trim(),
            address: String(iAddress.value||'').trim(),
            logo_url: String(iLogo.value||'').trim(),
            status: String(stSel.value||'draft'),
            sort: iSort.value ? Number(iSort.value) : null,
          };
          if (!payload.name){ alert('Firma adı gerekli'); return; }
          try{
            if (iFile.files && iFile.files[0]){
              const f=iFile.files[0]; const ext=(f.name.split('.').pop()||'jpg').toLowerCase();
              const key=`benefits/${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
              payload.logo_url = await uploadToBucketGeneric(f, 'images', key);
            }
            if (payload.status==='published' && !row.published_at){ payload.published_at = new Date().toISOString(); }
            if (!row.id && (payload.sort == null)) {
              try{
                const { data: maxRow } = await sb().from('benefits').select('sort').order('sort', { ascending:false, nullsLast:true }).limit(1).maybeSingle();
                const next = (Number(maxRow?.sort || 0) || 0) + 1;
                payload.sort = next;
              }catch{}
            }
            const q = row.id
              ? sb().from('benefits').update(payload).eq('id', row.id)
              : sb().from('benefits').insert(payload).select('id').single();
            const { error } = await q; if (error) throw error;
            closeModal(); await loadAdminBenefits();
          }catch(err){ alert('Kaydedilemedi: ' + (err?.message||String(err))); }
        } finally {
          saveBtn.disabled = false;
          saveBtn.textContent = origSaveText;
          saving = false;
        }
      });
      (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
    }catch(e){ alert('Form açılamadı: ' + (e?.message||String(e))); }
  }

  // Lazy-load fontkit for pdf-lib embedding of custom fonts
  async function ensureFontkit(){
    if (window.fontkit) return window.fontkit;
    if (document.querySelector('script[data-fontkit]')){
      for (let i=0;i<10;i++){ await new Promise(r=> setTimeout(r,150)); if (window.fontkit) return window.fontkit; }
    }
    const cdns = [
      'vendor/fontkit/fontkit.umd.min.js',
      'https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@0.0.4/dist/fontkit.umd.min.js',
      'https://unpkg.com/@pdf-lib/fontkit@0.0.4/dist/fontkit.umd.min.js',
      'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/fontkit.umd.min.js'
    ];
    for (const url of cdns){
      try{
        const s = document.createElement('script'); s.src=url; s.async=true; s.defer=true; s.setAttribute('data-fontkit','1');
        const p = new Promise((resolve)=>{ s.onload=()=>resolve(window.fontkit||null); s.onerror=()=>resolve(null); });
        document.head.appendChild(s);
        const fk = await p; if (fk) return fk;
      }catch{}
    }
    return null;
  }

  async function openMembershipFormHtmlPreview(member){
    try{
      const plateCode = await getProvincePlateCodeByName(member?.work_province);
      $modalTitle().textContent = 'Üyelik Formu (Önizleme)';
      const form = $modalForm(); form.innerHTML = '';
      const tools = document.createElement('div');
      tools.style.display='flex'; tools.style.gap='8px'; tools.style.justifyContent='flex-end'; tools.style.marginBottom='8px';
      const btnPrint = document.createElement('button'); btnPrint.type='button'; btnPrint.className='btn btn-success'; btnPrint.textContent='Yazdır';
      const note = document.createElement('div'); note.className='muted'; note.style.marginRight='auto'; note.textContent='PDF kütüphanesi yüklenemedi. Yazdırılabilir görünüm gösteriliyor.';
      const grid = document.createElement('div');
      grid.style.display='grid'; grid.style.gridTemplateColumns='1fr 1fr'; grid.style.gap='8px'; grid.style.padding='12px'; grid.style.background='white'; grid.style.border='1px solid #e5e7eb'; grid.style.borderRadius='8px';
      const addRow = (label, value)=>{
        const l = document.createElement('div'); l.style.fontWeight='700'; l.style.color='#E03B3B'; l.textContent=label;
        const v = document.createElement('div'); v.style.whiteSpace='pre-wrap'; v.style.color='#0B3A60'; v.textContent=String(value||'');
        grid.appendChild(l); grid.appendChild(v);
      };
      const fmtDate = (d)=>{ try{ if(!d) return ''; const parts=String(d).split('-'); if(parts.length===3){ return `${parts[2]}.${parts[1]}.${parts[0]}`; } return String(d);}catch{return String(d||'');} };
      addRow('KURUMUN ADI', member?.institution_name);
      addRow('GÖREV YAPILAN BİRİMİN ADI', member?.work_unit);
      addRow('GÖREV YAPILAN BİRİMİN ADRESİ', member?.work_unit_address);
      addRow('İL ADI', member?.work_province);
      addRow('İL Kodu', plateCode);
      addRow('İLÇE ADI', member?.work_district);
      addRow('ADI', member?.first_name);
      addRow('SOYADI', member?.last_name);
      addRow('TC KİMLİK NO', member?.national_id);
      addRow('BABA ADI', member?.father_name);
      addRow('ANA ADI', member?.mother_name);
      addRow('DOĞUM TARİHİ', fmtDate(member?.birth_date));
      addRow('DOĞUM YERİ', member?.birth_place);
      addRow('CİNSİYETİ', member?.gender);
      addRow('ÖĞRENİM', member?.education);
      addRow('KURUM SİCİL', member?.corp_reg_no);
      addRow('KADRO ÜNVANI', member?.title);
      addRow('T.C.EMEKLİ SANDIĞI', member?.retirement_no);
      addRow('SOSYAL SİGORTALAR KURUMU', member?.ssk_no);
      addRow('E-POSTA', member?.email);
      addRow('CEP TEL', member?.phone);
      btnPrint.addEventListener('click', ()=>{
        try{
          const w = window.open('', '_blank');
          if (!w) return;
          const safeTitle = 'Üyelik Formu';
          w.document.open();
          w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${safeTitle}</title><style>body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;padding:16px} .grid{display:grid;grid-template-columns:1fr 1fr;gap:8px} .grid div{padding:6px 8px;border-bottom:1px dashed #e5e7eb} .grid div:nth-child(odd){font-weight:700;color:#E03B3B} .grid div:nth-child(even){color:#0B3A60}</style></head><body>`);
          w.document.write(`<h3 style="margin:0 0 12px">Üyelik Formu</h3><div class="grid">${grid.innerHTML}</div>`);
          w.document.write('</body></html>');
          w.document.close();
          w.focus(); w.print(); setTimeout(()=> w.close(), 400);
        }catch{}
      });
      const left = document.createElement('div'); left.style.display='flex'; left.style.alignItems='center'; left.style.gap='8px';
      left.appendChild(note);
      form.appendChild(left);
      tools.appendChild(btnPrint);
      form.appendChild(tools);
      form.appendChild(grid);
      try{ (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal())); }catch{}
    }catch{}
  }
  function closeModal(){
    const m = $modal();
    if (!m) return;
    m.style.display = 'none';
    editing = null;
  }
  try{ const closeBtn = qs('#modalClose'); if (closeBtn) closeBtn.addEventListener('click', closeModal); }catch{}
  // Also expose to window to avoid scope issues across handlers
  try{ window.openModal = openModal; window.closeModal = closeModal; }catch{}

  // Safe HTML escaper used in table renderers
  function escapeHtml(s){ const div=document.createElement('div'); div.textContent=String(s); return div.innerHTML; }

  let currentTab = 'news';
  let currentAdmin = { email:null, roles:[], allowed_tabs:[] };
  let editing = null; // { type: 'news'|'ann', id: string|null }
  let messagesById = new Map();
  let currentMsgCategory = 'all';
  let currentDocsSubTab = 'incoming';
  let incomingDocsById = new Map();
  let outgoingDocsById = new Map();

// Seçime inline stil uygula (ör. font-size)
// NOTE: Tek bir kez, IIFE içinde üst seviye scope’a ekleyin.
function applyInlineStyle(prop, val){
  const sel = window.getSelection(); 
  if (!sel || !sel.rangeCount) return;
  const range = sel.getRangeAt(0);
  if (range.collapsed){
    const span = document.createElement('span');
    span.style[prop] = val;
    // imleci span içinden çıkarmak için zero-width char bırak
    span.appendChild(document.createTextNode('\u200b'));
    range.insertNode(span);
    sel.removeAllRanges();
    const r = document.createRange();
    r.selectNodeContents(span);
    r.collapse(false);
    sel.addRange(r);
  } else {
    const span = document.createElement('span');
    span.style[prop] = val;
    try{ range.surroundContents(span); }
    catch { /* bazı karmaşık seçimlerde surround çalışmayabilir; burada no-op */ }
  }
}
  function showLogin(){
    $loginSection().hidden = false;
    $adminSection().hidden = true;
    $authStatus().textContent = 'Giriş yapınız';
  }

  // ========== USERS (RBAC) ==========
  async function loadAdminUsers(){
    const tbody = $usersTBody(); if (!tbody) return;
    tbody.innerHTML= '';
    try{
      const { data, error } = await sb().from('admin_users').select('email, roles, allowed_tabs').order('email');
      if (error) throw error;
      const tabNames = { news:'Haberler', ann:'Duyurular', msgs:'Mesajlar', pages:'Sayfalar', benefits:'Üyelere Özel', posters:'Afiş', reports:'Rapor', founders:'Kurucular', chairman:'Genel Başkan', members:'Üyeler', users:'Kullanıcılar', settings:'Ayarlar' };
      (data||[]).forEach(row=>{
        const tr=document.createElement('tr');
        const roles = Array.isArray(row.roles)? row.roles.join(', ') : '';
        const tabs = Array.isArray(row.allowed_tabs)
          ? row.allowed_tabs.map(t => tabNames[t] || t).join(', ')
          : '';
        tr.innerHTML = `<td>${escapeHtml(row.email|| '')}</td><td>${escapeHtml(roles)}</td><td>${escapeHtml(tabs)}</td><td class="actions"><button class="btn btn-warning" data-edit-user="${row.email}">Düzenle</button><button class="btn btn-danger" data-del-user="${row.email}">Sil</button></td>`;
        tbody.appendChild(tr);
      });
      wireUsersRowActions();
    }catch(e){ alert('Kullanıcılar yüklenemedi: ' + (e?.message||String(e))); }
  }

  function wireUsersRowActions(){
    const tbody=$usersTBody(); if(!tbody) return;
    tbody.querySelectorAll('button[data-edit-user]').forEach(btn=>{
      btn.addEventListener('click', async ()=>{
        const email = btn.getAttribute('data-edit-user');
        const { data } = await sb().from('admin_users').select('email, roles, allowed_tabs').eq('email', email).maybeSingle();
        openUserModal(data || { email, roles:[], allowed_tabs:[] });
      });
    });
    tbody.querySelectorAll('button[data-del-user]').forEach(btn=>{
      btn.addEventListener('click', async ()=>{
        const email = btn.getAttribute('data-del-user');
        if (!confirm('Silinsin mi?')) return;
        const { error } = await sb().from('admin_users').delete().eq('email', email);
        if (error) return alert('Silme hatası: ' + error.message);
        loadAdminUsers();
      });
    });
    const addBtn = qs('#newAdminUserBtn'); if (addBtn) addBtn.onclick = ()=> openUserModal({ email:'', roles:[], allowed_tabs:[] });
  }

  // ========== ADMIN USERS MODAL ==========
  function openUserModal(row){
    editing = { type:'user', id: row.email };
    $modalTitle().textContent = row.email ? 'Kullanıcıyı Düzenle' : 'Yeni Kullanıcı';
    $modalForm().innerHTML= '';
    const form = $modalForm();

    // Helpers for building labeled inputs/selects
    function inputEl(labelText, name, value, type='text'){
      const label = document.createElement('label');
      label.style.display='grid';
      label.style.gap='6px';
      label.innerHTML = `<span>${labelText}</span>`;
      const input = document.createElement('input');
      input.name = name;
      input.type = type;
      input.value = value || '';
      input.className = 'btn btn-outline';
      input.style.padding='8px';
      input.style.border='1px solid #e5e7eb';
      input.style.borderRadius='6px';
      label.appendChild(input);
      return label;
    }
    function selectEl(labelText, name, value, options){
      const label = document.createElement('label');
      label.style.display='grid';
      label.style.gap='6px';
      label.innerHTML = `<span>${labelText}</span>`;
      const sel = document.createElement('select');
      sel.name=name;
      sel.className='btn btn-outline';
      sel.style.padding='8px';
      sel.style.border='1px solid #e5e7eb';
      sel.style.borderRadius='6px';
      (options||[]).forEach(o=>{ const opt=document.createElement('option'); opt.value=o.v; opt.textContent=o.t; if(String(o.v)===String(value)) opt.selected=true; sel.appendChild(opt); });
      label.appendChild(sel);
      return label;
    }

    // Email
    form.appendChild(inputEl('E‑posta', 'email', row.email|| ''));
    form.appendChild(inputEl('Yeni Şifre', 'password1', '', 'password'));
    form.appendChild(inputEl('Yeni Şifre (Tekrar)', 'password2', '', 'password'));
    // Roles (at least superadmin is supported by applyTabPermissions)
    const rolesLbl = document.createElement('div'); rolesLbl.style.display='grid'; rolesLbl.style.gap='6px'; rolesLbl.innerHTML='<span>Roller</span>';
    const rolesWrap = document.createElement('div'); rolesWrap.style.display='flex'; rolesWrap.style.flexWrap='wrap'; rolesWrap.style.gap='10px';
    const roleOptions = [
      { v:'superadmin', t:'Süper Yönetici' },
      { v:'editor', t:'Editör' },
      { v:'viewer', t:'Görüntüleyici' }
    ];
    roleOptions.forEach(r=>{
      const lab=document.createElement('label'); lab.style.display='inline-flex'; lab.style.alignItems='center'; lab.style.gap='6px';
      const cb=document.createElement('input'); cb.type='checkbox'; cb.value=r.v; cb.checked=row.roles && row.roles.includes(r.v);
      lab.appendChild(cb); lab.appendChild(document.createTextNode(r.t)); rolesWrap.appendChild(lab);
    });
    rolesLbl.appendChild(rolesWrap); form.appendChild(rolesLbl);

    // Allowed tabs
    const tabsLbl = document.createElement('div'); tabsLbl.style.display='grid'; tabsLbl.style.gap='6px'; tabsLbl.innerHTML='<span>İzinli Sekmeler</span>';
    const tabsWrap = document.createElement('div'); tabsWrap.style.display='flex'; tabsWrap.style.flexWrap='wrap'; tabsWrap.style.gap='10px';
    const tabOptions = [
      { v:'news', t:'Haberler' },
      { v:'ann', t:'Duyurular' },
      { v:'msgs', t:'Mesajlar' },
      { v:'pages', t:'Sayfalar' },
      { v:'docs', t:'Evrak Modülü' },
      { v:'benefits', t:'Üyelere Özel' },
      { v:'posters', t:'Afiş' },
      { v:'reports', t:'Rapor' },
      { v:'founders', t:'Kurucular' },
      { v:'chairman', t:'Genel Başkan' },
      { v:'members', t:'Üyeler' },
      { v:'users', t:'Kullanıcılar' },
      { v:'settings', t:'Ayarlar' }
    ];
    tabOptions.forEach(t=>{
      const lab=document.createElement('label'); lab.style.display='inline-flex'; lab.style.alignItems='center'; lab.style.gap='6px';
      const cb=document.createElement('input'); cb.type='checkbox'; cb.value=t.v; cb.checked=row.allowed_tabs && row.allowed_tabs.includes(t.v);
      lab.appendChild(cb); lab.appendChild(document.createTextNode(t.t)); tabsWrap.appendChild(lab);
    });
    tabsLbl.appendChild(tabsWrap); form.appendChild(tabsLbl);

    // Actions
    const actions=document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
    const saveBtn=document.createElement('button'); saveBtn.className='btn btn-success'; saveBtn.textContent='Kaydet';
    const cancelBtn=document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
    actions.appendChild(saveBtn); actions.appendChild(cancelBtn); form.appendChild(actions);

    cancelBtn.addEventListener('click', (e)=>{ e.preventDefault(); closeModal(); });
    saveBtn.addEventListener('click', async (e)=>{
      e.preventDefault();
      const email = String(qs('input[name="email"]', form).value|| '').trim();
      if (!email) return alert('E‑posta gerekli');
      const pass1 = String(qs('input[name="password1"]', form)?.value|| '').trim();
      const pass2 = String(qs('input[name="password2"]', form)?.value|| '').trim();
      if ((pass1 || pass2) && pass1 !== pass2) return alert("Şifreler aynı değil");
      if (pass1 && pass1.length < 8) return alert("Şifre en az 8 karakter olmalı");
      const selectedRoles = Array.from(rolesWrap.querySelectorAll('input[type="checkbox"]'))
        .filter(i=>i.checked).map(i=>i.value);
      const selectedTabs = Array.from(tabsWrap.querySelectorAll('input[type="checkbox"]'))
        .filter(i=>i.checked).map(i=>i.value);
      try{
        const payload = { email, roles: selectedRoles, allowed_tabs: selectedTabs };
        const { error } = await sb().from('admin_users').upsert(payload, { onConflict:'email' });
        if (error) throw error;
        if (pass1) {
          try {
            const { data: sess } = await sb().auth.getSession();
            const token = sess?.session?.access_token || '';
            const headers = token ? { Authorization: `Bearer ${token}` } : {};
            const { data: resp, error: funcErr } = await sb().functions.invoke("auth_admin", {
              body: { action: "upsert_user", email, password: pass1 },
              headers
            });
            if (funcErr) throw funcErr;
            if (resp && resp.error) throw new Error(resp.error);
          } catch (e) {
            return alert("Şifre ayarlanamadı: " + (e?.message || String(e)));
          }
        }
        closeModal();
        if (typeof loadAdminUsers === 'function') await loadAdminUsers();
      }catch(err){ alert('Kaydedilemedi: ' + (err?.message||String(err))); }
    });

    // Open modal
    (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
  }

  // ========== MEMBERS ==========
  async function loadAdminMembers(){
    const tbody = $membersTBody(); if (!tbody) return;
    // Ensure search toolbar exists
    try{
      const table = tbody.closest('table');
      const host = table && table.parentElement ? table.parentElement : tbody.parentElement;
      if (host && !qs('#membersSearchBar')){
        const bar = document.createElement('div');
        bar.id = 'membersSearchBar';
        bar.style.display='flex'; bar.style.gap='8px'; bar.style.alignItems='center'; bar.style.margin='8px 0';
        const input = document.createElement('input'); input.type='text'; input.placeholder='İsim veya TC ile ara'; input.id='membersSearchInput'; input.className='btn btn-outline'; input.style.flex='1';
        const btn = document.createElement('button'); btn.id='membersSearchBtn'; btn.textContent='Ara'; btn.className='btn';
        bar.appendChild(input); bar.appendChild(btn);
        host.insertBefore(bar, table || host.firstChild);
        // Events
        const reload = ()=> loadAdminMembers();
        let tmr; input.addEventListener('keyup', (e)=>{ clearTimeout(tmr); tmr = setTimeout(reload, 300); if (e.key==='Enter') reload(); });
        btn.addEventListener('click', reload);
      }
      // Wire pending/tümü filters once
      const filt = qs('#membersFilters');
      if (filt && !filt.dataset.wired){
        filt.dataset.wired = '1';
        filt.addEventListener('click', (ev)=>{
          const b = ev.target && ev.target.closest ? ev.target.closest('button[data-mem-filter]') : null;
          if (!b) return;
          Array.from(filt.querySelectorAll('button')).forEach(x=> x.classList.remove('active'));
          b.classList.add('active');
          loadAdminMembers();
        });
      }
    }catch{}

    tbody.innerHTML= '';
    try{
      const term = (qs('#membersSearchInput')?.value || '').trim();
      const currentFilter = (qs('#membersFilters .btn.active')?.getAttribute('data-mem-filter')) || 'all';
      let q = sb().from('members').select('id, member_no, first_name, last_name, national_id, email, phone, status, join_date').order('member_no', { ascending:true, nullsFirst:true });
      if (term){
        const like = `%${term}%`;
        q = q.or(
          `first_name.ilike.${like},last_name.ilike.${like},national_id.ilike.${like}`
        );
      }
      // Filters
      if (currentFilter === 'pending'){
        // Pending by status
        q = q.eq('status', 'pending');
      } else if (currentFilter === 'active'){
        // Approved active members only
        q = q.eq('status', 'active').not('member_no', 'is', null);
      } else if (currentFilter === 'passive'){
        // Approved passive members only
        q = q.eq('status', 'passive').not('member_no', 'is', null);
      }
      const { data, error } = await q;
      if (error) throw error;
      (data||[]).forEach(row=>{
        const tr=document.createElement('tr');
        const name = `${escapeHtml(row.first_name|| '')} ${escapeHtml(row.last_name|| '')}`.trim();
        const status = row.status || 'active';
        const statusTr = status === 'active' ? 'Aktif' : (status === 'pending' ? 'Onay bekliyor' : 'Pasif');
        const toggleText = status === 'active' ? 'Pasife Al' : 'Aktife Al';
        const jd = row.join_date ? new Date(row.join_date).toLocaleDateString('tr-TR') : '-';
        const isPending = (status === 'pending');
        tr.innerHTML = `
          <td>${isPending ? '-' : (row.member_no||'-')}</td>
          <td>${name}<div class="muted" style="font-size:12px">${escapeHtml(row.national_id|| '')}</div></td>
          <td>${escapeHtml(row.email|| '')}</td>
          <td>${escapeHtml(row.phone|| '')}</td>
          <td>${escapeHtml(statusTr)}</td>
          <td>${jd}</td>
          <td class="actions">
            <div class="row-1">
              <button class="btn btn-warning" data-edit-mem="${row.id}">Düzenle</button>
              <button class="btn btn-primary" data-printform-mem="${row.id}">Üyelik Formu Yazdır</button>
            </div>
            <div class="row-2" style="margin-top:6px">
              ${!isPending ? `<button class="btn btn-${status==='active'?'danger':'success'}" data-toggle-mem="${row.id}" data-status="${status}">${toggleText}</button>` : ''}
              ${isPending ? `<button class="btn btn-success" data-approve-mem="${row.id}">Üyeliği Onayla</button>` : ''}
              <button class="btn btn-primary" data-idcard-mem="${row.id}">Kimlik Göster</button>
            </div>
          </td>`;
        tbody.appendChild(tr);
      });
      wireMembersRowActions();
    }catch(e){ alert('Üyeler yüklenemedi: ' + (e?.message||String(e))); }
  }

  function wireMembersRowActions(){
    const tbody=$membersTBody(); if(!tbody) return;
    tbody.querySelectorAll('button[data-edit-mem]').forEach(btn=>{
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-edit-mem');
        const { data } = await sb().from('members').select('*').eq('id', id).maybeSingle();
        openMemberModal(data || { id:null, status:'active' });
      });
    });
    tbody.querySelectorAll('button[data-approve-mem]').forEach(btn=>{
      if (btn.dataset.wired) return; btn.dataset.wired='1';
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-approve-mem');
        try{
          if (!confirm('Bu üyeliği onaylamak istediğinize emin misiniz?')) return;
          const memberNo = await nextMemberNo();
          const today = new Date().toISOString().slice(0,10);
          const { error } = await sb().from('members').update({ member_no: memberNo, join_date: today, status:'active' }).eq('id', id);
          if (error) throw error;
          loadAdminMembers();
        }catch(err){ alert('Onaylama başarısız: ' + (err?.message||String(err))); }
      });
    });
    tbody.querySelectorAll('button[data-toggle-mem]').forEach(btn=>{
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-toggle-mem');
        const cur = btn.getAttribute('data-status') || 'active';
        try{
          if (cur === 'active'){
            if (!confirm('Üyelik pasife alınsın mı?')) return;
            const today = new Date().toISOString().slice(0,10);
            const { error } = await sb().from('members').update({ status:'passive', leave_date: today }).eq('id', id);
            if (error) throw error;
          } else {
            if (!confirm('Üyelik aktifleştirilsin mi?')) return;
            const { error } = await sb().from('members').update({ status:'active', leave_date: null }).eq('id', id);
            if (error) throw error;
          }
          loadAdminMembers();
        }catch(err){ alert('İşlem başarısız: ' + (err?.message||String(err))); }
      });
    });
    tbody.querySelectorAll('button[data-idcard-mem]').forEach(btn=>{
      if (btn.dataset.wired) return;
      btn.dataset.wired = '1';
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-idcard-mem');
        try{
          const { data } = await sb().from('members').select('*').eq('id', id).maybeSingle();
          if (!data) return alert('Üye bulunamadı');
          openIdCardPreview(data);
        }catch(e){ alert('Kimlik gösterilemedi: ' + (e?.message||String(e))); }
      });
    });
    tbody.querySelectorAll('button[data-printform-mem]').forEach(btn=>{
      if (btn.dataset.wired) return; btn.dataset.wired='1';
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-printform-mem');
        try{
          const { data } = await sb().from('members').select('*').eq('id', id).maybeSingle();
          if (!data) return alert('Üye bulunamadı');
          await generateMembershipFormPDF(data);
        }catch(e){ alert('Form yazdırılamadı: ' + (e?.message||String(e))); }
      });
    });
    const addBtn = qs('#newMemberBtn'); if (addBtn) addBtn.onclick = ()=> openMemberModal({ id:null, status:'active' });
  }

  async function nextMemberNo(){
    try{
      const { data, error } = await sb()
        .from('members')
        .select('member_no,status')
        .not('member_no', 'is', null)
        .neq('status', 'pending')
        .order('member_no', { ascending:false, nullsFirst:false })
        .limit(1);
      if (error) throw error;
      const raw = (data && data[0] && data[0].member_no) != null ? Number(data[0].member_no) : 0;
      const max = Number.isFinite(raw) ? raw : 0;
      return max + 1;
    }catch{ return 1; }
  }

  // Lazy-load pdf-lib if it isn't already available
  async function ensurePdfLib(){
    if (window.PDFLib && window.PDFLib.PDFDocument) return window.PDFLib;
    if (window.pdfLib && window.pdfLib.PDFDocument) return window.pdfLib;
    // prevent duplicate loads
    if (document.querySelector('script[data-pdf-lib]')){
      // wait for an existing load
      for (let i=0;i<10;i++){
        await new Promise(r=> setTimeout(r, 150));
        if (window.PDFLib?.PDFDocument || window.pdfLib?.PDFDocument) return window.PDFLib || window.pdfLib;
      }
    }
    const cdns = [
      // Prefer local vendor first (works offline and bypasses CSP/CDN blocks)
      'vendor/pdf-lib/pdf-lib.min.js',
      'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js',
      'https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js',
      'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js'
    ];
    for (const url of cdns){
      try{
        const s = document.createElement('script');
        s.src = url; s.async = true; s.defer = true; s.setAttribute('data-pdf-lib', '1');
        const p = new Promise((resolve)=>{ s.onload = ()=> resolve(window.PDFLib || window.pdfLib || null); s.onerror = ()=> resolve(null); });
        document.head.appendChild(s);
        const mod = await p;
        if (mod && mod.PDFDocument) return mod;
      }catch{}
    }
    return null;
  }

  // Province plate code lookup with simple cache
  async function getProvincePlateCodeByName(name){
    const key = String(name||'').trim().toLowerCase();
    if (!key) return '';
    try{
      window.__provMap = window.__provMap || null;
      if (!window.__provMap){
        // Try to load from localStorage cache first
        try{ const raw = localStorage.getItem('ilkesen_prov_map_v1'); if (raw) window.__provMap = JSON.parse(raw); }catch{}
      }
      if (!window.__provMap){
        const { data, error } = await sb().from('provinces').select('name,plate_code');
        if (!error && data){
          window.__provMap = {};
          data.forEach(p=>{ const k=String(p.name||'').trim().toLowerCase(); if (k) window.__provMap[k] = p.plate_code; });
          try{ localStorage.setItem('ilkesen_prov_map_v1', JSON.stringify(window.__provMap)); }catch{}
        }
      }
      if (window.__provMap && key in window.__provMap){
        const v = window.__provMap[key];
        return v != null ? String(v) : '';
      }
      // Last attempt: single query ilike
      try{
        const { data } = await sb().from('provinces').select('plate_code,name').ilike('name', key).maybeSingle();
        if (data && (data.plate_code!=null)){
          window.__provMap = window.__provMap || {}; window.__provMap[key] = data.plate_code;
          try{ localStorage.setItem('ilkesen_prov_map_v1', JSON.stringify(window.__provMap)); }catch{}
          return String(data.plate_code);
        }
      }catch{}
    }catch{}
    return '';
  }

  async function generateMembershipFormPDF(member){
    try{
      const PDFLib = await ensurePdfLib();
      if (!PDFLib || !PDFLib.PDFDocument) { await openMembershipFormHtmlPreview(member); return; }
      const { PDFDocument, StandardFonts } = PDFLib;
      // Resolve province plate code using cached list or DB
      let plateCode = '';
      try{ plateCode = await getProvincePlateCodeByName(member.work_province); }catch{}
      // Prepare mapping
      const fmtDate = (d)=>{ try{ if(!d) return ''; const parts=String(d).split('-'); if(parts.length===3){ return `${parts[2]}.${parts[1]}.${parts[0]}`; } return String(d);}catch{return String(d||'');} };
      const values = {
        'KURUMUN ADI': member.institution_name||'',
        'GÖREV YAPILAN BİRİMİN ADI': member.work_unit||'',
        'GÖREV YAPILAN BİRİMİN ADRESİ': member.work_unit_address||'',
        'İL ADI': member.work_province||'',
        'iL ADI': member.work_province||'',
        'İL Kodu': plateCode||'',
        'iL Kodu': plateCode||'',
        'İLÇE ADI': member.work_district||'',
        'ADI': member.first_name||'',
        'SOYADI': member.last_name||'',
        'TC KİMLİK NO': member.national_id||'',
        'BABA ADI': member.father_name||'',
        'ANA ADI': member.mother_name||'',
        'DOĞUM TARİHİ': fmtDate(member.birth_date),
        'DOĞUM YERİ': member.birth_place||'',
        'CİNSİYETİ': member.gender||'',
        'ÖĞRENİM': member.education||'',
        'KURUM SİCİL': member.corp_reg_no||'',
        'KADRO ÜNVANI': member.title||'',
        'T.C.EMEKLİ SANDIĞI': member.retirement_no||'',
        'SOSYAL SİGORTALAR KURUMU': member.ssk_no||'',
        'E-POSTA': member.email||'',
        'CEP TEL': member.phone||'',
      };
      // Load template with fallbacks for non-ascii filename encodings
      async function fetchTpl(){
        const urls = ['assets/üyelik.pdf', 'assets/%C3%BCyelik.pdf', 'assets/uyelik.pdf'];
        for (const u of urls){ try{ const r = await fetch(u, { cache:'no-store' }); if (r.ok) return await r.arrayBuffer(); }catch{} }
        throw new Error('Şablon PDF yüklenemedi');
      }
      const tplBytes = await fetchTpl();
      const pdfDoc = await PDFDocument.load(tplBytes);
      // Optional custom font for Turkish
      let customFont = null;
      try{
        if (window.fontkit && pdfDoc.registerFontkit) { pdfDoc.registerFontkit(window.fontkit); }
        const fontResp = await fetch('assets/fonts/Roboto-Regular.ttf');
        if (fontResp.ok){
          const fontBytes = await fontResp.arrayBuffer();
          customFont = await pdfDoc.embedFont(fontBytes);
        }
      }catch{}
      const form = pdfDoc.getForm();
      const fields = form.getFields();
      const norm = (s)=> String(s||'').normalize('NFKD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Za-z0-9ÇĞİÖŞÜçğıöşü]/g,'').toLocaleUpperCase('tr-TR');
      const catalog = (fields||[]).map(f=>({ n: norm(f.getName()), name: f.getName() }));
      function setFieldSmart(label, val){
        const v = String(val ?? '');
        // 1) direct
        try{ const tf = form.getTextField(label); tf.setText(v); return true; }catch{}
        const key = norm(label);
        // 2) match normalized equal
        let hit = catalog.find(c=> c.n === key);
        if (hit){ try{ form.getTextField(hit.name).setText(v); return true; }catch{} }
        // 3) contains either way
        hit = catalog.find(c=> c.n.includes(key) || key.includes(c.n));
        if (hit){ try{ form.getTextField(hit.name).setText(v); return true; }catch{} }
        return false;
      }
      Object.entries(values).forEach(([k,v])=>{ try{ setFieldSmart(k, v); }catch{} });
      try{
        if (!customFont){
          try{ await ensureFontkit(); }catch{}
        }
        if (window.fontkit && pdfDoc.registerFontkit){ try{ pdfDoc.registerFontkit(window.fontkit); }catch{} }
        if (customFont){ form.updateFieldAppearances(customFont); }
        else { const helv = await pdfDoc.embedFont(StandardFonts.Helvetica); form.updateFieldAppearances(helv); }
      }catch{}
      try{ form.flatten(); }catch{}
      // Fallback for non-form template: use coordinate map (with calibration support)
      try{
        const noFields = !form || (form.getFields && form.getFields().length === 0);
        if (noFields){
          const pages = pdfDoc.getPages(); const page = pages[0];
          const { width, height } = page.getSize();
          const drawerFont = customFont || await pdfDoc.embedFont(StandardFonts.Helvetica);
          const defaultPageFrame = { x:0.29034749329320253, y:0.08080627461685956, w:0.40250965250965254, h:0.9048178200970485 };
          const defaultCoords = {"kurum_adi":{"x":0.32115696712360275,"y":0.3054318129533846},"gorev_birim_adi":{"x":0.3192376374534848,"y":0.35759530361587766},"gorev_birim_adresi":{"x":0.3192376374534848,"y":0.38400000000000006},"il_adi":{"x":0.32691495613395666,"y":0.5119992359768569},"il_kodu":{"x":0.6800716154356627,"y":0.5099126963503573},"ilce_adi":{"x":0.8182633516841563,"y":0.5099126963503573},"adi":{"x":0.3000443407523052,"y":0.6010249532392351},"soyadi":{"x":0.29812501108218714,"y":0.6364961268897303},"tc":{"x":0.2942863517419513,"y":0.6858775676187977},"baba":{"x":0.2962056814120692,"y":0.7195404103920329},"ana":{"x":0.726135527518494,"y":0.7216269500185325},"dogum_tarihi":{"x":0.2942863517419513,"y":0.7612712029220272},"dogum_yeri":{"x":0.7338128461989657,"y":0.7612712029220272},"kurum_sicil":{"x":0.3000443407523052,"y":0.8669892212773691},"unvan":{"x":0.3019636704224231,"y":0.9042687175968724},"eposta":{"x":0.6570396593942471,"y":0.8468193276085157},"cep":{"x":0.20983584625676072,"y":0.882290501259011},"sgk_sicil":{"x":0.2040778572464068,"y":0.8452892059779652},"uye_kayit_no":{"x":0.36146289019608013,"y":0.7091077122595342},"kabul_tarihi":{"x":0.3461082528351363,"y":0.7918737947233792},"cinsiyet_erkek":{"x":0.4977352967744557,"y":0.1732843298067924},"cinsiyet_kadin":{"x":0.8643272637669875,"y":0.17537086943329214},"ogrenim_ilkogretim":{"x":0.37489819788690587,"y":0.21292858271028706},"ogrenim_lise":{"x":0.5860244615998822,"y":0.2150151223367868},"ogrenim_yuksekokul":{"x":0.9161491648601727,"y":0.2150151223367868},"sgk_emekli":{"x":0.4919773077641018,"y":0.3881979113362636},"sgk_ssk":{"x":0.4900579780939839,"y":0.4236690849867588}};
          const defaultStyles = {"kurum_adi":{"dx":-0.006000000000000001,"dy":-0.05500000000000001,"size":9.5},"gorev_birim_adi":{"dx":-0.003,"dy":-0.07900000000000003,"size":9.5},"gorev_birim_adresi":{"dx":-0.003,"dy":-0.07200000000000003,"size":9.5},"il_adi":{"dx":0.07500000000000002,"dy":-0.14500000000000007,"size":9.5},"il_kodu":{"dx":-0.026999999999999996,"dy":-0.14200000000000007,"size":9.5},"ilce_adi":{"dx":-0.042,"dy":-0.14200000000000007,"size":9.5},"adi":{"dx":0.012,"dy":-0.1780000000000001,"size":9.5},"soyadi":{"dx":0.012,"dy":-0.19300000000000012,"size":9.5},"tc":{"dx":0.015,"dy":-0.21700000000000014,"size":9.5},"baba":{"dx":0.015,"dy":-0.22600000000000015,"size":9.5},"ana":{"dx":-0.018,"dy":-0.22900000000000015,"size":9.5},"dogum_tarihi":{"dx":0.018,"dy":-0.24700000000000016,"size":9.5},"dogum_yeri":{"dx":-0.026999999999999996,"dy":-0.24700000000000016,"size":9.5},"kurum_sicil":{"dx":0.012,"dy":-0.2890000000000002,"size":9.5},"unvan":{"dx":0.009000000000000001,"dy":-0.3070000000000002,"size":9.5},"eposta":{"dx":-0.045000000000000005,"dy":-0.12100000000000005,"size":9.5},"cep":{"dx":0,"dy":-0.13300000000000006,"size":9.5},"sgk_sicil":{"dx":0.24900000000000017,"dy":-0.19300000000000012,"size":9.5},"uye_kayit_no":{"dx":0.023999999999999997,"dy":0.11900000000000005,"size":16},"kabul_tarihi":{"dx":0.012,"dy":0.08600000000000003,"size":9.5},"cinsiyet_erkek":{"dx":-0.008999999999999994,"dy":0.3630000000000002,"size":10.5},"ogrenim_yuksekokul":{"dx":-0.06000000000000002,"dy":0.3420000000000002,"size":10},"cinsiyet_kadin":{"dx":-0.05400000000000001,"dy":0.36000000000000026,"size":10},"ogrenim_ilkogretim":{"dx":0.003,"dy":0.34500000000000025,"size":10},"ogrenim_lise":{"dx":-0.020999999999999998,"dy":0.34200000000000025,"size":10},"sgk_ssk":{"dx":0,"dy":-0.1680000000000001,"size":10}};
          let coords = JSON.parse(JSON.stringify(defaultCoords));
          try{ const raw = localStorage.getItem('ilkesen_pdf_coords_v1'); if (raw){ const m = JSON.parse(raw)||{}; Object.assign(coords, m); } }catch{}
          let styles = JSON.parse(JSON.stringify(defaultStyles));
          try{ const raw = localStorage.getItem('ilkesen_pdf_styles_v1'); if (raw){ const m = JSON.parse(raw)||{}; Object.assign(styles, m); } }catch{}
          const needCal = false;
          if (needCal){
            const tplUrl = URL.createObjectURL(new Blob([tplBytes], { type:'application/pdf' }));
            openPdfPreview('Üyelik Formu Kalibrasyon', tplUrl, member);
            return;
          }
          function drawTxt(key, text, sz){
            const c = coords[key]; if (!c) return; const st = styles[key]||{};
            const dx = Number.isFinite(st.dx) ? st.dx : 0; const dy = Number.isFinite(st.dy) ? st.dy : 0; const fs = Number.isFinite(st.size) ? st.size : (sz||10);
            const tx = Math.max(8, (c.x + dx) * width);
            const ty = Math.max(8, height - ((c.y + dy) * height));
            page.drawText(String(text||''), { x: tx, y: ty, size: fs, font: drawerFont });
          }
          function markX(key, sz){
            const c = coords[key]; if (!c) return; const st = styles[key]||{};
            const dx = Number.isFinite(st.dx) ? st.dx : 0; const dy = Number.isFinite(st.dy) ? st.dy : 0; const fs = Number.isFinite(st.size) ? st.size : (sz||12);
            const tx = Math.max(6, (c.x + dx) * width); const ty = Math.max(6, height - ((c.y + dy) * height));
            page.drawText('X', { x: tx, y: ty, size: fs, font: drawerFont });
          }
          drawTxt('kurum_adi', values['KURUMUN ADI']);
          drawTxt('gorev_birim_adi', values['GÖREV YAPILAN BİRİMİN ADI']);
          drawTxt('gorev_birim_adresi', values['GÖREV YAPILAN BİRİMİN ADRESİ']);
          drawTxt('il_adi', values['İL ADI']);
          drawTxt('il_kodu', values['İL Kodu']);
          drawTxt('ilce_adi', values['İLÇE ADI']);
          drawTxt('adi', values['ADI']);
          drawTxt('soyadi', values['SOYADI']);
          drawTxt('tc', values['TC KİMLİK NO']);
          drawTxt('baba', values['BABA ADI']);
          drawTxt('ana', values['ANA ADI']);
          drawTxt('dogum_tarihi', values['DOĞUM TARİHİ']);
          drawTxt('dogum_yeri', values['DOĞUM YERİ']);
          drawTxt('kurum_sicil', values['KURUM SİCİL']);
          drawTxt('unvan', values['KADRO ÜNVANI']);
          drawTxt('eposta', values['E-POSTA']);
          drawTxt('cep', values['CEP TEL']);
          // Optional: admin tarafında mevcutsa
          drawTxt('uye_kayit_no', member.member_no || '');
          let debugMarks = false; try{ debugMarks = localStorage.getItem('ilkesen_pdf_debug_marks_v1') === 'all'; }catch{}
          if (debugMarks){
            markX('cinsiyet_erkek'); markX('cinsiyet_kadin');
            markX('ogrenim_ilkogretim'); markX('ogrenim_lise'); markX('ogrenim_yuksekokul');
          } else {
            let gk = '';
            try{
              const forceG = localStorage.getItem('ilkesen_pdf_force_gender_v1');
              if (forceG==='erkek') gk = 'cinsiyet_erkek';
              else if (forceG==='kadin') gk = 'cinsiyet_kadin';
              else if (forceG==='none') gk = '';
            }catch{}
            if (!gk){
              const g = String(member.gender||'').toLowerCase();
              if (g.includes('erk') || g==='male') gk = 'cinsiyet_erkek';
              else if (g.includes('kad') || g==='female') gk = 'cinsiyet_kadin';
            }
            if (gk) markX(gk);
            function eduKey(v){
              const s = String(v||'').toLowerCase();
              if (!s) return '';
              if (s.includes('ilkö') || s.includes('ilko') || s.includes('ilkokul') || s.includes('primary')) return 'ogrenim_ilkogretim';
              if (s.includes('lise') || s.includes('ortaöğ') || s.includes('ortaog')) return 'ogrenim_lise';
              if (s.includes('yüksek') || s.includes('yuksek') || s.includes('univers') || s.includes('ünivers') || s.includes('uni') || s.includes('yuksekokul') || s.includes('yüksekokul') || s.includes('lisans') || s.includes('onlisans') || s.includes('önlisans')) return 'ogrenim_yuksekokul';
              const only = s.replace(/[^0-9]/g,'');
              if (only==='1') return 'ogrenim_ilkogretim';
              if (only==='2') return 'ogrenim_lise';
              if (only==='3') return 'ogrenim_yuksekokul';
              return '';
            }
            let ek = '';
            try{
              const force = localStorage.getItem('ilkesen_pdf_force_edu_v1');
              if (force==='ilkogretim') ek = 'ogrenim_ilkogretim';
              else if (force==='lise') ek = 'ogrenim_lise';
              else if (force==='yuksekokul') ek = 'ogrenim_yuksekokul';
            }catch{}
            if (!ek){ ek = eduKey(member.education); }
            if (ek) markX(ek);
          }
          const sgkNo = member.retirement_no || member.ssk_no || '';
          drawTxt('sgk_sicil', sgkNo);
        }
      }catch{}
      const pdfBytes = await pdfDoc.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      // Open directly in a new tab and trigger print
      try{
        const w = window.open(url, '_blank');
        if (w) {
          try{ w.focus(); }catch{}
          setTimeout(()=>{ try{ w.print(); }catch{} }, 600);
        } else {
          alert('Yeni sekme engellendi. Lütfen bu site için açılır pencerelere izin verin ve tekrar deneyin.');
        }
      }catch{}
    }catch(err){ alert('Form oluşturulamadı: ' + (err?.message||String(err))); }
  }

  function openPdfPreview(title, url, member){
    try{
      $modalTitle().textContent = title || 'PDF';
      const form = $modalForm(); form.innerHTML = '';
      const tools = document.createElement('div');
      tools.style.display='flex'; tools.style.gap='8px'; tools.style.justifyContent='flex-end'; tools.style.marginBottom='8px';
      const btnPrint = document.createElement('button'); btnPrint.type='button'; btnPrint.className='btn btn-success'; btnPrint.textContent='Yazdır';
      const btnDownload = document.createElement('button'); btnDownload.type='button'; btnDownload.className='btn btn-primary'; btnDownload.textContent='İndir';
      const btnOpen = document.createElement('button'); btnOpen.type='button'; btnOpen.className='btn btn-warning'; btnOpen.textContent='Yeni Sekmede Aç';
      const btnCal = document.createElement('button'); btnCal.type='button'; btnCal.className='btn btn-secondary'; btnCal.textContent='Kalibrasyon';
      const btnFull = document.createElement('button'); btnFull.type='button'; btnFull.className='btn'; btnFull.textContent='Tam Ekran';
      const btnZoomOut = document.createElement('button'); btnZoomOut.type='button'; btnZoomOut.className='btn'; btnZoomOut.textContent='-';
      const btnZoomIn = document.createElement('button'); btnZoomIn.type='button'; btnZoomIn.className='btn'; btnZoomIn.textContent='+';
      const btnScroll = document.createElement('button'); btnScroll.type='button'; btnScroll.className='btn'; btnScroll.textContent='Kaydırma Modu';
      const btnCorners = document.createElement('button'); btnCorners.type='button'; btnCorners.className='btn'; btnCorners.textContent='Sayfa Köşeleri';
      const btnReset = document.createElement('button'); btnReset.type='button'; btnReset.className='btn btn-danger'; btnReset.textContent='Sıfırla';
      const btnTune = document.createElement('button'); btnTune.type='button'; btnTune.className='btn'; btnTune.textContent='İnce Ayar';
      const btnMarks = document.createElement('button'); btnMarks.type='button'; btnMarks.className='btn';
      btnMarks.textContent = (localStorage.getItem('ilkesen_pdf_debug_marks_v1')==='all') ? 'İşaret Testi: Açık' : 'İşaret Testi';
      const btnEdu = document.createElement('button'); btnEdu.type='button'; btnEdu.className='btn'; btnEdu.textContent='Öğrenim Seç';
      const btnGender = document.createElement('button'); btnGender.type='button'; btnGender.className='btn'; btnGender.textContent='Cinsiyet Seç';
      const wrap = document.createElement('div');
      wrap.style.position='relative';
      wrap.style.transformOrigin = 'top left';
      let __scale = 1;
      const iframe = document.createElement('iframe');
      iframe.src = url; iframe.style.width='100%'; iframe.style.height='80vh'; iframe.style.border='1px solid #e5e7eb';
      const overlay = document.createElement('div');
      overlay.style.position='absolute'; overlay.style.left='0'; overlay.style.top='0'; overlay.style.right='0'; overlay.style.bottom='0'; overlay.style.display='none'; overlay.style.cursor='crosshair';
      overlay.style.background='transparent';
      wrap.appendChild(iframe); wrap.appendChild(overlay);
      const scroller = document.createElement('div');
      scroller.style.position='relative'; scroller.style.height='80vh'; scroller.style.overflow='auto';
      scroller.appendChild(wrap);
      btnPrint.addEventListener('click', ()=>{
        try{ if (iframe.contentWindow) { iframe.contentWindow.focus(); iframe.contentWindow.print(); return; } }catch{}
        try{ const w = window.open(url, '_blank'); if (w) { w.focus(); w.print(); } }catch{}
      });
      btnDownload.addEventListener('click', ()=>{
        const a = document.createElement('a');
        const safeName = `${String(member?.first_name||'').trim()}_${String(member?.last_name||'').trim()}`.replace(/[^A-Za-z0-9_çğıöşüİıÇĞÖŞÜ-]/g,'_');
        a.href = url; a.download = `uyelik_formu_${safeName||member?.national_id||member?.id||''}.pdf`;
        document.body.appendChild(a); a.click(); setTimeout(()=> a.remove(), 200);
      });
      btnOpen.addEventListener('click', ()=>{
        try{ const w = window.open(url, '_blank'); if (w) { w.focus(); setTimeout(()=>{ try{ w.print(); }catch{} }, 600); } }catch{}
      });
      let calCtx = null;
      btnCal.addEventListener('click', ()=>{
        try{
          const seq = [
            ['kurum_adi','KURUMUN ADI'],
            ['gorev_birim_adi','GÖREV YAPILAN BİRİMİN ADI'],
            ['gorev_birim_adresi','GÖREV YAPILAN BİRİMİN ADRESİ'],
            ['il_adi','İL ADI'],
            ['il_kodu','İL Kodu'],
            ['ilce_adi','İLÇE ADI'],
            ['adi','ADI'],
            ['soyadi','SOYADI'],
            ['tc','TC KİMLİK NO'],
            ['baba','BABA ADI'],
            ['ana','ANA ADI'],
            ['dogum_tarihi','DOĞUM TARİHİ'],
            ['dogum_yeri','DOĞUM YERİ'],
            ['kurum_sicil','KURUM SİCİL'],
            ['unvan','KADRO ÜNVANI'],
            ['eposta','E-POSTA'],
            ['cep','CEP TEL'],
            ['sgk_sicil','SGK SİCİL NO'],
            ['uye_kayit_no','ÜYE KAYIT NUMARASI'],
            ['kabul_tarihi','ÜYELİĞE KABUL TARİHİ'],
            ['cinsiyet_erkek','CİNSİYET ERKEK kutusu'],
            ['cinsiyet_kadin','CİNSİYET KADIN kutusu'],
            ['ogrenim_ilkogretim','ÖĞRENİM İLKÖĞRETİM kutusu'],
            ['ogrenim_lise','ÖĞRENİM LİSE kutusu'],
            ['ogrenim_yuksekokul','ÖĞRENİM YÜKSEK OKUL kutusu']
          ];
          let i = 0;
          let coords = {"kurum_adi":{"x":0.32115696712360275,"y":0.3054318129533846},"gorev_birim_adi":{"x":0.3192376374534848,"y":0.35759530361587766},"il_adi":{"x":0.32691495613395666,"y":0.5119992359768569},"il_kodu":{"x":0.6800716154356627,"y":0.5099126963503573},"ilce_adi":{"x":0.8182633516841563,"y":0.5099126963503573},"adi":{"x":0.3000443407523052,"y":0.6010249532392351},"soyadi":{"x":0.29812501108218714,"y":0.6364961268897303},"tc":{"x":0.2942863517419513,"y":0.6858775676187977},"baba":{"x":0.2962056814120692,"y":0.7195404103920329},"ana":{"x":0.726135527518494,"y":0.7216269500185325},"dogum_tarihi":{"x":0.2942863517419513,"y":0.7612712029220272},"dogum_yeri":{"x":0.7338128461989657,"y":0.7612712029220272},"kurum_sicil":{"x":0.3000443407523052,"y":0.8669892212773691},"unvan":{"x":0.3019636704224231,"y":0.9042687175968724},"eposta":{"x":0.6570396593942471,"y":0.8468193276085157},"cep":{"x":0.20983584625676072,"y":0.882290501259011},"sgk_sicil":{"x":0.2040778572464068,"y":0.8452892059779652},"uye_kayit_no":{"x":0.36146289019608013,"y":0.7091077122595342},"kabul_tarihi":{"x":0.3461082528351363,"y":0.7918737947233792},"cinsiyet_erkek":{"x":0.4977352967744557,"y":0.1732843298067924},"cinsiyet_kadin":{"x":0.8643272637669875,"y":0.17537086943329214},"ogrenim_ilkogretim":{"x":0.37489819788690587,"y":0.21292858271028706},"ogrenim_lise":{"x":0.5860244615998822,"y":0.2150151223367868},"ogrenim_yuksekokul":{"x":0.9161491648601727,"y":0.2150151223367868}};
          let pageFrame = { x:0.29034749329320253, y:0.08080627461685956, w:0.40250965250965254, h:0.9048178200970485 };
          try{ const raw = localStorage.getItem('ilkesen_pdf_coords_v1'); if (raw){ const m = JSON.parse(raw)||{}; Object.assign(coords, m); } }catch{}
          try{ const raw2 = localStorage.getItem('ilkesen_pdf_pageframe_v1'); if (raw2){ const pf = JSON.parse(raw2)||null; if (pf) pageFrame = pf; } }catch{}
          const bar = document.createElement('div');
          bar.style.position='absolute'; bar.style.left='8px'; bar.style.top='8px'; bar.style.padding='6px 8px'; bar.style.background='rgba(17,24,39,.8)'; bar.style.color='#fff'; bar.style.borderRadius='6px'; bar.style.fontSize='12px'; bar.style.zIndex='10';
          let stage = pageFrame ? 'fields' : 'corners';
          let cornerStep = 0; // 0: TL, 1: BR
          let tl = null, br = null;
          function updateMsg(){
            if (stage==='corners'){
              bar.textContent = cornerStep===0 ? 'Sayfanın sol-üst köşesine tıklayın' : 'Sayfanın sağ-alt köşesine tıklayın';
            } else {
              bar.textContent = 'Tıklayın: ' + (seq[i] ? seq[i][1] : 'bitti') + '  |  Kaydır: alanı kaydırmak için Kaydırma Modu';
            }
          }
          updateMsg();
          overlay.innerHTML=''; overlay.appendChild(bar);
          overlay.style.display='block';
          function onClick(ev){
            const r = overlay.getBoundingClientRect();
            const ox = (ev.clientX - r.left) / r.width; const oy = (ev.clientY - r.top) / r.height;
            if (stage==='corners'){
              if (cornerStep===0){ tl = { x: ox, y: oy }; cornerStep = 1; }
              else { br = { x: ox, y: oy }; stage='fields';
                const x = Math.min(tl.x, br.x), y = Math.min(tl.y, br.y);
                const w = Math.max(0.01, Math.abs(br.x - tl.x));
                const h = Math.max(0.01, Math.abs(br.y - tl.y));
                pageFrame = { x, y, w, h };
                try{ localStorage.setItem('ilkesen_pdf_pageframe_v1', JSON.stringify(pageFrame)); }catch{}
              }
              updateMsg();
              return;
            }
            if (!seq[i]){ finish(); return; }
            let px = ox, py = oy;
            if (pageFrame){
              px = (ox - pageFrame.x) / pageFrame.w;
              py = (oy - pageFrame.y) / pageFrame.h;
              px = Math.max(0, Math.min(1, px));
              py = Math.max(0, Math.min(1, py));
            }
            const key = seq[i][0]; coords[key] = { x: px, y: py };
            i++; updateMsg();
            if (i>=seq.length) finish();
          }
          function finish(){
            overlay.style.display='none'; overlay.removeEventListener('click', onClick);
            try{ localStorage.setItem('ilkesen_pdf_coords_v1', JSON.stringify(coords)); }catch{}
            try{ if (member) generateMembershipFormPDF(member); }catch{}
          }
          overlay.addEventListener('click', onClick);
          // Allow scrolling the calibration area while overlay is visible
          overlay.addEventListener('wheel', (e)=>{ /* let it bubble to scroller */ }, { passive:true });
          calCtx = {
            pause(){ overlay.style.display='none'; },
            resume(){ overlay.style.display='block'; }
          };
        }catch{}
      });
      btnCorners.addEventListener('click', ()=>{
        try{
          // Force re-capture of page frame only
          localStorage.removeItem('ilkesen_pdf_pageframe_v1');
          btnCal.click();
        }catch{}
      });
      btnReset.addEventListener('click', ()=>{
        try{ localStorage.removeItem('ilkesen_pdf_pageframe_v1'); localStorage.removeItem('ilkesen_pdf_coords_v1'); alert('Kalibrasyon sıfırlandı. Lütfen tekrar Kalibrasyon yapın.'); }catch{}
      });
      btnScroll.addEventListener('click', ()=>{ try{ if (!calCtx){ btnCal.click(); return; } const hidden = overlay.style.display==='none'; if (hidden){ calCtx.resume(); btnScroll.textContent='Kaydırma Modu'; } else { calCtx.pause(); btnScroll.textContent='Devam Et'; } }catch{} });
      btnMarks.addEventListener('click', ()=>{
        try{
          const cur = localStorage.getItem('ilkesen_pdf_debug_marks_v1');
          const next = cur==='all' ? '' : 'all';
          if (next) localStorage.setItem('ilkesen_pdf_debug_marks_v1', next); else localStorage.removeItem('ilkesen_pdf_debug_marks_v1');
          btnMarks.textContent = next==='all' ? 'İşaret Testi: Açık' : 'İşaret Testi';
          if (member) generateMembershipFormPDF(member);
        }catch{}
      });

      btnEdu.addEventListener('click', ()=>{
        try{
          const panel = document.createElement('div');
          panel.style.position='absolute'; panel.style.left='8px'; panel.style.bottom='8px'; panel.style.background='rgba(17,24,39,.9)'; panel.style.color='#fff'; panel.style.padding='8px'; panel.style.borderRadius='8px'; panel.style.zIndex='11';
          panel.style.display='flex'; panel.style.gap='6px';
          function mkBtn(t, val){ const b=document.createElement('button'); b.type='button'; b.className='btn'; b.textContent=t; b.onclick=()=>{ try{ if(val) localStorage.setItem('ilkesen_pdf_force_edu_v1', val); else localStorage.removeItem('ilkesen_pdf_force_edu_v1'); if (member) generateMembershipFormPDF(member); panel.remove(); }catch{} }; return b; }
          panel.appendChild(mkBtn('İlköğretim','ilkogretim'));
          panel.appendChild(mkBtn('Lise','lise'));
          panel.appendChild(mkBtn('Yüksek Okul','yuksekokul'));
          panel.appendChild(mkBtn('Kaldır',null));
          wrap.appendChild(panel);
        }catch{}
      });

      

      btnGender.addEventListener('click', ()=>{
        try{
          const panel = document.createElement('div');
          panel.style.position='absolute'; panel.style.left='8px'; panel.style.bottom='48px'; panel.style.background='rgba(17,24,39,.9)'; panel.style.color='#fff'; panel.style.padding='8px'; panel.style.borderRadius='8px'; panel.style.zIndex='11';
          panel.style.display='flex'; panel.style.gap='6px';
          function mkBtn(t, val){ const b=document.createElement('button'); b.type='button'; b.className='btn'; b.textContent=t; b.onclick=()=>{ try{ if(val) localStorage.setItem('ilkesen_pdf_force_gender_v1', val); else localStorage.removeItem('ilkesen_pdf_force_gender_v1'); if (member) generateMembershipFormPDF(member); panel.remove(); }catch{} }; return b; }
          panel.appendChild(mkBtn('Erkek','erkek'));
          panel.appendChild(mkBtn('Kadın','kadin'));
          panel.appendChild(mkBtn('Kaldır','none'));
          wrap.appendChild(panel);
        }catch{}
      });

      // Fine-tune UI for per-field offsets and font size
      btnTune.addEventListener('click', ()=>{
        try{
          const panel = document.createElement('div');
          panel.style.position='absolute'; panel.style.right='8px'; panel.style.top='8px'; panel.style.background='rgba(17,24,39,.9)'; panel.style.color='#fff'; panel.style.padding='8px'; panel.style.borderRadius='8px'; panel.style.zIndex='11'; panel.style.minWidth='220px';
          const title = document.createElement('div'); title.textContent='İnce Ayar'; title.style.fontWeight='bold'; title.style.marginBottom='6px'; panel.appendChild(title);
          const select = document.createElement('select'); select.style.width='100%'; select.style.marginBottom='6px';
          const fields = [
            'kurum_adi','gorev_birim_adi','il_adi','il_kodu','ilce_adi','adi','soyadi','tc','baba','ana','dogum_tarihi','dogum_yeri','kurum_sicil','unvan','eposta','cep','sgk_sicil','uye_kayit_no','kabul_tarihi','cinsiyet_erkek','cinsiyet_kadin','ogrenim_ilkogretim','ogrenim_lise','ogrenim_yuksekokul'
          ];
          fields.splice(2, 0, 'gorev_birim_adresi');
          fields.forEach(k=>{ const o=document.createElement('option'); o.value=k; o.textContent=k; select.appendChild(o); });
          panel.appendChild(select);
          const row1 = document.createElement('div'); row1.style.display='flex'; row1.style.gap='6px'; row1.style.marginBottom='6px';
          const bL = document.createElement('button'); bL.type='button'; bL.className='btn'; bL.textContent='←';
          const bU = document.createElement('button'); bU.type='button'; bU.className='btn'; bU.textContent='↑';
          const bD = document.createElement('button'); bD.type='button'; bD.className='btn'; bD.textContent='↓';
          const bR = document.createElement('button'); bR.type='button'; bR.className='btn'; bR.textContent='→';
          [bL,bU,bD,bR].forEach(b=>{ b.style.minWidth='34px'; });
          row1.appendChild(bL); row1.appendChild(bU); row1.appendChild(bD); row1.appendChild(bR);
          panel.appendChild(row1);
          const row2 = document.createElement('div'); row2.style.display='flex'; row2.style.gap='6px'; row2.style.marginBottom='6px';
          const bSm = document.createElement('button'); bSm.type='button'; bSm.className='btn'; bSm.textContent='A-';
          const bLg = document.createElement('button'); bLg.type='button'; bLg.className='btn'; bLg.textContent='A+';
          const bClr = document.createElement('button'); bClr.type='button'; bClr.className='btn'; bClr.textContent='Sıfırla Alan';
          row2.appendChild(bSm); row2.appendChild(bLg); row2.appendChild(bClr);
          panel.appendChild(row2);
          const row3 = document.createElement('div'); row3.style.display='flex'; row3.style.gap='6px';
          const bApply = document.createElement('button'); bApply.type='button'; bApply.className='btn btn-success'; bApply.textContent='Uygula';
          const bClose = document.createElement('button'); bClose.type='button'; bClose.className='btn'; bClose.textContent='Kapat';
          row3.appendChild(bApply); row3.appendChild(bClose); panel.appendChild(row3);
          wrap.appendChild(panel);
          function loadStyles(){ try{ return JSON.parse(localStorage.getItem('ilkesen_pdf_styles_v1')||'{}'); }catch{ return {}; } }
          function saveStyles(m){ try{ localStorage.setItem('ilkesen_pdf_styles_v1', JSON.stringify(m||{})); }catch{} }
          function mutate(fn){ const key = select.value; const m = loadStyles(); m[key] = m[key] || { dx:0, dy:0, size:10 }; fn(m[key]); saveStyles(m); }
          const step = 0.003;
          bL.onclick = ()=> mutate(s=> s.dx = (s.dx||0) - step);
          bR.onclick = ()=> mutate(s=> s.dx = (s.dx||0) + step);
          bU.onclick = ()=> mutate(s=> s.dy = (s.dy||0) - step);
          bD.onclick = ()=> mutate(s=> s.dy = (s.dy||0) + step);
          bSm.onclick = ()=> mutate(s=> s.size = Math.max(6, (s.size||10) - 0.5));
          bLg.onclick = ()=> mutate(s=> s.size = Math.min(16, (s.size||10) + 0.5));
          bClr.onclick = ()=>{ const m = loadStyles(); delete m[select.value]; saveStyles(m); };
          bApply.onclick = ()=>{ try{ if (member) generateMembershipFormPDF(member); }catch{} };
          bClose.onclick = ()=>{ try{ panel.remove(); }catch{} };
        }catch{}
      });
      tools.appendChild(btnPrint); tools.appendChild(btnDownload);
      tools.appendChild(btnOpen); tools.appendChild(btnCal); tools.appendChild(btnScroll); tools.appendChild(btnCorners); tools.appendChild(btnFull); tools.appendChild(btnZoomOut); tools.appendChild(btnZoomIn); tools.appendChild(btnMarks); tools.appendChild(btnEdu); tools.appendChild(btnGender); tools.appendChild(btnTune); tools.appendChild(btnReset);
      form.appendChild(tools); form.appendChild(scroller);
      btnFull.addEventListener('click', ()=>{
        const fs = scroller.getAttribute('data-fullscreen') === '1';
        if (!fs){ scroller.setAttribute('data-fullscreen','1'); scroller.style.height = 'calc(100vh - 140px)'; iframe.style.height = 'calc(100vh - 140px)'; btnFull.textContent='Pencere'; }
        else { scroller.setAttribute('data-fullscreen','0'); scroller.style.height = '80vh'; iframe.style.height = '80vh'; btnFull.textContent='Tam Ekran'; }
      });
      function applyScale(){ wrap.style.transform = `scale(${__scale})`; }
      btnZoomOut.addEventListener('click', ()=>{ __scale = Math.max(0.5, Math.round((__scale-0.1)*10)/10); applyScale(); });
      btnZoomIn.addEventListener('click', ()=>{ __scale = Math.min(2.0, Math.round((__scale+0.1)*10)/10); applyScale(); });
      // If iframe gets blocked by browser policy, auto-open new tab fallback
      setTimeout(()=>{
        try{
          const cw = iframe.contentWindow;
          const ok = !!cw;
          if (!ok){ const w = window.open(url, '_blank'); if (w) w.focus(); }
        }catch{ try{ const w = window.open(url, '_blank'); if (w) w.focus(); }catch{} }
      }, 900);
      try{ (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal())); }catch{}
    }catch{}
  }

  async function uploadToBucket(file, filePath) {
    try {
      // Normalize path: remove leading slashes and accidental bucket prefix
      const cleanPath = String(filePath).replace(/^\/+/, '').replace(/^member-photos\//, '');
      
      // Upload to the root (or given subfolder) of the bucket
      const { error } = await sb()
        .storage
        .from('member-photos')
        .upload(cleanPath, file, {
          cacheControl: '3600',
          upsert: true
        });
      
      if (error) {
        console.error('Upload error:', error);
        throw error;
      }
      
      // Get the public URL for this object
      const { data: { publicUrl } } = sb()
        .storage
        .from('member-photos')
        .getPublicUrl(cleanPath);
        
      return publicUrl;
    } catch (error) {
      console.error('Error in uploadToBucket:', error);
      throw new Error('Dosya yüklenirken bir hata oluştu: ' + (error.message || 'Bilinmeyen hata'));
    }
  }

  // Resolve a usable image URL for member photos. If bucket is private, create a signed URL.
  async function resolveMemberPhotoUrl(src){
    if (!src) return '';
    try{
      // Remove crop meta if present
      src = stripPhotoKey(src);
      // Try to extract the object key from a Supabase storage URL
      if (/^https?:\/\//i.test(src)){
        // Match: /storage/v1/object/(public|authenticated)/<bucket>/<key>
        const mm = src.match(/\/storage\/v1\/object\/(public|authenticated)\/([^\/]+)\/([^?]+)/);
        if (mm && mm[1] && mm[2] && mm[3]) {
          const kind = mm[1]; // 'public' | 'authenticated'
          const bucket = decodeURIComponent(mm[2]);
          const key = decodeURIComponent(mm[3]);
          if (kind === 'public') {
            // Public URL'yi olduğu gibi kullan (CORS düzgün gelir)
            return src;
          }
          // authenticated vb. ise bulunduğu bucket üzerinden imzala
          const { data, error } = await sb().storage.from(bucket).createSignedUrl(key, 60*60);
          if (!error && data?.signedUrl) return data.signedUrl;
          return src;
        }
        // Non-supabase URL: return as-is
        return src;
      }
      // If it's not a URL, treat it as an object key inside the default member-photos bucket
      const key = String(src).replace(/^\/+/, '');
      const { data, error } = await sb().storage.from('member-photos').createSignedUrl(key, 60*60);
      if (!error && data?.signedUrl) return data.signedUrl;
      return src;
    }catch(e){
      console.warn('resolveMemberPhotoUrl failed:', e);
      return src;
    }
  }

  // === Photo crop meta helpers (compact: key|z=1.00,ox=0,oy=0) ===
  function stripPhotoKey(v){
    if (!v) return '';
    const s = String(v);
    const i = s.indexOf('|');
    return i >= 0 ? s.slice(0, i) : s;
  }
  function parsePhotoMeta(v){
    const meta = { z: 1.0, ox: 0, oy: 0, oxp: null, oyp: null };
    if (!v) return meta;
    const s = String(v);
    const i = s.indexOf('|');
    if (i < 0) return meta;
    const tail = s.slice(i+1);
    tail.split(',').forEach(p=>{
      const [k, val] = p.split('=');
      const n = Number(val);
      if (k === 'z' && isFinite(n)) meta.z = n;
      if (k === 'ox' && isFinite(n)) meta.ox = Math.round(n);
      if (k === 'oy' && isFinite(n)) meta.oy = Math.round(n);
      if (k === 'oxp' && isFinite(n)) meta.oxp = n;
      if (k === 'oyp' && isFinite(n)) meta.oyp = n;
    });
    return meta;
  }
function parsePhotoMetaMobile(v){
  const meta = { zm: 1.0, oxm: 0, oym: 0, oxpm: null, oypm: null };
  if (!v) return meta;
  const s = String(v);
  const i = s.indexOf('|');
  if (i < 0) return meta;
  const tail = s.slice(i+1);
  tail.split(',').forEach(p=>{
    const [k, val] = p.split('=');
    const n = Number(val);
    if (k === 'zm' && isFinite(n)) meta.zm = n;
    if (k === 'oxm' && isFinite(n)) meta.oxm = Math.round(n);
    if (k === 'oym' && isFinite(n)) meta.oym = Math.round(n);
    if (k === 'oxpm' && isFinite(n)) meta.oxpm = n;
    if (k === 'oypm' && isFinite(n)) meta.oypm = n;
  });
  return meta;
}  function buildPhotoValue(key, meta){
  const k = stripPhotoKey(key);
  const z = (meta?.z ?? 1).toFixed(3);
  const ox = Math.round(meta?.ox ?? 0);
  const oy = Math.round(meta?.oy ?? 0);
  const oxp = (meta?.oxp != null) ? Number(meta.oxp) : null;
  const oyp = (meta?.oyp != null) ? Number(meta.oyp) : null;
  const parts = [`z=${z}`, `ox=${ox}`, `oy=${oy}`];
  if (oxp != null && isFinite(oxp)) parts.push(`oxp=${Math.round(oxp * 10) / 10}`);
  if (oyp != null && isFinite(oyp)) parts.push(`oyp=${Math.round(oyp * 10) / 10}`);
  // Mobile-specific values (optional)
  if (meta && (meta.zm != null || meta.oxm != null || meta.oym != null || meta.oxpm != null || meta.oypm != null)){
    const zm = (meta?.zm ?? 1).toFixed(3);
    const oxm = Math.round(meta?.oxm ?? 0);
    const oym = Math.round(meta?.oym ?? 0);
    parts.push(`zm=${zm}`, `oxm=${oxm}`, `oym=${oym}`);
    if (meta?.oxpm != null && isFinite(meta.oxpm)) parts.push(`oxpm=${Math.round(meta.oxpm * 10) / 10}`);
    if (meta?.oypm != null && isFinite(meta.oypm)) parts.push(`oypm=${Math.round(meta.oypm * 10) / 10}`);
  }
  return `${k}|${parts.join(',')}`;
}

  function openMemberModal(row){
    editing = { type:'member', id: row.id };
    $modalTitle().textContent = row.id ? 'Üyeyi Düzenle' : 'Üyelik Başvuru Formu';
    $modalForm().innerHTML= '';
    const form = $modalForm();

    // Helpers for building labeled inputs/selects
    function inputEl(labelText, name, value, type='text'){
      const label = document.createElement('label');
      label.style.display='grid';
      label.style.gap='6px';
      label.innerHTML = `<span>${labelText}</span>`;
      const input = document.createElement('input');
      input.name = name;
      input.type = type;
      input.value = value || '';
      input.className = 'btn btn-outline';
      input.style.padding='8px';
      input.style.border='1px solid #e5e7eb';
      input.style.borderRadius='6px';
      label.appendChild(input);
      return label;
    }
    function selectEl(labelText, name, value, options){
      const label = document.createElement('label');
      label.style.display='grid';
      label.style.gap='6px';
      label.innerHTML = `<span>${labelText}</span>`;
      const sel = document.createElement('select');
      sel.name=name;
      sel.className='btn btn-outline';
      sel.style.padding='8px';
      sel.style.border='1px solid #e5e7eb';
      sel.style.borderRadius='6px';
      (options||[]).forEach(o=>{ const opt=document.createElement('option'); opt.value=o.v; opt.textContent=o.t; if(String(o.v)===String(value)) opt.selected=true; sel.appendChild(opt); });
      label.appendChild(sel);
      return label;
    }

    // Identity
    form.appendChild(inputEl('Ad', 'first_name', row.first_name|| ''));
    form.appendChild(inputEl('Soyad', 'last_name', row.last_name|| ''));
    form.appendChild(inputEl('TC Kimlik No', 'national_id', row.national_id|| ''));
    // Family
    form.appendChild(inputEl('Baba Adı', 'father_name', row.father_name|| ''));
    form.appendChild(inputEl('Anne Adı', 'mother_name', row.mother_name|| ''));
    // Birth
    form.appendChild(inputEl('Doğum Yeri', 'birth_place', row.birth_place|| ''));
    form.appendChild(inputEl('Doğum Tarihi', 'birth_date', row.birth_date||'', 'date'));
    // Personal
    form.appendChild(selectEl('Cinsiyet', 'gender', row.gender||'', [ {v:'',t:'Seçiniz'}, {v:'erkek',t:'Erkek'}, {v:'kadin',t:'Kadın'} ]));
    form.appendChild(selectEl('Öğrenim Durumu', 'education', row.education||'', [ {v:'',t:'Seçiniz'},{v:'ilkokul',t:'İlkokul'},{v:'ortaokul',t:'Ortaokul'},{v:'lise',t:'Lise'},{v:'onlisans',t:'Önlisans'},{v:'lisans',t:'Lisans'},{v:'yukseklisans',t:'Yüksek Lisans'},{v:'doktora',t:'Doktora'} ]));
    // Work (Province/District dynamic selects)
    const provLabel = document.createElement('label'); provLabel.style.display='grid'; provLabel.style.gap='6px'; provLabel.innerHTML='<span>İl</span>';
    const provSel = document.createElement('select'); provSel.name='work_province'; provSel.className='btn btn-outline'; provSel.style.padding='8px'; provSel.style.border='1px solid #e5e7eb'; provSel.style.borderRadius='6px'; provLabel.appendChild(provSel); form.appendChild(provLabel);
    const distLabel = document.createElement('label'); distLabel.style.display='grid'; distLabel.style.gap='6px'; distLabel.innerHTML='<span>İlçe</span>';
    const distSel = document.createElement('select'); distSel.name='work_district'; distSel.className='btn btn-outline'; distSel.style.padding='8px'; distSel.style.border='1px solid #e5e7eb'; distSel.style.borderRadius='6px'; distLabel.appendChild(distSel); form.appendChild(distLabel);
    // Helpers for geo dropdowns and cache-busting
    const bust = (u)=>{ try{ const x=new URL(u, location.origin); x.searchParams.set('v', Date.now()); return x.toString(); }catch{ return u; } };
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
      try{
        const { data: provs, error } = await sb().from('provinces').select('id,name,plate_code').order('name');
        if (error) throw error;
        const out = (provs||[]).map(p => ({ id: String(p.id), name: upper(p.name), plate_code: p.plate_code }));
        window.__ilkesenGeoCache.provinces = out;
        window.__ilkesenGeoCache.ts = Date.now();
        await saveCache();
        return window.__ilkesenGeoCache.provinces;
      }catch{
        // Fallback to local JSON keys as province list if DB is unavailable
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
    }
    async function fetchDistrictsSmart({ provinceId, plateCode, provinceName }){
      const cache = await loadCache();
      // Prefer cache by provinceId
      const cacheKey = provinceId ? `pid:${provinceId}` : (plateCode ? `pc:${plateCode}` : (provinceName ? `pn:${provinceName}` : ''));
      if (cacheKey && cache.districtsByProv[cacheKey]) return cache.districtsByProv[cacheKey];
      // 1) Try DB by province id
      if (provinceId != null && provinceId !== ''){
        try{
          const { data: d1, error: e1 } = await sb().from('districts').select('id,name,province_id').eq('province_id', provinceId).order('name');
          if (!e1 && d1 && d1.length){
            window.__ilkesenGeoCache.districtsByProv[cacheKey] = d1;
            window.__ilkesenGeoCache.ts = Date.now(); await saveCache();
            return d1;
          }
        }catch{}
      }
      // 2) Try DB by plate code if provided
      if (plateCode != null && plateCode !== ''){
        try{
          const { data: d2, error: e2 } = await sb().from('districts').select('id,name,province_id').eq('province_id', plateCode).order('name');
          if (!e2 && d2 && d2.length){
            const key = cacheKey || `pc:${plateCode}`;
            window.__ilkesenGeoCache.districtsByProv[key] = d2; window.__ilkesenGeoCache.ts = Date.now(); await saveCache();
            return d2;
          }
        }catch{}
      }
      // 3) Fallback to local JSON by province name
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
    async function populateProvincesAndDistricts(){
      try{
        const provs = await fetchProvinces();
        provSel.innerHTML = '';
        const optNone = document.createElement('option'); optNone.value= ''; optNone.textContent='İl Seçiniz'; provSel.appendChild(optNone);
        (provs||[]).forEach(p=>{ const o=document.createElement('option'); o.value=String(p.id); o.textContent=p.name; o.dataset.plate = String(p.plate_code||p.plate|| ''); provSel.appendChild(o); });
        // Preselect by name if row has stored province name
        if (row.work_province){
          const found = (provs||[]).find(p=> String(p.name|| '').toLowerCase() === String(row.work_province|| '').toLowerCase());
          if (found) provSel.value = String(found.id);
        }
        // Load districts for selected province
        async function loadDistrictsForSelected(){
          distSel.innerHTML= '';
          const o0=document.createElement('option'); o0.value= ''; o0.textContent='İlçe Seçiniz'; distSel.appendChild(o0);
          const selOpt = provSel.selectedOptions && provSel.selectedOptions[0];
          const pid = provSel.value;
          const plate = selOpt ? (selOpt.dataset.plate || '') : '';
          const pname = selOpt ? (selOpt.textContent || '').trim() : '';
          if (!pid && !plate && !pname){ return; }
          let dists = await fetchDistrictsSmart({ provinceId: pid, plateCode: plate, provinceName: pname });
          // If nothing came and we have a province id but missing plate code, try to resolve plate_code from DB and retry
          if ((!dists || !dists.length) && pid && !plate){
            try{
              const { data: p } = await sb().from('provinces').select('plate_code').eq('id', pid).maybeSingle();
              const plate2 = p && (p.plate_code != null) ? String(p.plate_code) : '';
              if (plate2){
                dists = await fetchDistrictsSmart({ provinceId: pid, plateCode: plate2, provinceName: pname });
              }
            }catch{}
          }
          (dists||[]).forEach(d=>{ const o=document.createElement('option'); o.value=String(d.id); o.textContent=d.name; distSel.appendChild(o); });
          if (!dists || !dists.length){
            const o=document.createElement('option'); o.value= ''; o.textContent='(İlçeler yüklenemedi)'; distSel.appendChild(o);
          }
          if (row.work_district){
            const foundD = (dists||[]).find(d=> String(d.name|| '').toLowerCase() === String(row.work_district|| '').toLowerCase());
            if (foundD) distSel.value = String(foundD.id);
          }
        }
        await loadDistrictsForSelected();
        provSel.addEventListener('change', loadDistrictsForSelected);
      }catch{}
    }
    // Kick off population
    populateProvincesAndDistricts();
    form.appendChild(inputEl('Çalıştığınız Kurum Tam Adı', 'institution_name', row.institution_name|| ''));
    form.appendChild(inputEl('Görev Yapılan Birim', 'work_unit', row.work_unit|| ''));
    form.appendChild(inputEl('Görev Yapılan Birimin Adresi', 'work_unit_address', row.work_unit_address|| ''));
    // Corp
    form.appendChild(inputEl('Kurum Sicil No', 'corp_reg_no', row.corp_reg_no|| ''));
    form.appendChild(inputEl('Unvan', 'title', row.title|| ''));
    form.appendChild(selectEl('Kan Grubu', 'blood_type', row.blood_type||'', [ {v:'',t:'Seçiniz'},{v:'0 Rh+',t:'0 Rh+'},{v:'0 Rh-',t:'0 Rh-'},{v:'A Rh+',t:'A Rh+'},{v:'A Rh-',t:'A Rh-'},{v:'B Rh+',t:'B Rh+'},{v:'B Rh-',t:'B Rh-'},{v:'AB Rh+',t:'AB Rh+'},{v:'AB Rh-',t:'AB Rh-'} ]));
    // Other
    form.appendChild(inputEl('Emekli Sandığı Sicil No', 'retirement_no', row.retirement_no|| ''));
    form.appendChild(inputEl('SSK No', 'ssk_no', row.ssk_no|| ''));
    try{
      ['first_name','last_name','father_name','mother_name','birth_place','institution_name','work_unit','title'].forEach(n=>{
        const el = form.querySelector(`input[name="${n}"]`);
        if (!el) return;
        el.addEventListener('input', ()=>{
          let v = el.value || '';
          v = v.replace(/[^A-Za-zÇĞİÖŞÜçğıöşü\s]/g,'');
          try{ v = v.toLocaleUpperCase('tr-TR'); }catch{ v = v.toUpperCase(); }
          v = v.replace(/\s+/g,' ');
          el.value = v;
        });
        el.addEventListener('blur', ()=>{ el.value = (el.value||'').trim().replace(/\s+/g,' '); });
      });
      const nid = form.querySelector('input[name="national_id"]');
      if (nid){
        try{ nid.maxLength = 11; nid.setAttribute('inputmode','numeric'); }catch{}
        nid.addEventListener('input', ()=>{ nid.value = (nid.value||'').replace(/\D/g,'').slice(0,11); });
      }
      ;['corp_reg_no','retirement_no','ssk_no'].forEach(n=>{
        const el = form.querySelector(`input[name="${n}"]`);
        if (!el) return;
        try{ el.setAttribute('inputmode','numeric'); }catch{}
        el.addEventListener('input', ()=>{ el.value = (el.value||'').replace(/\D/g,''); });
      });
    }catch{}
    // Contact
    form.appendChild(inputEl('E-posta', 'email', row.email|| ''));
    const phoneInput = inputEl('Telefon (5XX...)', 'phone', row.phone|| '');
    form.appendChild(phoneInput);
    // Phone input mask: (XXX) XXX XX XX (visual only)
    try{
      const phoneEl = phoneInput.querySelector('input');
      if (phoneEl){
        phoneEl.placeholder = '5XX XXX XX XX';
        try{ phoneEl.setAttribute('inputmode','numeric'); phoneEl.autocomplete='off'; }catch{}
        phoneEl.maxLength = 10;
        phoneEl.addEventListener('input', ()=>{
          let d = String(phoneEl.value||'').replace(/\D/g,'').slice(0,10);
          phoneEl.value = d;
        });
      }
    }catch{}
    // Dates & status (only in edit mode)
    if (row.id){
      form.appendChild(inputEl('Üye Kayıt Tarihi', 'join_date', row.join_date||new Date().toISOString().slice(0,10), 'date'));
      form.appendChild(inputEl('Üyelikten Ayrılış Tarihi', 'leave_date', row.leave_date||'', 'date'));
      form.appendChild(selectEl('Durum', 'status', row.status||'active', [ {v:'active',t:'Aktif'}, {v:'passive',t:'Pasif'} ]));
    }
        // Files
    const photoLabel = document.createElement("label"); photoLabel.style.display="grid"; photoLabel.style.gap="6px"; photoLabel.innerHTML = "<span>Fotoğraf</span>";
    const photoPrev = document.createElement("img");
    photoPrev.style.maxWidth = "160px";
    photoPrev.style.maxHeight = "160px";
    photoPrev.style.objectFit = "cover";
    photoPrev.style.borderRadius = "10px";
    photoPrev.style.border = "1px solid #e5e7eb";
    photoPrev.style.display = "block";
    const genderSel = form.querySelector("select[name=\"gender\"]");
    function __computeMemberFallbackByGender(){
      const g = (genderSel && genderSel.value) || row.gender || "";
      return /kadin|kadın|female/i.test(String(g)) ? "kadin.jpeg" : "erkek.jpeg";
    }
    if (row.photo_url){
      // Resolve to signed URL if needed; avoid cache-busting on signed URLs
      resolveMemberPhotoUrl(row.photo_url).then(u=>{ photoPrev.src = u; }).catch(()=>{ try{ photoPrev.src = stripPhotoKey(row.photo_url); }catch{} });
    } else {
      try{ photoPrev.src = __computeMemberFallbackByGender(); }catch{}
    }
    photoLabel.appendChild(photoPrev);
    photoPrev.addEventListener("error", ()=>{ try{ photoPrev.src = __computeMemberFallbackByGender(); photoPrev.style.display="block"; }catch{} });
    const photoInput = document.createElement("input"); photoInput.type="file"; photoInput.accept="image/*"; photoLabel.appendChild(photoInput); form.appendChild(photoLabel);
    // Live preview on file select (does not upload yet)
    photoInput.addEventListener("change", ()=>{
      const f = photoInput.files && photoInput.files[0];
      if (!f){ try{ photoPrev.src = __computeMemberFallbackByGender(); photoPrev.style.display="block"; }catch{} return; }
      const url = URL.createObjectURL(f);
      photoPrev.src = url; photoPrev.style.display="block";
    });
    // Crop controls (only in edit mode)
    if (row.id){
      const cropWrap = document.createElement("div"); cropWrap.className="card"; cropWrap.style.padding="10px"; cropWrap.style.display="grid"; cropWrap.style.gap="8px"; cropWrap.style.marginTop="8px";
      const cropTitle = document.createElement("div"); cropTitle.textContent="Kimlik Fotoğraf Kadrajı"; cropTitle.style.fontWeight="600"; cropWrap.appendChild(cropTitle);
      const initMeta = parsePhotoMeta(row.photo_url||"");
      const zoomLbl = document.createElement("label"); zoomLbl.style.display="grid"; zoomLbl.style.gap="4px"; zoomLbl.innerHTML="<span>Yakınlık</span>";
      const zoomRange = document.createElement("input"); zoomRange.type="range"; zoomRange.min="0.30"; zoomRange.max="2.00"; zoomRange.step="0.01"; zoomRange.value=String(initMeta.z||1); zoomRange.setAttribute("data-photo-zoom","1");
      zoomLbl.appendChild(zoomRange); cropWrap.appendChild(zoomLbl);
      const oyLbl = document.createElement("label"); oyLbl.style.display="grid"; oyLbl.style.gap="4px"; oyLbl.innerHTML="<span>Dikey Ofset</span>";
      const oyRange = document.createElement("input"); oyRange.type="range"; oyRange.min="-160"; oyRange.max="160"; oyRange.step="1"; oyRange.value=String(initMeta.oy||0); oyRange.setAttribute("data-photo-oy","1");
      oyLbl.appendChild(oyRange); cropWrap.appendChild(oyLbl);
      const oxLbl = document.createElement("label"); oxLbl.style.display="grid"; oxLbl.style.gap="4px"; oxLbl.innerHTML="<span>Yatay Ofset</span>";
      const oxRange = document.createElement("input"); oxRange.type="range"; oxRange.min="-160"; oxRange.max="160"; oxRange.step="1"; oxRange.value=String(initMeta.ox||0); oxRange.setAttribute("data-photo-ox","1");
      oxLbl.appendChild(oxRange); cropWrap.appendChild(oxLbl);
      const circPrev = document.createElement("canvas"); circPrev.width=160; circPrev.height=160; circPrev.style.borderRadius="50%"; circPrev.style.border="1px solid #e5e7eb"; cropWrap.appendChild(circPrev);
      form.appendChild(cropWrap);

      let circImg = null;
      async function ensureCircImg(){
        if (circImg) return circImg;
        try{
          if (photoInput.files && photoInput.files[0]){
            circImg = await new Promise((res)=>{ const i=new Image(); i.onload=()=>res(i); try{ i.src = URL.createObjectURL(photoInput.files[0]); }catch{ res(null);} });
            return circImg;
          }
          if (row.photo_url){
            try{
              const u = await resolveMemberPhotoUrl(row.photo_url);
              circImg = await new Promise((res)=>{ const i=new Image(); i.onload=()=>res(i); i.onerror=()=>res(null); i.src=u; });
              if (circImg) return circImg;
            }catch{}
          }
          const fb = __computeMemberFallbackByGender();
          circImg = await new Promise((res)=>{ const i=new Image(); i.onload=()=>res(i); i.onerror=()=>res(null); i.src=fb; });
          return circImg;
        }catch{}
        return null;
      }
      async function drawCircPrev(){
        const ctx = circPrev.getContext("2d");
        ctx.clearRect(0,0,160,160);
        ctx.save(); ctx.beginPath(); ctx.arc(80,80,78,0,Math.PI*2); ctx.closePath(); ctx.clip();
        const img = await ensureCircImg();
        if (!img){ ctx.fillStyle="#e5e7eb"; ctx.fillRect(0,0,160,160); ctx.restore(); return; }
        const R = 78; const baseScale = Math.max((R*2)/img.width, (R*2)/img.height);
        let scale = baseScale * Number(zoomRange.value||1);
        if (img.width*scale < R*2 || img.height*scale < R*2) scale = baseScale;
        const w = img.width*scale, h = img.height*scale;
        const x = 80 - w/2 + Number(oxRange.value||0);
        const y = 80 - h/2 + Number(oyRange.value||0);
        try{ ctx.drawImage(img, x, y, w, h); }catch{}
        ctx.restore();
        ctx.beginPath(); ctx.arc(80,80,78,0,Math.PI*2); ctx.closePath(); ctx.strokeStyle="#0ea5b1"; ctx.lineWidth=3; ctx.stroke();
        ctx.beginPath(); ctx.arc(80,80,72,0,Math.PI*2); ctx.closePath(); ctx.strokeStyle="#fb923c"; ctx.lineWidth=2; ctx.stroke();
      }
      drawCircPrev();
      [zoomRange, oyRange, oxRange].forEach(el=> el.addEventListener("input", drawCircPrev));
      if (genderSel){
        genderSel.addEventListener("change", ()=>{
          if (!(photoInput.files && photoInput.files[0]) && !row.photo_url){
            try{ photoPrev.src = __computeMemberFallbackByGender(); photoPrev.style.display="block"; }catch{}
            circImg = null; ensureCircImg().then(()=>drawCircPrev());
          }
        });
      }
      photoInput.addEventListener("change", ()=>{ circImg=null; ensureCircImg().then(()=>drawCircPrev()); try{ if (photoInput.files && photoInput.files[0]){ photoPrev.src = URL.createObjectURL(photoInput.files[0]); photoPrev.style.display="block"; } }catch{} });
    }
    const docLabel = document.createElement('label'); docLabel.style.display='grid'; docLabel.style.gap='6px'; docLabel.innerHTML = '<span>Belgeler</span>';
    const docInput = document.createElement('input'); docInput.type='file'; docInput.multiple=true; docLabel.appendChild(docInput); form.appendChild(docLabel);

    const docsPrev = document.createElement('div');
    docsPrev.style.display='grid';
    docsPrev.style.gridTemplateColumns='repeat(auto-fill, minmax(120px, 1fr))';
    docsPrev.style.gap='8px';
    docsPrev.style.margin='6px 0 10px';
    form.appendChild(docsPrev);
    function isImg(url){ return /(\.png|\.jpe?g|\.gif|\.webp|\.bmp|\.svg)$/i.test(String(url|| '')); }
    function fileName(url){ try{ const u=new URL(url, location.origin); return u.pathname.split('/').pop()||'dosya'; }catch{ const parts=String(url|| '').split('/'); return parts[parts.length-1]||'dosya'; } }
    function renderExistingDocs(){
      docsPrev.innerHTML= '';
      let arr=[]; try{ arr = JSON.parse(row.documents_urls||'[]'); if(!Array.isArray(arr)) arr=[]; }catch{}
      arr.forEach(url=>{
        const card = document.createElement('div'); card.className='card'; card.style.padding='6px'; card.style.display='flex'; card.style.flexDirection='column'; card.style.gap='6px';
        if (isImg(url)){
          const img = document.createElement('img'); img.alt= ''; img.style.width='100%'; img.style.height='86px'; img.style.objectFit='cover'; img.style.borderRadius='8px';
          img.src = bust(url);
          card.appendChild(img);
        } else {
          const a=document.createElement('a'); a.href=url; a.target='_blank'; a.rel='noopener noreferrer'; a.textContent=fileName(url); a.style.wordBreak='break-all'; card.appendChild(a);
        }
        docsPrev.appendChild(card);
      });
    }
    renderExistingDocs();
    const docsNewPrevTitle = document.createElement('div'); docsNewPrevTitle.className='muted'; docsNewPrevTitle.textContent='Seçilen Belgeler (önizleme)'; docsNewPrevTitle.style.margin='6px 0 2px';
    const docsNewPrev = document.createElement('div');
    docsNewPrev.style.display='grid'; docsNewPrev.style.gridTemplateColumns='repeat(auto-fill, minmax(120px, 1fr))'; docsNewPrev.style.gap='8px';
    form.appendChild(docsNewPrevTitle); form.appendChild(docsNewPrev);
    function renderNewDocs(files){
      docsNewPrev.innerHTML= '';
      const fs = Array.from(files||[]);
      fs.forEach(f=>{
        const card = document.createElement('div'); card.className='card'; card.style.padding='6px'; card.style.display='flex'; card.style.flexDirection='column'; card.style.gap='6px';
        if (/^image\//i.test(f.type)){
          const img = document.createElement('img'); img.alt= ''; img.style.width='100%'; img.style.height='86px'; img.style.objectFit='cover'; img.style.borderRadius='8px';
          try{ img.src = URL.createObjectURL(f); }catch{}
          card.appendChild(img);
        } else {
          const span=document.createElement('span'); span.textContent=f.name; span.title=f.name; span.style.wordBreak='break-all'; card.appendChild(span);
        }
        docsNewPrev.appendChild(card);
      });
      docsNewPrevTitle.style.display = fs.length ? 'block' : 'none';
    }
    docsNewPrevTitle.style.display='none';
    docInput.addEventListener('change', ()=>{ renderNewDocs(docInput.files||[]); });

    const actions = document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
    const saveBtn = document.createElement('button');
    saveBtn.type = 'button'; // avoid form submit
    saveBtn.className = 'btn btn-success';
    saveBtn.textContent = 'Kaydet';
    actions.appendChild(saveBtn);
    form.appendChild(actions);

    saveBtn.addEventListener('click', async (e)=>{
      e.preventDefault();
      const fd=new FormData(form);
      const payload={
        first_name: String(fd.get('first_name')|| '').trim(),
        last_name: String(fd.get('last_name')|| '').trim(),
        national_id: String(fd.get('national_id')|| '').trim(),
        father_name: String(fd.get('father_name')|| '').trim(),
        mother_name: String(fd.get('mother_name')|| '').trim(),
        birth_place: String(fd.get('birth_place')|| '').trim(),
        birth_date: String(fd.get('birth_date')|| '').trim() || null,
        gender: String(fd.get('gender')|| '').trim()||null,
        education: String(fd.get('education')|| '').trim()||null,
        // Convert selected province/district IDs to names for storage
        work_province: null,
        work_district: null,
        institution_name: String(fd.get('institution_name')|| '').trim(),
        work_unit: String(fd.get('work_unit')|| '').trim(),
        work_unit_address: String(fd.get('work_unit_address')|| '').trim(),
        corp_reg_no: String(fd.get('corp_reg_no')|| '').trim(),
        title: String(fd.get('title')|| '').trim(),
        blood_type: String(fd.get('blood_type')|| '').trim()||null,
        retirement_no: String(fd.get('retirement_no')|| '').trim(),
        ssk_no: String(fd.get('ssk_no')|| '').trim(),
        email: String(fd.get('email')|| '').trim().toLowerCase() || null, // Normalize optional email: lowercase and send null if empty
        phone: String(fd.get('phone')|| '').trim(),
        join_date: row.id ? (String(fd.get('join_date')|| '').trim() || new Date().toISOString().slice(0,10)) : null,
        leave_date: row.id ? (String(fd.get('leave_date')|| '').trim() || null) : null,
        status: row.id ? String(fd.get('status')||'active') : 'active'
      };
      try{
        ['first_name','last_name','father_name','mother_name','birth_place','institution_name','work_unit','title'].forEach(k=>{
          if (payload[k]!=null){ try{ payload[k] = String(payload[k]).toLocaleUpperCase('tr-TR'); }catch{ payload[k] = String(payload[k]).toUpperCase(); } }
        });
        if (payload.national_id!=null) payload.national_id = String(payload.national_id).replace(/\D/g,'');
        if (payload.corp_reg_no!=null) payload.corp_reg_no = String(payload.corp_reg_no).replace(/\D/g,'');
        if (payload.retirement_no!=null) payload.retirement_no = String(payload.retirement_no).replace(/\D/g,'');
        if (payload.ssk_no!=null) payload.ssk_no = String(payload.ssk_no).replace(/\D/g,'');
      }catch{}
      // Generic cleanup: convert empty strings to null to avoid DB CHECK issues
      Object.keys(payload).forEach(k => { if (payload[k] === '') payload[k] = null; });
      if (!payload.first_name || !payload.last_name) return alert('Ad ve Soyad zorunlu');

      // Resolve province/district NAMES from selected options (no extra round-trips)
      try{
        payload.work_province = provSel && provSel.value ? (provSel.selectedOptions[0]?.textContent || null) : null;
        payload.work_district = distSel && distSel.value ? (distSel.selectedOptions[0]?.textContent || null) : null;
      }catch{}

      // Basic client-side validations
      if (payload.national_id){
        if (!/^\d{11}$/.test(payload.national_id)) return alert('TC Kimlik No 11 haneli olmalı');
      }
      // Phone normalize for DB check: expect exactly 10 digits and store as +90XXXXXXXXXX
      if (payload.phone){
        let d = String(payload.phone|| '').replace(/\D/g,'');
        if (d.startsWith('90')) d = d.slice(2);
        if (d.startsWith('0')) d = d.slice(1);
        d = d.slice(0,10);
        if (d.length === 10){
          payload.phone = '+90' + d;
        } else {
          payload.phone = null; // let DB accept null if user didn't provide full number
        }
      }
      if (payload.phone === null) { delete payload.phone; }
      // Email normalize strictly before regex: NFKC, strip whitespace/zero-width, lowercase
      if (payload.email){
        try{
          const cleaned = String(payload.email)
            .normalize('NFKC')
            .replace(/[\s\u00A0\u200B-\u200D\uFEFF]/g, '')
            .toLowerCase();
          payload.email = cleaned || null;
        }catch{ payload.email = String(payload.email|| '').trim().toLowerCase() || null; }
      }
      if (payload.email && !/^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(payload.email)) return alert('Geçerli bir e‑posta giriniz');
      if (payload.email === null) { delete payload.email; }
      // Business rule: at least one of retirement_no or ssk_no must be provided
      if (!payload.retirement_no && !payload.ssk_no){ return alert('Emekli Sandığı Sicil No veya SSK No alanlarından en az biri dolu olmalıdır.'); }

      // Handle uploads: photo and documents
      try{
        // Photo: preserve existing if none selected
        let photoUrl = row.photo_url || null;
        const fileInput = form.querySelector('input[type="file"]');
        const readPhotoMeta = () => {
          const zEl = form.querySelector('[data-photo-zoom]');
          const oxEl = form.querySelector('[data-photo-ox]');
          const oyEl = form.querySelector('[data-photo-oy]');
          const z = Number((zEl && zEl.value) || 1);
          const ox = Number((oxEl && oxEl.value) || 0);
          const oy = Number((oyEl && oyEl.value) || 0);
          return { z, ox, oy };
        };
        const photoMeta = readPhotoMeta();
        if (fileInput.files && fileInput.files[0]) {
          const file = fileInput.files[0];
          const fileExt = file.name.split('.').pop();
          const fileName = `${Date.now()}.${fileExt}`;
          const filePath = fileName;
          // Store only the object key; previews will resolve with a signed URL if needed
          const key = await uploadToBucket(file, filePath);
          photoUrl = buildPhotoValue(key, photoMeta);
        }
        // Normalize legacy full public URLs to keys and attach current meta
        if (photoUrl && /^https?:\/\//i.test(photoUrl)){
          const m = String(photoUrl).match(/\/storage\/v1\/object\/(public|authenticated)\/member-photos\/([^?]+)/);
          if (m && m[1] && m[2]) photoUrl = buildPhotoValue(decodeURIComponent(m[2]), photoMeta);
        }
        // If no new file but sliders changed, persist meta on existing key
        if (!fileInput.files?.length && row.photo_url){
          const existingKey = stripPhotoKey(row.photo_url);
          photoUrl = buildPhotoValue(existingKey, photoMeta);
        }
        payload.photo_url = photoUrl;

        // Documents: start with existing list
        let documentsUrls = [];
        try{ documentsUrls = JSON.parse(row.documents_urls||'[]'); if(!Array.isArray(documentsUrls)) documentsUrls=[]; }catch{}
        const filesDocs = docInput.files ? Array.from(docInput.files) : [];
        for (const f of filesDocs){
          const fileExt = f.name.split('.').pop();
          const fileName = `${Date.now()}.${fileExt}`;
          const filePath = `member-docs/${fileName}`;
          const url = await uploadToBucketGeneric(f, 'member-docs', filePath);
          if (url) documentsUrls.push(url);
        }
        payload.documents_urls = JSON.stringify(documentsUrls);
      }catch(upErr){ console.warn('upload failed', upErr); }

      let resp;
if (editing && editing.id) {
  // Mevcut üye: UPDATE
  resp = await sb().from('members')
    .update(payload)
    .eq('id', editing.id)
    .select('id')
    .maybeSingle();
} else {
  // Yeni üye: Admin panelinden ekleniyor -> hemen üye numarası ve katılım tarihi atanır
  try{
    const memberNo = await nextMemberNo();
    const today = new Date().toISOString().slice(0,10);
    payload.member_no = memberNo;
    if (!payload.join_date) payload.join_date = today;
    if (!payload.status) payload.status = 'active';
  }catch{}
  resp = await sb().from('members')
    .insert(payload)
    .select('id')
  .maybeSingle();
}

if (resp.error) {
  const msg = String(resp.error?.message || '').toLowerCase();
  const code = String(resp.error?.code || '');
  // Graceful fallback if new column does not exist yet
  if ((code === '42703' || /undefined column|column .* does not exist/.test(msg)) && /work_unit_address/.test(msg)){
    try{
      delete payload.work_unit_address;
      let resp2;
      if (editing && editing.id){
        resp2 = await sb().from('members').update(payload).eq('id', editing.id).select('id').maybeSingle();
      } else {
        resp2 = await sb().from('members').insert(payload).select('id').maybeSingle();
      }
      if (!resp2.error){
        closeModal();
        await loadAdminMembers();
        return;
      }
    }catch{}
  }
  if (code === '23505' || /uq_members_national_id|unique|duplicate key/.test(msg)) {
    alert('Bu TC Kimlik No başka bir üyede zaten kayıtlı.');
  } else {
    alert('Kaydetme hatası: ' + (resp.error.message || String(resp.error)));
  }
  return;
}

// Başarılıysa kapat ve listeyi yenile
closeModal();
await loadAdminMembers();
    });

    (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
  }

  async function sendDigitalIdCard(memberId){
    // Sends request to Edge Function to email the rendered card
    try{
      const { data: m } = await sb().from('members').select('*').eq('id', memberId).maybeSingle();
      if (!m) throw new Error('Üye bulunamadı');
  //    const { data: s } = await sb().from('settings').select('value').eq('key','email_edge_url').maybeSingle();
  //    const edgeUrl = s?.value || '';
  //    if (!edgeUrl) throw new Error('E-posta gönderim servisi yapılandırılmamış (email_edge_url).');
      const imgData = window.__idcardLastPng || null;
      const payload = { type:'digital_id', member: m, image_base64: imgData };
      const { data, error } = await window.ilkeSupabase.functions.invoke('idcard', {
        body: payload
      });
      if (error) {
        throw new Error('Sunucu hatası: ' + (error.message || JSON.stringify(error)));
      }
    }catch(e){ throw e; }
  }

  // Show member ID card: Canvas render using PSD-exported background, logo overlays, QR and join date
  function openIdCardPreview(member){
    try{
      $modalTitle().textContent = 'Dijital Kimlik';
      const form = $modalForm(); form.innerHTML= '';
      const wrap = document.createElement('div'); wrap.style.display='grid'; wrap.style.gap='12px';
      form.appendChild(wrap);
      const canvas = document.createElement('canvas'); canvas.width = 638; canvas.height = 1012; canvas.style.maxWidth='100%'; canvas.style.borderRadius='12px'; canvas.style.border='1px solid #e5e7eb'; wrap.appendChild(canvas);
      const actions = document.createElement('div'); actions.style.display='flex'; actions.style.flexWrap='wrap'; actions.style.gap='8px'; wrap.appendChild(actions);
      const sendBtn = document.createElement('button'); sendBtn.className='btn btn-success'; sendBtn.textContent='E‑posta ile Gönder';
      const dlBtn = document.createElement('button'); dlBtn.className='btn btn-primary'; dlBtn.textContent='JPEG İndir';
      const printBtn = document.createElement('button'); printBtn.className='btn btn-primary'; printBtn.textContent='Yazdır';
      const closeBtn = document.createElement('button'); closeBtn.className='btn btn-danger'; closeBtn.textContent='Kapat';
      actions.appendChild(sendBtn); actions.appendChild(dlBtn); actions.appendChild(printBtn); actions.appendChild(closeBtn);

      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';

      (async ()=>{
        // Decide style: PSD if idcard_bg_url is set and idcard_style != 'vector'; otherwise draw vector theme
        let bgUrl = await getSettingValue('idcard_bg_url');
        if (!bgUrl) { bgUrl = 'edge/idcard/id_card.png'; } // default to provided template
        const stylePref = (await getSettingValue('idcard_style')) || '';
        const useVector = (!bgUrl || String(stylePref).toLowerCase() === 'vector') ? true : false;

        // Common settings
        const companyName = (await getSettingValue('company_name')) || (await getSettingValue('site_name')) || '';
        const companyTag = (await getSettingValue('company_tagline')) || (await getSettingValue('site_tagline')) || '';
        let logoUrl = await getSettingValue('logo_url'); if(!logoUrl){ logoUrl = await getSettingValue('footer_logo_url'); }

        // Colors similar to sample
        const C_TEAL = '#0ea5b1';
        const C_ORANGE = '#fb923c';
        const C_TEXT = '#0f172a';
        const C_MUTED = '#6b7280';
        const C_RED = '#E03B3B';
        const C_BLUE = '#0B3A60';

        // Background + rounded corners clip (exported PNG is also rounded)
        ctx.clearRect(0,0,canvas.width,canvas.height);
        // Clip entire card to rounded rect
        ctx.save();
        roundRect(ctx, 0, 0, canvas.width, canvas.height, 24);
        ctx.clip();
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0,0,canvas.width,canvas.height);

        if (!useVector){
          // Draw the PNG template exactly as background
          try{ const bg=await loadImg(bust(bgUrl)); ctx.drawImage(bg,0,0,canvas.width,canvas.height);}catch{}
        } else {
          // Draw vector waves adapted for vertical card
          // Top wave (teal)
          ctx.save();
          const gradTop = ctx.createLinearGradient(0,0,canvas.width,0);
          gradTop.addColorStop(0, '#14b8a6'); gradTop.addColorStop(1, '#0ea5b1');
          ctx.fillStyle = gradTop;
          ctx.beginPath();
          ctx.moveTo(0, 110);
          ctx.bezierCurveTo(canvas.width*0.15, 40, canvas.width*0.7, 170, canvas.width, 100);
          ctx.lineTo(canvas.width, 0);
          ctx.lineTo(0, 0);
          ctx.closePath();
          ctx.fill();
          ctx.restore();

          // Middle wave (orange)
          ctx.save();
          const gradMid = ctx.createLinearGradient(0,0,canvas.width,0);
          gradMid.addColorStop(0, '#fdba74'); gradMid.addColorStop(1, C_ORANGE);
          ctx.fillStyle = gradMid;
          ctx.beginPath();
          ctx.moveTo(0, canvas.height*0.42);
          ctx.bezierCurveTo(canvas.width*0.25, canvas.height*0.50, canvas.width*0.7, canvas.height*0.35, canvas.width, canvas.height*0.46);
          ctx.lineTo(canvas.width, canvas.height*0.58);
          ctx.bezierCurveTo(canvas.width*0.7, canvas.height*0.48, canvas.width*0.25, canvas.height*0.64, 0, canvas.height*0.56);
          ctx.closePath();
          ctx.fill();
          ctx.restore();

          // Bottom accent ribbon
          ctx.save();
          ctx.fillStyle = C_ORANGE;
          ctx.beginPath();
          ctx.moveTo(0, canvas.height-52);
          ctx.lineTo(80, canvas.height);
          ctx.lineTo(0, canvas.height);
          ctx.closePath();
          ctx.fill();
          ctx.restore();
        }
        // Small logo plate at top-left (white plate with logo)
        try{
          if (logoUrl){
            const lg=await loadImg(bust(logoUrl));

            const lx = 26, ly = 58, lr = 56; // konum ve yarıçap
            // Daire kırp
            ctx.save();
            ctx.beginPath();
            ctx.arc(lx + lr, ly + lr, lr, 0, Math.PI * 2);
            ctx.closePath();
            ctx.clip();

            // Görseli hafif taşırarak çiz: şeffaf padding olsa bile daireyi tam doldursun
            ctx.drawImage(lg, lx - 6, ly - 6, lr * 2 + 12, lr * 2 + 12);
            ctx.restore();
          
            // Border yok (çerçeveyi tamamen kaldırdık)
           }
        }catch{}
          // Headings: first two orange, third teal 'İLKE-SEN'
         // Logoyla çakışmayı önlemek için min sol kenar: logo sağ ucu + 8px
        const baseX = canvas.width / 2 + 20;
        const lx = 26, ly = 58, lr = 56;                 // logo konumları (yukarıda kullandıklarımız)
        const circleRight = lx + lr * 2;
        const minLeft = circleRight + 8;                  // metnin sol kenarı en az bu değerin sağında olmalı

        function centeredX(text) {
          // Ölç ve gerekirse merkezi sağa kaydır
          const w = ctx.measureText(text).width;
          return Math.max(baseX, minLeft + w / 2);
        }

        ctx.textAlign = 'center';

        // İLKE-SEN – biraz daha yukarı
        ctx.fillStyle = '#ffffff';
        ctx.font = 'bold 42px Inter, Arial, sans-serif';
        ctx.fillText('İLKE-SEN', centeredX('İLKE-SEN'), ly + lr - 28);

        // Turuncu başlık 1 – logonun yatay merkezine hizalı (ly + lr)
        ctx.fillStyle = C_RED;
        ctx.font = 'bold 20px Inter, Arial, sans-serif';
        const line1 = 'İLKELİ YEREL YÖNETİM HİZMETLERİ KOLU';
        ctx.fillText(line1, centeredX(line1), ly + lr);        // 156

        // Turuncu başlık 2 – 1. satırın altına
        const line2 = 'KAMU GÖREVLİLERİ SENDİKASI';
        ctx.fillText(line2, centeredX(line2), ly + lr + 26);   // 182
                          
        // Photo circle centered (continues…)
        const photoCX = canvas.width/2;
        // Match template circular hole (bigger, slightly lower) and eliminate edge gaps
        const photoCY = useVector ? 312 : 406; // 1 satır (~26px) yukarı
        const photoR = useVector ? 84 : 180; // significantly increased to fill the frame
        ctx.save();
        if (useVector){
          ctx.beginPath(); ctx.arc(photoCX, photoCY, photoR+10, 0, Math.PI*2); ctx.closePath(); ctx.strokeStyle=C_BLUE; ctx.lineWidth=6; ctx.stroke();
          ctx.beginPath(); ctx.arc(photoCX, photoCY, photoR+2, 0, Math.PI*2); ctx.closePath(); ctx.strokeStyle=C_RED; ctx.lineWidth=4; ctx.stroke();
        }
        ctx.beginPath(); ctx.arc(photoCX, photoCY, photoR, 0, Math.PI*2); ctx.closePath(); ctx.clip();
        try{
          if (member.photo_url){
            const meta = parsePhotoMeta(member.photo_url);
            const src = await resolveMemberPhotoUrl(member.photo_url);
            const ph = await loadImg(src);
            const baseScale = Math.max((photoR*2) / ph.width, (photoR*2) / ph.height);
            let scale = baseScale * (Number(meta.z)||1);
            if (ph.width * scale < photoR*2 || ph.height * scale < photoR*2){ scale = baseScale; }
            const newWidth = ph.width * scale;
            const newHeight = ph.height * scale;
            // Scale offsets from preview (R=78) to actual card radius
            const offScale = photoR / 78; // preview used 160x160 circle with R=78
            const offX = (Number(meta.ox)||0) * offScale;
            const offY = (Number(meta.oy)||0) * offScale;
            const x = photoCX - (newWidth / 2) + offX;
            const y = photoCY - (newHeight / 2) + offY;
            // Draw the photo
            ctx.drawImage(ph, x, y, newWidth, newHeight);
          }
          else { try{ const g = String(member.gender|| '').toLowerCase(); const fb = /kadin|kadın|female/.test(g) ? 'kadin.jpeg' : 'erkek.jpeg'; const ph = await loadImg(fb); const baseScale = Math.max((photoR*2)/ph.width, (photoR*2)/ph.height); const scale = baseScale; const newWidth = ph.width*scale; const newHeight = ph.height*scale; const x = photoCX - newWidth/2; const y = photoCY - newHeight/2; ctx.drawImage(ph, x, y, newWidth, newHeight); }catch{
  try{
    const g = String(member.gender|| '').toLowerCase();
    const fb = /kadin|kadın|female/.test(g) ? 'kadin.jpeg' : 'erkek.jpeg';
    const ph = await loadImg(fb);
    const baseScale = Math.max((photoR*2)/ph.width, (photoR*2)/ph.height);
    const newWidth = ph.width*baseScale;
    const newHeight = ph.height*baseScale;
    const x = photoCX - newWidth/2;
    const y = photoCY - newHeight/2;
    ctx.drawImage(ph, x, y, newWidth, newHeight);
  }catch{
    ctx.fillStyle='#e5e7eb';
    ctx.fillRect(photoCX-photoR, photoCY-photoR, photoR*2, photoR*2);
  }
}}
        }catch{
  try{
    const g = String(member.gender|| '').toLowerCase();
    const fb = /kadin|kadın|female/.test(g) ? 'kadin.jpeg' : 'erkek.jpeg';
    const ph = await loadImg(fb);
    const baseScale = Math.max((photoR*2)/ph.width, (photoR*2)/ph.height);
    const newWidth = ph.width*baseScale;
    const newHeight = ph.height*baseScale;
    const x = photoCX - newWidth/2;
    const y = photoCY - newHeight/2;
    ctx.drawImage(ph, x, y, newWidth, newHeight);
  }catch{
    ctx.fillStyle='#e5e7eb';
    ctx.fillRect(photoCX-photoR, photoCY-photoR, photoR*2, photoR*2);
  }
}
        ctx.restore();
        // Ensure a visible circular frame on template mode as well
        if (!useVector){
          ctx.save();
          // outer white halo for contrast
          ctx.beginPath(); ctx.arc(photoCX, photoCY, photoR+3, 0, Math.PI*2); ctx.closePath();
          ctx.strokeStyle = '#ffffff'; ctx.lineWidth = 6; ctx.stroke();
          // teal ring
          ctx.beginPath(); ctx.arc(photoCX, photoCY, photoR+1, 0, Math.PI*2); ctx.closePath();
          ctx.strokeStyle = C_BLUE; ctx.lineWidth = 4; ctx.stroke();
          // inner red ring
          ctx.beginPath(); ctx.arc(photoCX, photoCY, photoR-4, 0, Math.PI*2); ctx.closePath();
          ctx.strokeStyle = C_RED; ctx.lineWidth = 3; ctx.stroke();
          ctx.restore();
        }
        // Name area: in template we only print texts (no banner), in vector we draw banner
        ctx.textAlign='center';
        if (useVector){
          const bannerW = canvas.width - 120, bannerH = 54, bannerX=(canvas.width-bannerW)/2, bannerY=photoCY+photoR+24 - 26; // isim 2 satır yukarı: +26 ekstra
          ctx.save(); roundRect(ctx, bannerX, bannerY, bannerW, bannerH, 10); ctx.clip();
          ctx.fillStyle=C_TEAL; ctx.fill(); ctx.restore();
          ctx.fillStyle = C_BLUE; ctx.font = 'bold 28px Inter, Arial, sans-serif';
          ctx.fillText(`${(member.first_name|| '').toUpperCase()} ${(member.last_name|| '').toUpperCase()}`.trim(), canvas.width/2, bannerY+34);
          if (member.title){ ctx.fillStyle = C_RED; ctx.font='bold 22px Inter, Arial, sans-serif'; ctx.fillText(String(member.title).toUpperCase(), canvas.width/2, bannerY+58); }
        } else {
          // Shifted further down by ~24px from previous
          ctx.fillStyle = C_BLUE; ctx.font = 'bold 28px Inter, Arial, sans-serif';
          ctx.fillText(`${(member.first_name|| '').toUpperCase()} ${(member.last_name|| '').toUpperCase()}`.trim(), canvas.width/2, photoCY+photoR+88 - 26);
          if (member.title){ ctx.fillStyle = C_RED; ctx.font='bold 22px Inter, Arial, sans-serif'; ctx.fillText(String(member.title).toUpperCase(), canvas.width/2, photoCY+photoR+114 - 26); }
        }
        // Details and QR: if vector, draw white card; if template, print directly in template box
        const uyelik = member.join_date ? new Date(member.join_date).toLocaleDateString('tr-TR') : '-';
        if (useVector){
          const cardX=32, cardY=(photoCY+photoR+24)+54+78 - 26, cardW=canvas.width-64, cardH=260; // bilgiler 2 satır yukarı: +26 ekstra
          ctx.save(); roundRect(ctx, cardX, cardY, cardW, cardH, 12); ctx.fillStyle='#ffffff'; ctx.fill(); ctx.strokeStyle='#e5e7eb'; ctx.lineWidth=1; ctx.stroke(); ctx.restore();
          ctx.textAlign='left'; const pad=16; const leftX=Math.floor(canvas.width*0.5); let lineY=cardY+36;
          ctx.fillStyle=C_RED; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText('Üye Numarası', leftX, lineY); ctx.fillStyle=C_BLUE; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText(`: ${member.member_no||'-'}`, leftX+160, lineY);
          lineY+=36; ctx.fillStyle=C_RED; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText('Üyelik Tarihi', leftX, lineY); ctx.fillStyle=C_BLUE; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText(`: ${uyelik}`, leftX+160, lineY);
          lineY+=36; ctx.fillStyle=C_RED; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText('TC. No', leftX, lineY); ctx.fillStyle=C_BLUE; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText(`: ${member.national_id||'-'}`, leftX+160, lineY);
          try{
            async function buildVerifyUrl(){
              try{
                const base = (await getSettingValue('verify_base_url')) || '';
                const fallbackBase = location.origin + '/verify.html';
                let page = base && /^https?:\/\//i.test(base) ? base : fallbackBase;
                const u = new URL(page, location.origin);
                if (!/verify\.html$/i.test(u.pathname)) u.pathname = u.pathname.replace(/\/$/, '') + '/verify.html';
                if (member.verify_token) { u.searchParams.set('t', member.verify_token); }
                else { u.searchParams.set('id', member.id); }
                return u.toString();
              }catch{ return location.origin + '/verify.html?id=' + encodeURIComponent(member.id); }
            }
            const url = await buildVerifyUrl();

const qrSize = 160;
const padBR = 28;
const cb = Date.now() + '-' + encodeURIComponent(member.verify_token || member.id); // cache-buster
const qr = await loadImg(`https://api.qrserver.com/v1/create-qr-code/?size=${qrSize}x${qrSize}&margin=0&data=${encodeURIComponent(url)}&cb=${cb}`);
const qx = padBR, qy = canvas.height - padBR - qrSize;
ctx.drawImage(qr, qx, qy, qrSize, qrSize);}catch{}
        } else {
          // template coordinates
          ctx.textAlign='left'; let leftX=Math.floor(canvas.width/2), lineY=808-52; // başlangıç: merkeze hizala, 2 satır yukarı
          ctx.fillStyle=C_RED; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText('Üye Numarası', leftX, lineY); ctx.fillStyle=C_BLUE; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText(`: ${member.member_no||'-'}`, leftX+160, lineY);
          lineY+=36; ctx.fillStyle=C_RED; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText('Üyelik Tarihi', leftX, lineY); ctx.fillStyle=C_BLUE; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText(`: ${uyelik}`, leftX+160, lineY);
          lineY+=36; ctx.fillStyle=C_RED; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText('TC. No', leftX, lineY); ctx.fillStyle=C_BLUE; ctx.font='bold 20px Inter, Arial, sans-serif'; ctx.fillText(`: ${member.national_id||'-'}`, leftX+160, lineY);
          try{
            async function buildVerifyUrl(){
              try{
                const base = (await getSettingValue('verify_base_url')) || '';
                const fallbackBase = location.origin + '/verify.html';
                let page = base && /^https?:\/\//i.test(base) ? base : fallbackBase;
                const u = new URL(page, location.origin);
                if (!/verify\.html$/i.test(u.pathname)) u.pathname = u.pathname.replace(/\/$/, '') + '/verify.html';
                if (member.verify_token) { u.searchParams.set('t', member.verify_token); }
                else { u.searchParams.set('id', member.id); }
                return u.toString();
              }catch{ return location.origin + '/verify.html?id=' + encodeURIComponent(member.id); }
            }
            const url = await buildVerifyUrl();

const qrSize = 160;
const padBR = 28;
const cb = Date.now() + '-' + encodeURIComponent(member.verify_token || member.id); // cache-buster
const qr = await loadImg(`https://api.qrserver.com/v1/create-qr-code/?size=${qrSize}x${qrSize}&margin=0&data=${encodeURIComponent(url)}&cb=${cb}`);
const qx = padBR, qy = canvas.height - padBR - qrSize;
ctx.drawImage(qr, qx, qy, qrSize, qrSize);}catch{}
        }
        // finalize snapshot
        try{ window.__idcardLastPng = canvas.toDataURL('image/jpeg', 0.92); }catch{}
      })();
      function roundRect(ctx,x,y,w,h,r){ r=Math.min(r,w/2,h/2); ctx.save(); ctx.beginPath(); ctx.moveTo(x+r,y); ctx.lineTo(x+w-r,y); ctx.quadraticCurveTo(x+w,y,x+w,y+r); ctx.lineTo(x+w,y+h-r); ctx.quadraticCurveTo(x+w,y+h,x+w-r,y+h); ctx.lineTo(x+r,y+h); ctx.quadraticCurveTo(x,y+h,x,y+h-r); ctx.lineTo(x,y+r); ctx.quadraticCurveTo(x,y,x+r,y); ctx.closePath(); }
      function loadImg(src){
        return new Promise((res, rej) => {
          (async () => {
            try{
              const raw = String(src||'');
              let final = raw;
              // Prefer blob via safeImageUrl when possible (prevents tainted canvas)
              try{ if (window.safeImageUrl && /^https?:/i.test(raw)){ const su = await window.safeImageUrl(raw); if (su) final = su; } }catch{}
              const i = new Image();
              try{
                // Only set CORS for known CORS-enabled hosts; blob/data/same-origin need no CORS
                const isBlob = /^blob:/i.test(final);
                const isData = /^data:/i.test(final);
                if (!isBlob && !isData){
                  const u = new URL(final, location.href);
                  const same = (u.origin === location.origin);
                  const host = u.host.toLowerCase();
                  const shouldCORS = !same && /\.supabase\.co$/i.test(host);
                  if (shouldCORS) i.crossOrigin = 'anonymous';
                }
              }catch{}
              i.onload = () => res(i);
              i.onerror = () => rej(new Error('img load failed'));
              i.src = final;
            }catch(e){ rej(e); }
          })();
        });
      }
      function fileName(url){ try{ const u=new URL(url, location.origin); return u.pathname.split('/').pop()||'dosya'; }catch{ const parts=String(url|| '').split('/'); return parts[parts.length-1]||'dosya'; } }
      function renderExistingDocs(){
        docsPrev.innerHTML= '';
        let arr=[]; try{ arr = JSON.parse(row.documents_urls||'[]'); if(!Array.isArray(arr)) arr=[]; }catch{}
        arr.forEach(url=>{
          const card = document.createElement('div'); card.className='card'; card.style.padding='6px'; card.style.display='flex'; card.style.flexDirection='column'; card.style.gap='6px';
          if (isImg(url)){
            const img = document.createElement('img'); img.alt= ''; img.style.width='100%'; img.style.height='86px'; img.style.objectFit='cover'; img.style.borderRadius='8px';
            img.src = bust(url);
            card.appendChild(img);
          } else {
            const a=document.createElement('a'); a.href=url; a.target='_blank'; a.rel='noopener noreferrer'; a.textContent=fileName(url); a.style.wordBreak='break-all'; card.appendChild(a);
          }
          docsPrev.appendChild(card);
        });
      }
      dlBtn.addEventListener('click',(e)=>{ e.preventDefault(); try{ const data=canvas.toDataURL('image/jpeg', 0.92); try{ window.__idcardLastPng = data; }catch{} const a=document.createElement('a'); const fullName = `${(member.first_name||'').trim()} ${(member.last_name||'').trim()}`.trim() || (member.member_no||member.id||'uye'); let base = fullName.normalize('NFD').replace(/[\u0300-\u036f]/g,''); base = base.replace(/ı/g,'i').replace(/İ/g,'I').replace(/ş/g,'s').replace(/Ş/g,'S').replace(/ğ/g,'g').replace(/Ğ/g,'G').replace(/ç/g,'c').replace(/Ç/g,'C').replace(/ö/g,'o').replace(/Ö/g,'O').replace(/ü/g,'u').replace(/Ü/g,'U'); const safe = base.toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_+|_+$/g,''); a.download=`kimlik_${safe||'uye'}.jpg`; a.href=data; a.click(); }catch{} });
      printBtn.addEventListener('click',(e)=>{ e.preventDefault(); try{ const w=window.open('','_blank'); if(!w) return; const data=canvas.toDataURL('image/jpeg', 0.92); try{ window.__idcardLastPng = data; }catch{} w.document.write(`<img src="${data}" style="max-width:100%" onload="window.print();setTimeout(()=>window.close(),300);"/>`); w.document.close(); }catch{} });
      sendBtn.addEventListener('click', async (e)=>{ e.preventDefault(); try{ window.__idcardLastPng = canvas.toDataURL('image/jpeg', 0.92); await sendDigitalIdCard(member.id); closeModal(); alert('Kimlik gönderim kuyruğa alındı.'); }catch(err){ alert('Kimlik gönderilemedi: ' + (err?.message||String(err))); } });
      closeBtn.addEventListener('click', (e)=>{ e.preventDefault(); closeModal(); });

      (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
    }catch(e){ alert('Kimlik önizleme açılamadı: ' + (e?.message||String(e))); }
  }

  // Settings helper
  async function getSettingValue(key){
    try{ const { data } = await sb().from('settings').select('value').eq('key', key).maybeSingle(); return data?.value || ''; }catch{ return ''; }
  }

  // ========== SETTINGS ==========
  async function loadSettingsForm(){
    const panel = qs('#tab-settings'); if (!panel) return;
    // Try to find an existing form area; otherwise create one lightly
    let form = panel.querySelector('form');
    if (!form){
      form = document.createElement('form');
      form.id = 'settingsForm';
      form.style.display='grid'; form.style.gap='10px';
      panel.appendChild(form);
      const note = document.createElement('div'); note.className='muted'; note.textContent = 'Ayarlar'; form.appendChild(note);
    }

    // Fetch settings and populate matching inputs by name
    let rows = [];
    try{
      const { data, error } = await sb().from('settings').select('key, value');
      if (error) throw error;
      rows = data || [];
    }catch(e){ alert('Ayarlar yüklenemedi: ' + (e?.message||String(e))); }
    const byKey = new Map(rows.map(r=>[String(r.key), r.value]));

    // Helpers
    const norm = (s)=>String(s|| '').toLowerCase().replace(/\s+/g,' ').trim();
    const labelToKey = new Map(Object.entries({
      'recaptcha site key':'recaptcha_site_key',
      'recaptcha secret key':'recaptcha_secret_key',
      'kurum adı':'company_name',
      'slogan':'company_tagline',
      'logo url':'logo_url',
      'kart stili':'idcard_style',
      'kart arkaplan png (şablon yolu)':'idcard_bg_url',
      'doğrulama sayfası url (verify.html)':'verify_base_url'   // <-- EKLE
    }));
    labelToKey.set('e-posta gönderim servisi url (idcard)', 'email_edge_url');
    const findBtnByText = (text)=>{
      const t = norm(text);
      return Array.from(panel.querySelectorAll('button, input[type="button"], input[type="submit"]')).find(b=>norm(b.textContent||b.value)===t);
    };

    // Populate any input/select/textarea that has a name matching the key
    const inputs = Array.from(panel.querySelectorAll('input, select, textarea'));
    const bound = new Set();
    inputs.forEach(el => {
      const key = el.getAttribute('name') || el.getAttribute('data-setting') || el.id;
      if (!key) return;
      bound.add(el);
      const val = byKey.has(key) ? byKey.get(key) : '';
      if (el.type === 'checkbox'){
        el.checked = val === true || val === 'true' || val === '1' || val === 1;
      } else if (el.type === 'radio'){
        if (String(el.value) === String(val)) el.checked = true;
      } else if (val != null){
        el.value = String(val);
      }
    });
    // Populate via labels if inputs have no name/id
    Array.from(panel.querySelectorAll('label')).forEach(lbl =>{
      const k = labelToKey.get(norm(lbl.textContent)); if (!k) return;
      const el = lbl.querySelector('input,select,textarea'); if (!el || bound.has(el)) return;
      const val = byKey.has(k) ? byKey.get(k) : '';
      if (el.type === 'checkbox') el.checked = val === true || val === 'true' || val === '1' || val === 1;
      else if (el.type === 'radio'){ if (!el.checked) return; el.checked = String(el.value) === String(val); }
      else if (val != null) el.value = String(val);
      el.setAttribute('data-setting', k);
      bound.add(el);
    });

    // Ensure common fields exist visually (so they appear even if not in HTML)
    const ensureField = (label, name, type='text', placeholder='') => {
      if (form.querySelector(`[name="${name}"]`)) return;
      const row = document.createElement('label');
      row.style.display='grid';
      row.style.gap='6px';
      row.innerHTML = `<span style="font-weight:600">${label}</span>`;
      const input = document.createElement(type==='select'?'select':'input');
      if (type==='select'){
        const opt1=document.createElement('option'); opt1.value= ''; opt1.textContent='Varsayılan';
        const opt2=document.createElement('option'); opt2.value='vector'; opt2.textContent='Vector (Önerilir)';
        const opt3=document.createElement('option'); opt3.value='psd'; opt3.textContent='Arkaplan (PSD PNG)';
        input.appendChild(opt1); input.appendChild(opt2); input.appendChild(opt3);
      } else { input.type='text'; input.placeholder=placeholder; }
      input.name = name;
      input.style.padding='8px';
      input.style.border='1px solid #e5e7eb';
      input.style.borderRadius='6px';
      row.appendChild(input);
      form.appendChild(row);
    };
    ensureField('Kurum Adı', 'company_name');
    ensureField('Slogan', 'company_tagline');
    ensureField('Logo URL', 'logo_url');
    ensureField('Kart Stili', 'idcard_style', 'select');
    ensureField('Kart Arkaplan PNG (şablon yolu)', 'idcard_bg_url');
    ensureField('Doğrulama Sayfası URL (verify.html)', 'verify_base_url', 'text', 'https://ilkesen.org.tr/verify.html');
    ensureField('E‑posta Gönderim Servisi URL (ID Kart)', 'email_edge_url', 'text', 'https://YOUR-PROJECT-REF.functions.supabase.co/idcard');
    // verify_base_url alanı populate turundan sonra oluşturulduğu için burada set ediyoruz
    try {
      const el = form.querySelector('input[name="verify_base_url"]');
      if (el && byKey.has('verify_base_url')) {
        el.value = String(byKey.get('verify_base_url') || '');
        el.setAttribute('data-setting', 'verify_base_url');
      }
    } catch {}
    // Wire Save and Reload buttons if exist
    const saveBtn = panel.querySelector('#settingsSaveBtn') || panel.querySelector('button[data-action="save-settings"]') || findBtnByText('kaydet');
    const reloadBtn = panel.querySelector('#settingsReloadBtn') || panel.querySelector('button[data-action="reload-settings"]') || findBtnByText('yenile');
    if (reloadBtn && !reloadBtn.dataset.wired){ reloadBtn.dataset.wired='1'; reloadBtn.addEventListener('click', (e)=>{ e.preventDefault(); loadSettingsForm(); }); }
    if (saveBtn && !saveBtn.dataset.wired){
      saveBtn.dataset.wired='1';
      saveBtn.addEventListener('click', async (e)=>{
        e.preventDefault();
        // collect values
        const els = Array.from(panel.querySelectorAll('input, select, textarea'));
        const used = new Set();
        const payload = [];
        els.forEach(el => {
          let key = el.getAttribute('name') || el.getAttribute('data-setting') || el.id;
          if (!key){
            const lbl = el.closest('label'); if (lbl){ const lk = labelToKey.get(norm(lbl.textContent)); if (lk){ key = lk; el.setAttribute('data-setting', lk); } }
          }
          if (!key || used.has(key)) return; used.add(key);
          let value;
          if (el.type === 'checkbox') value = el.checked ? 'true' : 'false';
          else if (el.type === 'radio'){ if (!el.checked) return; value = el.value; }
          else value = el.value;
          payload.push({ key, value });
        });
        try{
          const { error } = await sb().from('settings').upsert(payload, { onConflict: 'key' });
          if (error) throw error;
          alert('Ayarlar kaydedildi');
          await loadSettingsForm();
        }catch(err){ alert('Ayarlar kaydedilemedi: ' + (err?.message||String(err))); }
      });
    }
    try {
      const el2 = form.querySelector('input[name="email_edge_url"]');
      if (el2 && byKey.has('email_edge_url')) {
        el2.value = String(byKey.get('email_edge_url') || '');
        el2.setAttribute('data-setting', 'email_edge_url');
      }
    } catch {}
  }

  // ========== PAGES ==========
  async function loadAdminPages(){
    const tbody = qs('#pagesTableBody'); if (!tbody) return; tbody.innerHTML= '';
    try{
      const table = tbody.closest('table');
      const headRow = table && table.tHead && table.tHead.rows && table.tHead.rows[0];
      if (headRow){
        headRow.innerHTML = `
          <th>Başlık</th>
          <th>Slug</th>
          <th>Durum</th>
          <th>Yayınlanma Tarihi</th>
          <th>Yayından Kaldırılma Tarihi</th>
          <th>İşlem</th>
        `;
      }
    }catch{}
    try{
      const { data, error } = await sb()
        .from('pages')
        .select('id, title, slug, status, published_at, unpublish_at')
        .order('title', { ascending: true });
      if (error) throw error;
      (data||[]).forEach(row => {
        const tr = document.createElement('tr');
        const statusMap = { published:'Yayımlandı', draft:'Taslak', scheduled:'Planlı', archived:'Arşivli', unpublished:'Yayından Kaldırıldı' };
        const statusTr = statusMap[String(row.status|| '').toLowerCase()] || (row.status|| '');
        tr.innerHTML = `
          <td>${escapeHtml(row.title|| '')}</td>
          <td>${escapeHtml(row.slug|| '')}</td>
          <td>${escapeHtml(statusTr)}</td>
          <td>${row.published_at ? new Date(row.published_at).toLocaleString('tr-TR') : '-'}</td>
          <td>${row.unpublish_at ? new Date(row.unpublish_at).toLocaleString('tr-TR') : '-'}</td>
          <td class="actions">
            <button class="btn btn-warning" data-edit-page="${row.id}">Düzenle</button>
          </td>`;
        tbody.appendChild(tr);
      });
       // Wire: Yeni Sayfa
    const addBtn = qs('#newPageBtn');
    if (addBtn && !addBtn.dataset.wired){
      addBtn.dataset.wired='1';
      addBtn.onclick = () => openPageModal({ id:null, title:'', slug:'', status:'draft', body:'', published_at:null, unpublish_at:null });
    }
      wirePagesRowActions();
    }catch(e){ alert('Sayfalar yüklenemedi: ' + (e?.message||String(e))); }
  }

  function wirePagesRowActions(){
    try{
      const tbody = qs('#pagesTableBody'); if (!tbody) return;

      // Edit
      tbody.querySelectorAll('button[data-edit-page]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-edit-page');
          try{
            const { data } = await sb().from('pages').select('*').eq('id', id).maybeSingle();
            openPageModal(data||{ id, title:'', slug:'', status:'draft', body:'', published_at:null, unpublish_at:null });
          }catch(e){ alert('Sayfa yüklenemedi: ' + (e?.message||String(e))); }
        });
      });
    }catch{}
  }
  

  function openPageModal(row){
    try{
      row = row || { id:null, title:'', slug:'', status:'draft', body:'', published_at:null, unpublish_at:null };
      $modalTitle().textContent = row.id ? 'Sayfayı Düzenle' : 'Yeni Sayfa';
      const form = $modalForm(); form.innerHTML= '';

      // Title
      const tLbl=document.createElement('label'); tLbl.style.display='grid'; tLbl.style.gap='6px'; tLbl.innerHTML='<span>Başlık</span>'; const tIn=document.createElement('input'); tIn.value=row.title|| ''; tLbl.appendChild(tIn); form.appendChild(tLbl);

      // Slug
      const sLbl=document.createElement('label'); sLbl.style.display='grid'; sLbl.style.gap='6px'; sLbl.innerHTML='<span>Slug</span>'; const sIn=document.createElement('input'); sIn.value=row.slug|| ''; sLbl.appendChild(sIn); form.appendChild(sLbl);

      // Body editor: Word-like toolbar + contentEditable area
      const bLbl=document.createElement('label'); bLbl.style.display='grid'; bLbl.style.gap='6px'; bLbl.innerHTML='<span>İçerik</span>';
      const tb=document.createElement('div'); tb.style.display='flex'; tb.style.flexWrap='wrap'; tb.style.gap='6px'; tb.style.margin='6px 0';
      const ed=document.createElement('div'); ed.contentEditable='true'; ed.className='card'; ed.style.minHeight='240px'; ed.style.padding='10px'; ed.style.overflow='auto';
      try{ ed.setAttribute('data-skip-safe-images','1'); }catch{}
      ed.innerHTML = (row.body|| '');
      function btn(label, title, on){ const b=document.createElement('button'); b.type='button'; b.className='btn btn-outline'; b.textContent=label; b.title=title; b.style.padding='6px 10px'; b.addEventListener('click', (e)=>{ e.preventDefault(); on(); ed.focus(); }); return b; }
      function applyInlineStyle(prop, val){
        const sel = window.getSelection(); if (!sel || !sel.rangeCount) return;
        const range = sel.getRangeAt(0);
        if (range.collapsed){
          const span=document.createElement('span'); span.style[prop]=val; span.appendChild(document.createTextNode('\u200b'));
          range.insertNode(span); sel.removeAllRanges(); const r=document.createRange(); r.selectNodeContents(span); r.collapse(false); sel.addRange(r);
        } else {
          const span=document.createElement('span'); span.style[prop]=val; try{ range.surroundContents(span); }catch{ document.execCommand('foreColor'); }
        }
      }
      function applyBlock(cmd){ document.execCommand(cmd,false,null); }
      // Font family
      const ff=document.createElement('select'); ff.className='btn btn-outline'; ['Default','Arial','Georgia','Tahoma','Times New Roman','Verdana','Courier New'].forEach(f=>{ const o=document.createElement('option'); o.value=(f==='Default')?'':f; o.textContent=f; ff.appendChild(o); }); ff.addEventListener('change',()=>{ if(ff.value) applyInlineStyle('fontFamily', ff.value); });
      // Font size (px)
      const fs=document.createElement('input'); fs.type='number'; fs.min='10'; fs.max='72'; fs.value='16'; fs.title='Yazı Boyutu (px)'; fs.className='btn btn-outline'; fs.style.width='84px';
      fs.addEventListener('change',()=> applyInlineStyle('fontSize', fs.value+'px'));
      // Line height
      const lh=document.createElement('select'); lh.className='btn btn-outline'; ['1.0','1.2','1.4','1.6','1.8','2.0'].forEach(v=>{ const o=document.createElement('option'); o.value=v; o.textContent='Satır: '+v; lh.appendChild(o); }); lh.addEventListener('change',()=> applyInlineStyle('lineHeight', lh.value));
      // Text color
      const col=document.createElement('input'); col.type='color'; col.className='btn btn-outline'; col.title='Yazı Rengi'; col.addEventListener('input',()=> applyInlineStyle('color', col.value));
      // Styles
      tb.appendChild(btn('B','Kalın', ()=> applyBlock('bold')));
      tb.appendChild(btn('I','İtalik', ()=> applyBlock('italic')));
      tb.appendChild(btn('U','Altı Çizili', ()=> applyBlock('underline')));
      // Alignment
      tb.appendChild(btn('Sol','Sola Hizala', ()=> applyBlock('justifyLeft')));
      tb.appendChild(btn('Ortala','Ortala', ()=> applyBlock('justifyCenter')));
      tb.appendChild(btn('Sağ','Sağa Hizala', ()=> applyBlock('justifyRight')));
      tb.appendChild(btn('İkiye','İki Yana Yasla', ()=> applyBlock('justifyFull')));
      // Lists
      tb.appendChild(btn('• Liste','Madde İşaretli', ()=> applyBlock('insertUnorderedList')));
      tb.appendChild(btn('1. Liste','Numaralı', ()=> applyBlock('insertOrderedList')));
      // Font family / size / line-height / color controls
      tb.appendChild(ff); tb.appendChild(fs); tb.appendChild(lh); tb.appendChild(col);
      bLbl.appendChild(tb); bLbl.appendChild(ed); form.appendChild(bLbl);
     
      // Status
      const stLbl=document.createElement('label'); stLbl.style.display='grid'; stLbl.style.gap='6px'; stLbl.innerHTML='<span>Durum</span>'; const stSel=document.createElement('select');
      const statusOpts = [
        { v:'draft', t:'Taslak' },
        { v:'scheduled', t:'Planlı' },
        { v:'published', t:'Yayımlandı' },
        { v:'archived', t:'Arşivli' },
        { v:'unpublished', t:'Yayından Kaldırıldı' },
      ];
      statusOpts.forEach(({v,t})=>{ const o=document.createElement('option'); o.value=v; o.textContent=t; if (v===(row.status||'draft')) o.selected=true; stSel.appendChild(o); });
      stLbl.appendChild(stSel); form.appendChild(stLbl);

      // Dates
      const pLbl=document.createElement('label'); pLbl.style.display='grid'; pLbl.style.gap='6px'; pLbl.innerHTML='<span>Yayın Tarihi</span>'; const pIn=document.createElement('input'); pIn.type='datetime-local'; pIn.value=row.published_at? new Date(row.published_at).toISOString().slice(0,16): ''; pLbl.appendChild(pIn); form.appendChild(pLbl);
      const uLbl=document.createElement('label'); uLbl.style.display='grid'; uLbl.style.gap='6px'; uLbl.innerHTML='<span>Yayından Kaldırma</span>'; const uIn=document.createElement('input'); uIn.type='datetime-local'; uIn.value=row.unpublish_at? new Date(row.unpublish_at).toISOString().slice(0,16): ''; uLbl.appendChild(uIn); form.appendChild(uLbl);

      // Actions
      const actions=document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
      const saveBtn=document.createElement('button'); saveBtn.className='btn btn-success'; saveBtn.textContent='Kaydet';
      const cancelBtn=document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
      actions.appendChild(saveBtn); actions.appendChild(cancelBtn);
      if (row.id){
        const delBtn=document.createElement('button'); delBtn.className='btn btn-danger'; delBtn.textContent='Sil'; actions.appendChild(delBtn);
        delBtn.addEventListener('click', async (e)=>{
          e.preventDefault();
          try{
            if (!confirm('Bu kaydı silmek istediğinize emin misiniz?')) return;
            const { error } = await sb().from('pages').delete().eq('id', row.id);
            if (error) throw error;
            closeModal(); if (typeof loadAdminPages==='function') await loadAdminPages();
          }catch(err){ alert('Silinemedi: ' + (err?.message||String(err))); }
        });
      }
      form.appendChild(actions);

      cancelBtn.addEventListener('click', (e)=>{ e.preventDefault(); closeModal(); });
      saveBtn.addEventListener('click', async (e)=>{
        e.preventDefault();
        try{
          const payload = {
            title: String(tIn.value|| '').trim(),
            slug: String(sIn.value|| '').trim(),
            body: String(ed.innerHTML|| '').trim(),
            status: String(stSel.value||'draft'),
            published_at: pIn.value ? new Date(pIn.value).toISOString() : null,
            unpublish_at: uIn.value ? new Date(uIn.value).toISOString() : null,
          };
          // Basic validation
          if (!payload.title) return alert('Başlık gerekli');
          if (!payload.slug) return alert('Slug gerekli');
          if (payload.status === 'scheduled' && !payload.published_at){ return alert('Planlı için Yayın Tarihi gerekli'); }
          if (payload.unpublish_at && payload.published_at && new Date(payload.unpublish_at) < new Date(payload.published_at)){
            return alert('Yayından kaldırma tarihi, yayın tarihinden önce olamaz');
          }
          if (payload.status === 'published' && !payload.published_at){ payload.published_at = new Date().toISOString(); }

          // Upsert by id
          const q = row.id ? sb().from('pages').update(payload).eq('id', row.id) : sb().from('pages').insert(payload).select('id').single();
          const { error } = await q; if (error) throw error;
          closeModal(); if (typeof loadAdminPages==='function') await loadAdminPages();
        }catch(err){ alert('Kaydedilemedi: ' + (err?.message||String(err))); }
      });

      (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
    }catch(e){ alert('Sayfa formu açılamadı: ' + (e?.message||String(e))); }
  }

  function openEmailCompose(to, subjOrig, msgId, afterSend){
    const subjDefault = /^\s*re:/i.test(subjOrig|| '') ? (subjOrig|| '') : `Re: ${subjOrig||'Yanıt'}`;

    $modalTitle().textContent = 'E‑posta Gönder';
    const form = $modalForm(); form.innerHTML = '';
    const wrap = document.createElement('div'); wrap.style.display='grid'; wrap.style.gap='8px';
    form.appendChild(wrap);

    // To (read-only) - mailto: temizle
    const toLbl = document.createElement('label'); toLbl.style.display='grid'; toLbl.style.gap='6px'; toLbl.innerHTML = '<span>Alıcı</span>';
    const toIn = document.createElement('input');
    toIn.type='email';
    toIn.value = String(to|| '').replace(/^mailto:/i,'').trim(); // <-- önemli
    toIn.disabled = true;
    toLbl.appendChild(toIn); wrap.appendChild(toLbl);

    // Subject
    const sLbl = document.createElement('label'); sLbl.style.display='grid'; sLbl.style.gap='6px'; sLbl.innerHTML = '<span>Konu</span>';
    const sIn = document.createElement('input'); sIn.type='text'; sIn.value = subjDefault; sLbl.appendChild(sIn); wrap.appendChild(sLbl);

    // Message
    const mLbl = document.createElement('label'); mLbl.style.display='grid'; mLbl.style.gap='6px'; mLbl.innerHTML = '<span>Mesaj</span>';
    const mIn = document.createElement('textarea'); mIn.rows = 10; mIn.placeholder = 'Yanıtınızı buraya yazınız...';
    mLbl.appendChild(mIn); wrap.appendChild(mLbl);

    // Actions
    const actions = document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
    const sendBtn = document.createElement('button'); sendBtn.className='btn btn-success'; sendBtn.textContent='Gönder';
    const cancelBtn = document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
    actions.appendChild(sendBtn); actions.appendChild(cancelBtn); wrap.appendChild(actions);

    cancelBtn.addEventListener('click', (e)=>{ e.preventDefault(); closeModal(); });

    // Send
    sendBtn.addEventListener('click', async (e)=>{
      e.preventDefault();

      const cleanTo = String(to|| '').replace(/^mailto:/i,'').trim();
      const subject = String(sIn.value|| '').trim();
      const message = String(mIn.value|| '').trim();

      if (!cleanTo || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanTo)){
        alert('Geçerli bir alıcı e‑posta bulunamadı.');
        return;
      }
      if (!subject || !message){
        alert('Konu ve mesaj gerekli');
        return;
      }

      sendBtn.disabled = true;
      try{
        // JWT’yi header’a ekle
        const { data: sess } = await sb().auth.getSession();
        const token = sess?.session?.access_token || '';
        const headers = token ? { Authorization: `Bearer ${token}` } : {};

        const { data, error } = await sb().functions.invoke('admin_send_mail', {
          body: { to: cleanTo, subject, message },
          headers
        });

        if (error) {
          let msg = error?.message || 'Bilinmeyen hata';
          try {
            const resp = error?.context?.response || error?.context;
            if (resp && typeof resp.text === 'function') {
              const details = await resp.text();
              msg += ' • ' + details;
            }
          } catch {}
          throw new Error(msg);
        }
        if (data && data.ok === false){ throw new Error(data.error || 'Gönderim hatası'); }

        if (msgId) {
          try{
            const repliedBy = (currentAdmin && currentAdmin.email) ? String(currentAdmin.email) : '';
            await sb().from('messages').update({
              is_replied: true,
              replied_at: new Date().toISOString(),
              replied_by: repliedBy,
              reply_subject: subject,
              reply_body: message
            }).eq('id', msgId);
          }catch{
            try{ await sb().from('messages').update({ is_replied: true }).eq('id', msgId); }catch{}
          }
        }
        closeModal();
        try{ alert('Mesajınız cevaplanmıştır.'); }catch{}
        if (typeof afterSend === 'function') try{ await afterSend(); }catch{}
      }catch(err){
        alert('E‑posta gönderilemedi: ' + (err?.message||String(err)));
      } finally {
        sendBtn.disabled = false;
      }
    });

    try{ (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal())); }catch{}
  }
 
  // ========== NEWS ==========
  async function loadAdminNews(){
    const tbody = $newsTBody(); if (!tbody) return; tbody.innerHTML= '';
    // Hard-set THEAD to ensure correct column order
    try{
      const table = tbody.closest('table');
      const headRow = table && table.tHead && table.tHead.rows && table.tHead.rows[0];
      if (headRow){
        headRow.innerHTML = `
          <th>Başlık</th>
          <th>Özet</th>
          <th>Yayınlanma Durumu</th>
          <th>Yayınlanma Tarihi</th>
          <th>Yayından Kaldırılma Tarihi</th>
          <th>İşlem</th>
        `;
      }
    }catch{}
    try{
      const { data, error } = await sb()
        .from('news')
        .select('id, title, summary, image_url, status, published_at, unpublish_at')
        .order('published_at', { ascending: false, nullsFirst: true });
      if (error) throw error;
      (data||[]).forEach(row => {
        const tr = document.createElement('tr');
        const thumb = row.image_url ? `<img src="${escapeHtml(bust(row.image_url))}" alt="${escapeHtml(row.title||'')}" style="width:40px;height:40px;object-fit:cover;border-radius:6px;margin-right:8px;vertical-align:middle"/>` : '';
        const statusMap = { published:'Yayımlandı', draft:'Taslak', scheduled:'Planlı', archived:'Arşivli', unpublished:'Yayından Kaldırıldı', active:'Yayımlandı' };
        const statusTr = statusMap[String(row.status|| '').toLowerCase()] || (row.status|| '');
        tr.innerHTML = `
          <td>${thumb}<span style="vertical-align:middle">${escapeHtml(row.title|| '')}</span></td>
          <td>${escapeHtml(row.summary|| '')}</td>
          <td>${escapeHtml(statusTr)}</td>
          <td>${row.published_at ? new Date(row.published_at).toLocaleString('tr-TR') : '-'}</td>
          <td>${row.unpublish_at ? new Date(row.unpublish_at).toLocaleString('tr-TR') : '-'}</td>
          <td class="actions">
            <button class="btn btn-warning" data-edit-news="${row.id}">Düzenle</button>
            ${String(row.status|| '').toLowerCase()==='published' ? `<button class="btn btn-danger" data-unpub-news="${row.id}">Yayından Kaldır</button>` : `<button class="btn btn-success" data-pub-news="${row.id}">Yayınla</button>`}
          </td>`;
        tbody.appendChild(tr);
      });
      const addBtn = qs('#newNewsBtn'); if (addBtn && !addBtn.dataset.wired){ addBtn.dataset.wired='1'; addBtn.onclick = ()=> openNewsModal({ id:null, title:'', summary:'', image_url:'', status:'draft' }); }
      wireNewsRowActions();
    }catch(e){ alert('Haberler yüklenemedi: ' + (e?.message||String(e))); }
  }

  // ========== ANNOUNCEMENTS ==========
  async function loadAdminAnnouncements(){
    const tbody = $annTBody(); if (!tbody) return; tbody.innerHTML= '';
    // Hard-set THEAD to ensure correct column order
    try{
      const table = tbody.closest('table');
      const headRow = table && table.tHead && table.tHead.rows && table.tHead.rows[0];
      if (headRow){
        headRow.innerHTML = `
          <th>Başlık</th>
          <th>İçerik</th>
          <th>Yayınlanma Durumu</th>
          <th>Yayınlanma Tarihi</th>
          <th>Yayından Kaldırılma Tarihi</th>
          <th>İşlem</th>
        `;
      }
    }catch{}
    try{
      const { data, error } = await sb()
        .from('announcements')
        .select('id, title, body, image_url, status, published_at, unpublish_at')
        .order('published_at', { ascending: false, nullsFirst: true });
      if (error) throw error;
      (data||[]).forEach(row => {
        const tr = document.createElement('tr');
        const thumb = row.image_url ? `<img src="${escapeHtml(bust(row.image_url))}" alt="${escapeHtml(row.title||'')}" style="width:40px;height:40px;object-fit:cover;border-radius:6px;margin-right:8px;vertical-align:middle"/>` : '';
        const statusMap = { published:'Yayımlandı', draft:'Taslak', scheduled:'Planlı', archived:'Arşivli', unpublished:'Yayından Kaldırıldı', active:'Yayımlandı' };
        const statusTr = statusMap[String(row.status|| '').toLowerCase()] || (row.status|| '');
        tr.innerHTML = `
          <td>${thumb}<span style="vertical-align:middle">${escapeHtml(row.title|| '')}</span></td>
          <td>${escapeHtml(row.body|| '')}</td>
          <td>${escapeHtml(statusTr)}</td>
          <td>${row.published_at ? new Date(row.published_at).toLocaleString('tr-TR') : '-'}</td>
          <td>${row.unpublish_at ? new Date(row.unpublish_at).toLocaleString('tr-TR') : '-'}</td>
          <td class="actions">
            <button class="btn btn-warning" data-edit-ann="${row.id}">Düzenle</button>
            ${String(row.status|| '').toLowerCase()==='published' ? `<button class="btn btn-danger" data-unpub-ann="${row.id}">Yayından Kaldır</button>` : `<button class="btn btn-success" data-pub-ann="${row.id}">Yayınla</button>`}
          </td>`;
        tbody.appendChild(tr);
      });
      const addBtn = qs('#newAnnBtn'); if (addBtn && !addBtn.dataset.wired){ addBtn.dataset.wired='1'; addBtn.onclick = ()=> openAnnModal({ id:null, title:'', body:'', image_url:'', status:'draft' }); }
      wireAnnRowActions();
    }catch(e){ alert('Duyurular yüklenemedi: ' + (e?.message||String(e))); }
  }

  async function loadAdminMessages(){
    const tbody = $msgsTBody(); if (!tbody) return;
    tbody.innerHTML = '';
    messagesById = new Map();
    try{
      const refreshBtn = document.getElementById('refreshMsgsBtn');
      if (refreshBtn && !refreshBtn.dataset.wired){
        refreshBtn.dataset.wired = '1';
        refreshBtn.addEventListener('click', (e)=>{ e.preventDefault(); loadAdminMessages(); });
      }
    }catch{}

    // Wire category filter buttons
    try{
      const group = document.getElementById('msgCatFilters');
      if (group && !group.dataset.wired){
        group.dataset.wired = '1';
        group.addEventListener('click', (ev)=>{
          const btn = ev.target && ev.target.closest ? ev.target.closest('[data-msg-cat]') : null;
          if (!btn) return;
          ev.preventDefault();
          currentMsgCategory = btn.getAttribute('data-msg-cat') || 'all';
          Array.from(group.querySelectorAll('button')).forEach(b=> b.classList.toggle('active', b===btn));
          loadAdminMessages();
        });
      }
    }catch{}

    // Load messages with server-side category filter if column exists; fallback to client-side
    let rows = [];
    try{
      let q = sb().from('messages').select('*').order('created_at', { ascending:false, nullsFirst:true });
      if (currentMsgCategory && currentMsgCategory !== 'all') q = q.eq('category', currentMsgCategory);
      { const { data, error } = await q; if (error) throw error; rows = data || []; }
    }catch(e){
      // Likely category column missing; fallback without filter and filter client-side by subject prefix
      try{
        const { data, error } = await sb().from('messages').select('id, name, email, subject, body, created_at, is_read, is_replied');
        if (error) throw error; rows = data || [];
      }catch(err){ alert('Mesajlar yüklenemedi: ' + (err?.message||String(err))); return; }
    }

    // Client-side filter if needed
    function deriveCategory(r){
      if (r && r.category) return String(r.category||'');
      const s = String(r?.subject||'');
      if (/^\s*\[Avukatl[ıi]k Talebi\]/i.test(s)) return 'lawyer';
      if (/^\s*\[Üyelik Başvurusu\]/i.test(s)) return 'membership';
      // Heuristics for other types can be added later
      return 'contact';
    }
    const filtered = (currentMsgCategory && currentMsgCategory !== 'all')
      ? rows.filter(r => (r?.category ? String(r.category) === currentMsgCategory : deriveCategory(r) === currentMsgCategory))
      : rows;

    try{
      filtered.forEach(row => {
        try{ if (row && row.id) messagesById.set(row.id, row); }catch{}
        const tr = document.createElement('tr');
        // Category + status badges
        const map = { contact:'İletişim', email:'Mail', membership:'Üyelik Başvurusu', lawyer:'Avukatlık Talebi', meeting:'Görüşme Talebi', all:'Tümü' };
        const cat = (row && row.category) ? String(row.category) : deriveCategory(row);
        const catKey = (cat && map[cat]) ? cat : null;
        const catClass = catKey && ['all','contact','email','membership','lawyer','meeting'].includes(catKey) ? `badge-${catKey}` : 'badge-neutral';
        let badges = '';
        if (catKey) badges += `<span class="badge ${catClass}">${escapeHtml(map[catKey])}</span>`;
        if (row && row.is_read) badges += `<span class=\"badge badge-neutral\">Okundu</span>`;
        if (row && row.is_replied) badges += `<span class=\"badge badge-neutral\">Cevaplandı</span>`;
        const tagHtml = badges ? `<div class=\"badges\">${badges}</div>` : '';
        const when = row && row.created_at ? new Date(row.created_at).toLocaleString('tr-TR') : '-';
        const subjectRaw = String(row?.subject || '');
        const subjectClean = subjectRaw.replace(/^\s*\[[^\]]+\]\s*/, '').trim();
        tr.innerHTML = `
          <td>${escapeHtml(row?.name || '')}</td>
          <td>${escapeHtml(row?.email || '')}</td>
          <td>${escapeHtml(subjectClean)}${tagHtml}</td>
          <td>${when}</td>
          <td class="actions">
            <button class="btn btn-warning" data-view-msg="${row.id}">Oku</button>
            <button class="btn btn-success" data-reply-msg="${row.id}">Cevapla</button>
            ${row && row.is_read ? '' : `<button class=\"btn btn-outline\" data-read-msg=\"${row.id}\">Okundu</button>`}
            <button class="btn btn-danger" data-del-msg="${row.id}">Sil</button>
          </td>
        `;
        tbody.appendChild(tr);
      });
      wireMsgsRowActions();
    }catch(e){
      alert('Mesajlar yüklenemedi: ' + (e?.message||String(e)));
    }
  }

  function wireMsgsRowActions(){
    const tbody = $msgsTBody(); if (!tbody) return;
    try{
      tbody.querySelectorAll('button[data-view-msg]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-view-msg');
          const row = messagesById.get(id);
          if (!row) return;
          try{
            if (!row.is_read){
              await sb().from('messages').update({ is_read: true }).eq('id', id);
              row.is_read = true;
            }
          }catch{}
          await openMessageModal(row);
          try{ await loadAdminMessages(); }catch{}
        });
      });
      tbody.querySelectorAll('button[data-read-msg]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-read-msg');
          try{ await sb().from('messages').update({ is_read: true }).eq('id', id); }catch(e){ alert('Güncellenemedi: ' + (e?.message||String(e))); }
          try{ await loadAdminMessages(); }catch{}
        });
      });
      tbody.querySelectorAll('button[data-reply-msg]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-reply-msg');
          const row = messagesById.get(id);
          if (!row) return;
          try{
            if (!row.is_read){
              await sb().from('messages').update({ is_read: true }).eq('id', id);
              row.is_read = true;
            }
          }catch{}
          openEmailCompose(row.email, row.subject, id, async ()=>{ try{ await loadAdminMessages(); }catch{} });
          try{ await loadAdminMessages(); }catch{}
        });
      });
      tbody.querySelectorAll('button[data-del-msg]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-del-msg');
          if (!id) return;
          if (!confirm('Bu mesajı silmek istediğinize emin misiniz?')) return;
          try{
            await sb().from('messages').delete().eq('id', id);
          }catch(e){ alert('Silinemedi: ' + (e?.message||String(e))); return; }
          try{ await loadAdminMessages(); }catch{}
        });
      });
    }catch{}
  }

  async function openMessageModal(row){
    try{
      let fullRow = row;
      try{
        if (row && row.id){
          const { data } = await sb().from('messages').select('*').eq('id', row.id).maybeSingle();
          if (data) fullRow = data;
        }
      }catch{}

      $modalTitle().textContent = 'Mesaj';
      const form = $modalForm();
      form.innerHTML = '';

      const wrap = document.createElement('div');
      wrap.style.display = 'grid';
      wrap.style.gap = '8px';
      form.appendChild(wrap);

      function roInput(label, val, type){
        const lbl = document.createElement('label');
        lbl.style.display='grid';
        lbl.style.gap='6px';
        lbl.innerHTML = `<span>${escapeHtml(label)}</span>`;
        const inp = document.createElement('input');
        inp.type = type || 'text';
        inp.value = String(val || '');
        inp.disabled = true;
        lbl.appendChild(inp);
        return lbl;
      }

      // Compute data used across rows
      const map = { contact:'İletişim', email:'Mail', membership:'Üyelik Başvurusu', lawyer:'Avukatlık Talebi', meeting:'Görüşme Talebi' };
      const subjStr = String(fullRow?.subject || '');
      const subjClean = subjStr.replace(/^\s*\[[^\]]+\]\s*/,'').trim();
      let cat = String(fullRow?.category||'');
      if (!cat){
        if (/^\s*\[Avukatl[ıi]k Talebi\]/i.test(subjStr)) cat = 'lawyer';
        else if (/^\s*\[Üyelik Başvurusu\]/i.test(subjStr)) cat = 'membership';
        else cat = 'contact';
      }
      const catLabel = map[cat] || '-';

      function parseMessageBody(raw){
        const out = { ad:null, soyad:null, uyeNo:null, telefon:null, kurum:null, aciklama:null };
        const txt = String(raw||'');
        const lines = txt.split(/\r?\n/);
        lines.forEach(line => {
          let m;
          if ((m = line.match(/^\s*Ad\s*:\s*(.*)$/i))) out.ad = m[1].trim();
          else if ((m = line.match(/^\s*Soyad\s*:\s*(.*)$/i))) out.soyad = m[1].trim();
          else if ((m = line.match(/^\s*Üye\s*No\s*:\s*(.*)$/i))) out.uyeNo = m[1].trim();
          else if ((m = line.match(/^\s*Uye\s*No\s*:\s*(.*)$/i))) out.uyeNo = m[1].trim();
          else if ((m = line.match(/^\s*Telefon\s*:\s*(.*)$/i))) out.telefon = m[1].trim();
          else if ((m = line.match(/^\s*Kurum\s*:\s*(.*)$/i))) out.kurum = m[1].trim();
        });
        const idxDesc = lines.findIndex(l => /^\s*Açıklama\s*:/i.test(l) || /^\s*Aciklama\s*:/i.test(l));
        if (idxDesc >= 0){ out.aciklama = lines.slice(idxDesc+1).join('\n').trim(); }
        else { out.aciklama = txt.trim(); }
        return out;
      }

      const parsed = parseMessageBody(fullRow?.body || '');
      const adSoyad = (fullRow?.name || [parsed.ad, parsed.soyad].filter(Boolean).join(' ')).trim() || '-';

      // Order: Kategori, Tarih, Kurum, Ad Soyad, Üye No, E‑posta, Telefon, Konu, Mesaj
      wrap.appendChild(roInput('Kategori', catLabel, 'text'));
      wrap.appendChild(roInput('Tarih', fullRow?.created_at ? new Date(fullRow.created_at).toLocaleString('tr-TR') : '-', 'text'));
      wrap.appendChild(roInput('Kurum', parsed.kurum || '-', 'text'));
      wrap.appendChild(roInput('Ad Soyad', adSoyad, 'text'));
      wrap.appendChild(roInput('Üye No', parsed.uyeNo || '-', 'text'));
      wrap.appendChild(roInput('E‑posta', fullRow?.email || '-', 'email'));
      wrap.appendChild(roInput('Telefon', parsed.telefon || '-', 'tel'));

      // Konu with category badge (color-coded)
      {
        const lbl = document.createElement('label');
        lbl.style.display='grid';
        lbl.style.gap='6px';
        const bar = document.createElement('div');
        bar.style.display='flex'; bar.style.alignItems='center'; bar.style.gap='8px';
        const title = document.createElement('span'); title.textContent = 'Konu';
        const badge = document.createElement('span');
        const catKey = (cat && map[cat]) ? cat : null;
        const catClass = catKey && ['all','contact','email','membership','lawyer','meeting'].includes(catKey) ? `badge-${catKey}` : 'badge-neutral';
        badge.className = `badge ${catClass}`; badge.textContent = catLabel;
        bar.appendChild(title); bar.appendChild(badge);
        lbl.appendChild(bar);
        const inp = document.createElement('input'); inp.type='text'; inp.value = subjClean; inp.disabled = true; lbl.appendChild(inp);
        wrap.appendChild(lbl);
      }

      // Mesaj: only description
      {
        const mLbl = document.createElement('label');
        mLbl.style.display='grid'; mLbl.style.gap='6px';
        mLbl.innerHTML = '<span>Mesaj</span>';
        const ta = document.createElement('textarea'); ta.rows = 12; ta.value = String(parsed.aciklama || ''); ta.disabled = true; mLbl.appendChild(ta);
        wrap.appendChild(mLbl);
      }

      try{
        const replyBody = (fullRow && (fullRow.reply_body ?? fullRow.replyBody ?? fullRow.reply_message ?? fullRow.replyMessage)) || '';
        const replySubj = (fullRow && (fullRow.reply_subject ?? fullRow.replySubject)) || '';
        const repliedAt = (fullRow && (fullRow.replied_at ?? fullRow.repliedAt)) || null;
        const repliedBy = (fullRow && (fullRow.replied_by ?? fullRow.repliedBy)) || '';
        if (String(replyBody|| '').trim() || String(replySubj|| '').trim() || repliedAt || String(repliedBy|| '').trim()){
          wrap.appendChild(roInput('Yanıtlayan', repliedBy || '-', 'text'));
          wrap.appendChild(roInput('Yanıt Tarihi', repliedAt ? new Date(repliedAt).toLocaleString('tr-TR') : '-', 'text'));
          wrap.appendChild(roInput('Yanıt Konusu', replySubj || '-', 'text'));
          const rLbl = document.createElement('label');
          rLbl.style.display='grid';
          rLbl.style.gap='6px';
          rLbl.innerHTML = '<span>Verilen Yanıt</span>';
          const rTa = document.createElement('textarea');
          rTa.rows = 10;
          rTa.value = String(replyBody || '');
          rTa.disabled = true;
          rLbl.appendChild(rTa);
          wrap.appendChild(rLbl);
        }
      }catch{}

      const actions = document.createElement('div');
      actions.style.display = 'flex';
      actions.style.gap = '8px';
      const replyBtn = document.createElement('button');
      replyBtn.className = 'btn btn-success';
      replyBtn.textContent = 'Cevapla';
      const delBtn = document.createElement('button');
      delBtn.className = 'btn btn-danger';
      delBtn.textContent = 'Sil';
      const closeBtn = document.createElement('button');
      closeBtn.className = 'btn btn-danger';
      closeBtn.textContent = 'Kapat';
      actions.appendChild(replyBtn);
      actions.appendChild(delBtn);
      actions.appendChild(closeBtn);
      wrap.appendChild(actions);

      closeBtn.addEventListener('click', (e)=>{ e.preventDefault(); closeModal(); });
      replyBtn.addEventListener('click', (e)=>{
        e.preventDefault();
        openEmailCompose(fullRow?.email, fullRow?.subject, fullRow?.id, async ()=>{ try{ await loadAdminMessages(); }catch{} });
      });
      delBtn.addEventListener('click', async (e)=>{
        e.preventDefault();
        const id = fullRow?.id;
        if (!id) return;
        if (!confirm('Bu mesajı silmek istediğinize emin misiniz?')) return;
        try{
          await sb().from('messages').delete().eq('id', id);
          closeModal();
          await loadAdminMessages();
        }catch(err){ alert('Silinemedi: ' + (err?.message||String(err))); }
      });

      (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
    }catch(e){
      alert('Mesaj açılamadı: ' + (e?.message||String(e)));
    }
  }

  function wireNewsRowActions(){
    try{
      const tbody = $newsTBody(); if (!tbody) return;
      tbody.querySelectorAll('button[data-edit-news]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-edit-news');
          try{ const { data } = await sb().from('news').select('*').eq('id', id).maybeSingle(); openNewsModal(data||{ id:null, title:'', summary:'', image_url:'', status:'draft' }); }catch(e){ alert('Haber yüklenemedi: ' + (e?.message||String(e))); }
        });
      });
      // Publish / Unpublish
      tbody.querySelectorAll('button[data-pub-news]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-pub-news');
          try{ const { error } = await sb().from('news').update({ status:'published', published_at: new Date().toISOString(), unpublish_at: null }).eq('id', id); if (error) throw error; await loadAdminNews(); }catch(e){ alert('Yayınlanamadı: '+(e?.message||String(e))); }
        });
      });
      tbody.querySelectorAll('button[data-unpub-news]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-unpub-news');
          try{ const { error } = await sb().from('news').update({ status:'unpublished', unpublish_at: new Date().toISOString() }).eq('id', id); if (error) throw error; await loadAdminNews(); }catch(e){ alert('Yayından kaldırılamadı: '+(e?.message||String(e))); }
        });
      });
    }catch{}
  }

  function wireAnnRowActions(){
    try{
      const tbody = $annTBody(); if (!tbody) return;
      tbody.querySelectorAll('button[data-edit-ann]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-edit-ann');
          try{ const { data } = await sb().from('announcements').select('*').eq('id', id).maybeSingle(); openAnnModal(data||{ id:null, title:'', body:'', image_url:'', status:'draft' }); }catch(e){ alert('Duyuru yüklenemedi: ' + (e?.message||String(e))); }
        });
      });
      // Publish / Unpublish
      tbody.querySelectorAll('button[data-pub-ann]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-pub-ann');
          try{ const { error } = await sb().from('announcements').update({ status:'published', published_at: new Date().toISOString(), unpublish_at: null }).eq('id', id); if (error) throw error; await loadAdminAnnouncements(); }catch(e){ alert('Yayınlanamadı: '+(e?.message||String(e))); }
        });
      });
      tbody.querySelectorAll('button[data-unpub-ann]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-unpub-ann');
          try{ const { error } = await sb().from('announcements').update({ status:'unpublished', unpublish_at: new Date().toISOString() }).eq('id', id); if (error) throw error; await loadAdminAnnouncements(); }catch(e){ alert('Yayından kaldırılamadı: '+(e?.message||String(e))); }
        });
      });
    }catch{}
  }

  function openNewsModal(row){
    try{
      row = row || { id:null, title:'', summary:'', image_url:'', status:'draft', published_at:null, unpublish_at:null };
      $modalTitle().textContent = row.id ? 'Haberi Düzenle' : 'Yeni Haber';
      const form = $modalForm(); form.innerHTML= '';

      // Title
      const tLbl=document.createElement('label'); tLbl.style.display='grid'; tLbl.style.gap='6px'; tLbl.innerHTML='<span>Başlık</span>'; const tIn=document.createElement('input'); tIn.value=row.title|| ''; tLbl.appendChild(tIn); form.appendChild(tLbl);
      // Summary
      const sLbl=document.createElement('label'); sLbl.style.display='grid'; sLbl.style.gap='6px'; sLbl.innerHTML='<span>Özet</span>'; const sIn=document.createElement('textarea'); sIn.rows=3; sIn.value=row.summary|| ''; sLbl.appendChild(sIn); form.appendChild(sLbl);
      // Body editor: Word-like toolbar + contentEditable area
      const bLbl=document.createElement('label'); bLbl.style.display='grid'; bLbl.style.gap='6px'; bLbl.innerHTML='<span>İçerik</span>';
      const tb=document.createElement('div'); tb.style.display='flex'; tb.style.flexWrap='wrap'; tb.style.gap='6px'; tb.style.margin='6px 0';
      const ed=document.createElement('div'); ed.contentEditable='true'; ed.className='card'; ed.style.minHeight='240px'; ed.style.padding='10px'; ed.style.overflow='auto'; ed.innerHTML = (row.body|| '');
      function btn(label, title, on){ const b=document.createElement('button'); b.type='button'; b.className='btn btn-outline'; b.textContent=label; b.title=title; b.style.padding='6px 10px'; b.addEventListener('click', (e)=>{ e.preventDefault(); on(); ed.focus(); }); return b; }
      function applyInlineStyle(prop, val){
        const sel = window.getSelection(); if (!sel || !sel.rangeCount) return;
        const range = sel.getRangeAt(0);
        if (range.collapsed){
          const span=document.createElement('span'); span.style[prop]=val; span.appendChild(document.createTextNode('\u200b'));
          range.insertNode(span); sel.removeAllRanges(); const r=document.createRange(); r.selectNodeContents(span); r.collapse(false); sel.addRange(r);
        } else {
          const span=document.createElement('span'); span.style[prop]=val; try{ range.surroundContents(span); }catch{ document.execCommand('foreColor'); }
        }
      }
      function applyBlock(cmd){ document.execCommand(cmd,false,null); }
      // Font family
      const ff=document.createElement('select'); ff.className='btn btn-outline'; ['Default','Arial','Georgia','Tahoma','Times New Roman','Verdana','Courier New'].forEach(f=>{ const o=document.createElement('option'); o.value=(f==='Default')?'':f; o.textContent=f; ff.appendChild(o); }); ff.addEventListener('change',()=>{ if(ff.value) applyInlineStyle('fontFamily', ff.value); });
      // Font size (px)
      const fs=document.createElement('input'); fs.type='number'; fs.min='10'; fs.max='72'; fs.value='16'; fs.title='Yazı Boyutu (px)'; fs.className='btn btn-outline'; fs.style.width='84px'; fs.addEventListener('change',()=> applyInlineStyle('fontSize', fs.value+'px'));
      // Line height
      const lh=document.createElement('select'); lh.className='btn btn-outline'; ['1.0','1.2','1.4','1.6','1.8','2.0'].forEach(v=>{ const o=document.createElement('option'); o.value=v; o.textContent='Satır: '+v; lh.appendChild(o); }); lh.addEventListener('change',()=> applyInlineStyle('lineHeight', lh.value));
      // Text color
      const col=document.createElement('input'); col.type='color'; col.className='btn btn-outline'; col.title='Yazı Rengi'; col.addEventListener('input',()=> applyInlineStyle('color', col.value));
      // Styles
      tb.appendChild(btn('B','Kalın', ()=> applyBlock('bold')));
      tb.appendChild(btn('I','İtalik', ()=> applyBlock('italic')));
      tb.appendChild(btn('U','Altı Çizili', ()=> applyBlock('underline')));
      // Alignment
      tb.appendChild(btn('Sol','Sola Hizala', ()=> applyBlock('justifyLeft')));
      tb.appendChild(btn('Ortala','Ortala', ()=> applyBlock('justifyCenter')));
      tb.appendChild(btn('Sağ','Sağa Hizala', ()=> applyBlock('justifyRight')));
      tb.appendChild(btn('İkiye','İki Yana Yasla', ()=> applyBlock('justifyFull')));
      // Lists
      tb.appendChild(btn('• Liste','Madde İşaretli', ()=> applyBlock('insertUnorderedList')));
      tb.appendChild(btn('1. Liste','Numaralı', ()=> applyBlock('insertOrderedList')));
      // Font family / size / line-height / color controls
      tb.appendChild(ff); tb.appendChild(fs); tb.appendChild(lh); tb.appendChild(col);
      bLbl.appendChild(tb); bLbl.appendChild(ed); form.appendChild(bLbl);
      // Yazı Boyutu (px)
      const fontSizeInput = document.createElement('input');
      fontSizeInput.type = 'number';
      fontSizeInput.min = '10';
      fontSizeInput.max = '72';
      fontSizeInput.value = '16';
      fontSizeInput.title = 'Yazı Boyutu (px)';
      fontSizeInput.className = 'btn btn-outline';
      fontSizeInput.style.width = '84px';
      fontSizeInput.addEventListener('change', () => {
        const v = Number(fontSizeInput.value || 16);
        applyInlineStyle('fontSize', Math.max(10, Math.min(72, v)) + 'px');
      });
      tb.appendChild(fontSizeInput);
      // Image
      const iLbl=document.createElement('label'); iLbl.style.display='grid'; iLbl.style.gap='6px'; iLbl.innerHTML='<span>Kapak Görseli</span>';
      const iIn=document.createElement('input'); iIn.placeholder='https://... (opsiyonel)'; iIn.value=row.image_url|| '';
      const iFile=document.createElement('input'); iFile.type='file'; iFile.accept='image/*';
      const iWrap=document.createElement('div'); iWrap.style.display='grid'; iWrap.style.gap='6px'; iWrap.appendChild(iIn); iWrap.appendChild(iFile); iLbl.appendChild(iWrap); form.appendChild(iLbl);
      // Cover crop controls (for slider)
      const coverLbl = document.createElement('div'); coverLbl.className='card'; coverLbl.style.padding='10px'; coverLbl.style.display='grid'; coverLbl.style.gap='8px'; coverLbl.style.margin='8px 0';
      const coverTitle = document.createElement('div'); coverTitle.textContent = 'Kapak (Slider) Kadrajı'; coverTitle.style.fontWeight='600'; coverLbl.appendChild(coverTitle);
      const initCover = parsePhotoMeta(row.cover_image_url|| '');
          const cZoom = document.createElement('input'); cZoom.type='hidden'; cZoom.value = String(Math.max(1, initCover.z || 1).toFixed(2)); cZoom.setAttribute('data-cover-zoom','1');
    const cOy = document.createElement('input'); cOy.type='hidden'; cOy.value = String(initCover.oy || 0); cOy.setAttribute('data-cover-oy','1');
    const cOx = document.createElement('input'); cOx.type='hidden'; cOx.value = String(initCover.ox || 0); cOx.setAttribute('data-cover-ox','1');      const initCoverM = parsePhotoMetaMobile(row.cover_image_url|| '');
      const mZoom = document.createElement('input'); mZoom.type='hidden'; mZoom.value = String(Math.max(1, initCoverM.zm || 1).toFixed(2)); mZoom.setAttribute('data-m-zoom','1');
      const mOy = document.createElement('input'); mOy.type='hidden'; mOy.value = String(initCoverM.oym || 0); mOy.setAttribute('data-m-oy','1');
      const mOx = document.createElement('input'); mOx.type='hidden'; mOx.value = String(initCoverM.oxm || 0); mOx.setAttribute('data-m-ox','1');
      // Rectangular cover preview
      const coverPrev = document.createElement('div');
      coverPrev.style.width='100%';               // fill modal width
      coverPrev.style.maxWidth='100%';
      coverPrev.style.height='460px';             // same as site slider
      coverPrev.style.border='1px solid #e5e7eb';
      coverPrev.style.borderRadius='10px';
      coverPrev.style.overflow='hidden';
      coverPrev.style.position='relative';
      coverPrev.style.background='#fff';          // white underlay like site
      const coverBg = document.createElement('div');
      coverBg.style.position='absolute';
      coverBg.style.inset='0';
      coverBg.style.backgroundRepeat='no-repeat';
      coverBg.style.backgroundColor='#fff';
      coverPrev.appendChild(coverBg);
      coverLbl.appendChild(coverPrev);
      // Mobile preview (approx. mobile slider aspect)
      const mobileTitle = document.createElement('div'); mobileTitle.textContent = 'Mobil Önizleme'; mobileTitle.style.fontWeight='600'; coverLbl.appendChild(mobileTitle);
      const coverPrevM = document.createElement('div');
      coverPrevM.style.width = '100%';
      coverPrevM.style.maxWidth='360px'; coverPrevM.style.margin='0 auto';
      coverPrevM.style.height = '460px';
      coverPrevM.style.border = '1px solid #e5e7eb';
      coverPrevM.style.borderRadius = '10px';
      coverPrevM.style.overflow = 'hidden';
      coverPrevM.style.position = 'relative';
      coverPrevM.style.background = '#fff';
      const coverBgM = document.createElement('div');
      coverBgM.style.position = 'absolute';
      coverBgM.style.inset = '0';
      coverBgM.style.backgroundRepeat = 'no-repeat';
      coverBgM.style.backgroundColor = '#fff';
      coverPrevM.appendChild(coverBgM);
      coverLbl.appendChild(coverPrevM);

      function currentCoverBase(){
        if (iFile.files && iFile.files[0]){ try{ return URL.createObjectURL(iFile.files[0]); }catch{} }
        return (String(iIn.value|| '').trim() || stripPhotoKey(row.cover_image_url||row.image_url|| ''));
      }
      function updateCoverPreview(){
  const base = currentCoverBase();
  if (!base){
    coverPrev.style.display='none';
    if (typeof coverPrevM!=='undefined') coverPrevM.style.display='none';
    return;
  }
  coverPrev.style.display='block';
  if (typeof coverPrevM!=='undefined') coverPrevM.style.display='block';
  coverBg.style.backgroundImage = `url(${base})`;
  if (typeof coverBgM!=='undefined') coverBgM.style.backgroundImage = `url(${base})`;
  // Desktop crop
  const zD = Number(cZoom.value||1);
  const oxD = Number(cOx.value||0);
  const oyD = Number(cOy.value||0);
  const sizeD = `${Math.max(100, Math.round(100*zD))}%`;
  coverBg.style.backgroundSize = sizeD;
  const boxWD = coverPrev && coverPrev.clientWidth ? coverPrev.clientWidth : 0;
  const boxHD = coverPrev && coverPrev.clientHeight ? coverPrev.clientHeight : 0;
  if (boxWD && boxHD){
    const oxp = Math.round((oxD / boxWD) * 1000) / 10;
    const oyp = Math.round((oyD / boxHD) * 1000) / 10;
    coverBg.style.backgroundPosition = `calc(50% + ${oxp}%) calc(50% + ${oyp}%)`;
  } else {
    coverBg.style.backgroundPosition = `calc(50% + ${oxD}px) calc(50% + ${oyD}px)`;
  }
  // Mobile crop (independent)
  if (typeof coverBgM!=='undefined'){
    const zM = (typeof mZoom!=='undefined' && mZoom.value) ? Number(mZoom.value) : zD;
    const oxM = (typeof mOx!=='undefined' && mOx.value!=null) ? Number(mOx.value) : 0;
    const oyM = (typeof mOy!=='undefined' && mOy.value!=null) ? Number(mOy.value) : 0;
    const sizeM = `${Math.max(100, Math.round(100*zM))}%`;
    coverBgM.style.backgroundSize = sizeM;
    const boxWM = coverPrevM && coverPrevM.clientWidth ? coverPrevM.clientWidth : 0;
    const boxHM = coverPrevM && coverPrevM.clientHeight ? coverPrevM.clientHeight : 0;
    if (boxWM && boxHM){
      const oxpm = Math.round((oxM / boxWM) * 1000) / 10;
      const oypm = Math.round((oyM / boxHM) * 1000) / 10;
      coverBgM.style.backgroundPosition = `calc(50% + ${oxpm}%) calc(50% + ${oypm}%)`;
    } else {
      coverBgM.style.backgroundPosition = `calc(50% + ${oxM}px) calc(50% + ${oyM}px)`;
    }
  }
}// === Slider kadraj aracı ===
(function(){
  function read(){ return { z:+cZoom.value||1, ox:+cOx.value||0, oy:+cOy.value||0 }; }
  function write(st){
    cZoom.value = Math.max(1, Math.min(3, st.z)).toFixed(2);
    cOx.value   = Math.round(st.ox);
    cOy.value   = Math.round(st.oy);
    updateCoverPreview();
  }

  const row = document.createElement('div');
  row.style.display='flex';
  row.style.gap='8px';
  row.style.marginTop='8px';

  const btnCrop   = document.createElement('button');
  btnCrop.type='button'; btnCrop.className='btn btn-outline'; btnCrop.textContent='Kadrajla (Slider)';
  const btnCenter = document.createElement('button');
  btnCenter.type='button'; btnCenter.className='btn btn-outline'; btnCenter.textContent='Ortala';
  const btnReset  = document.createElement('button');
  btnReset.type='button';  btnReset.className='btn btn-outline'; btnReset.textContent='Sıfırla';
  row.appendChild(btnCrop); row.appendChild(btnCenter); row.appendChild(btnReset);
  coverLbl.appendChild(row);
  form.appendChild(coverLbl);    

    /* — pan-zoom inline preview — */
  function attachPanZoom(el){
    let drag=false, sx=0, sy=0, ox0=0, oy0=0;
    el.style.cursor='grab';
    el.addEventListener('pointerdown',e=>{ drag=true; sx=e.clientX; sy=e.clientY; const st=read(); ox0=st.ox; oy0=st.oy; el.setPointerCapture(e.pointerId); el.style.cursor='grabbing'; });
    el.addEventListener('pointermove',e=>{ if(!drag) return; write({ z:read().z, ox:ox0+e.clientX-sx, oy:oy0+e.clientY-sy }); });
    el.addEventListener('pointerup',()=>{ drag=false; try{ el.releasePointerCapture && el.releasePointerCapture(); }catch{} el.style.cursor='grab'; });
    el.addEventListener('wheel',e=>{ e.preventDefault(); const step=e.deltaY<0?0.08:-0.08; write({ z:read().z+step, ox:read().ox, oy:read().oy }); },{passive:false});
  }
  attachPanZoom(coverPrev);
  function attachPanZoomMobile(el){
    let drag=false, sx=0, sy=0, ox0=0, oy0=0;
    el.style.cursor='grab';
    el.addEventListener('pointerdown',e=>{ drag=true; sx=e.clientX; sy=e.clientY; const ox=Number((typeof mOx!=='undefined'&&mOx.value)||0); const oy=Number((typeof mOy!=='undefined'&&mOy.value)||0); ox0=ox; oy0=oy; el.setPointerCapture(e.pointerId); el.style.cursor='grabbing'; });
    el.addEventListener('pointermove',e=>{ if(!drag) return; if(typeof mOx!=='undefined') mOx.value = String(Math.round(ox0 + e.clientX - sx)); if(typeof mOy!=='undefined') mOy.value = String(Math.round(oy0 + e.clientY - sy)); updateCoverPreview(); });
    el.addEventListener('pointerup',()=>{ drag=false; try{ el.releasePointerCapture && el.releasePointerCapture(); }catch{} el.style.cursor='grab'; });
    el.addEventListener('wheel',e=>{ e.preventDefault(); const step=e.deltaY<0?0.08:-0.08; if(typeof mZoom!=='undefined'){ const cur=Number(mZoom.value||1); mZoom.value = String(Math.max(1, Math.min(3, cur+step)).toFixed(2)); updateCoverPreview(); } },{passive:false});
  }
  if (typeof coverPrevM!=='undefined') attachPanZoomMobile(coverPrevM);
  btnCenter.onclick = ()=> write({ z:read().z, ox:0, oy:0 });
  btnReset .onclick = ()=> write({ z:1, ox:0, oy:0 });

  /* — full-screen overlay — */
  btnCrop.onclick = ()=>{
    const base=currentCoverBase(); if(!base) return;
    const ov=document.createElement('div');
    ov.style.position='fixed'; ov.style.inset='0'; ov.style.background='rgba(0,0,0,.6)';
    ov.style.display='flex'; ov.style.alignItems='center'; ov.style.justifyContent='center'; ov.style.zIndex='9999';
    const box=document.createElement('div'); box.style.width='min(1100px,96vw)'; box.style.aspectRatio='1100/460';
    box.style.position='relative'; box.style.background='#fff'; box.style.borderRadius='14px'; box.style.overflow='hidden';
    const bg=document.createElement('div'); bg.style.position='absolute'; bg.style.inset='0'; bg.style.backgroundImage=`url(${base})`;
    bg.style.backgroundRepeat='no-repeat'; box.appendChild(bg);
    const bar=document.createElement('div'); bar.style.position='absolute'; bar.style.top='10px'; bar.style.right='10px'; bar.style.display='flex'; bar.style.gap='8px';
    const ok=document.createElement('button'); ok.className='btn btn-success'; ok.textContent='Uygula';
    const cancel=document.createElement('button'); cancel.className='btn btn-danger'; cancel.textContent='İptal';
    bar.appendChild(ok); bar.appendChild(cancel); box.appendChild(bar); ov.appendChild(box); document.body.appendChild(ov);

    let st=read();
    const apply=()=>{
      bg.style.backgroundSize=`${Math.round(st.z*100)}%`;
      bg.style.backgroundPosition=`calc(50% + ${st.ox}px) calc(50% + ${st.oy}px)`;
    }; apply();

    /* overlay pan/zoom */
    let drag2=false,sx2=0,sy2=0,ox2=0,oy2=0;
    box.addEventListener('pointerdown',e=>{drag2=true;sx2=e.clientX;sy2=e.clientY;ox2=st.ox;oy2=st.oy;box.setPointerCapture(e.pointerId);});
    box.addEventListener('pointermove',e=>{if(drag2){st.ox=ox2+e.clientX-sx2; st.oy=oy2+e.clientY-sy2; apply();}});
    box.addEventListener('pointerup',()=>{drag2=false;});
    box.addEventListener('wheel',e=>{e.preventDefault(); st.z=Math.max(1,Math.min(3,st.z+(e.deltaY<0?0.08:-0.08))); apply();},{passive:false});

    ok.onclick     = ()=>{ write(st); document.body.removeChild(ov); };
    cancel.onclick = ()=>{ document.body.removeChild(ov); };
    ov.onclick     = e=>{ if(e.target===ov) document.body.removeChild(ov); };
    window.addEventListener('keydown',function esc(ev){ if(ev.key==='Escape'){document.body.removeChild(ov); window.removeEventListener('keydown',esc);} });
  };
})();
      updateCoverPreview();
      [cZoom, cOx, cOy].forEach(el => el.addEventListener('input', updateCoverPreview));
      iFile.addEventListener('change', updateCoverPreview);
      // Status
      const stLbl=document.createElement('label'); stLbl.style.display='grid'; stLbl.style.gap='6px'; stLbl.innerHTML='<span>Durum</span>'; const stSel=document.createElement('select');
      const statusOptsNews = [
        { v:'draft', t:'Taslak' },
        { v:'scheduled', t:'Planlı' },
        { v:'published', t:'Yayımlandı' },
        { v:'archived', t:'Arşivli' },
        { v:'unpublished', t:'Yayından Kaldırıldı' },
      ];
      statusOptsNews.forEach(({v,t})=>{ const o=document.createElement('option'); o.value=v; o.textContent=t; if (v===(row.status||'draft')) o.selected=true; stSel.appendChild(o); });
      stLbl.appendChild(stSel); form.appendChild(stLbl);
      // Dates
      const pLbl=document.createElement('label'); pLbl.style.display='grid'; pLbl.style.gap='6px'; pLbl.innerHTML='<span>Yayın Tarihi</span>'; const pIn=document.createElement('input'); pIn.type='datetime-local'; pIn.value=row.published_at? new Date(row.published_at).toISOString().slice(0,16): ''; pLbl.appendChild(pIn); form.appendChild(pLbl);
      const uLbl=document.createElement('label'); uLbl.style.display='grid'; uLbl.style.gap='6px'; uLbl.innerHTML='<span>Yayından Kaldırma</span>'; const uIn=document.createElement('input'); uIn.type='datetime-local'; uIn.value=row.unpublish_at? new Date(row.unpublish_at).toISOString().slice(0,16): ''; uLbl.appendChild(uIn); form.appendChild(uLbl);
      // Gallery
      const gLbl=document.createElement('label'); gLbl.style.display='grid'; gLbl.style.gap='6px'; gLbl.innerHTML='<span>Galeri URL’leri (JSON dizi, eski görseli silmek için URL’yi çıkarın)</span>';
      const gIn=document.createElement('textarea'); gIn.rows=3;
      try{ const arr = Array.isArray(row.gallery_urls)? row.gallery_urls : JSON.parse(row.gallery_urls||'[]'); gIn.value = JSON.stringify(Array.isArray(arr)?arr:[], null, 0); }
      catch{ gIn.value='[]'; }
      gLbl.appendChild(gIn); form.appendChild(gLbl);
      // Actions
      const actions=document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
      const saveBtn=document.createElement('button'); saveBtn.className='btn btn-success'; saveBtn.textContent='Kaydet';
      const cancelBtn=document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
      actions.appendChild(saveBtn); actions.appendChild(cancelBtn);
      if (row.id){
        const delBtn=document.createElement('button'); delBtn.className='btn btn-danger'; delBtn.textContent='Sil'; actions.appendChild(delBtn);
        delBtn.addEventListener('click', async (e)=>{
          e.preventDefault();
          try{
            if (!confirm('Bu kaydı silmek istediğinize emin misiniz?')) return;
            const { error } = await sb().from('news').delete().eq('id', row.id);
            if (error) throw error;
            closeModal(); if (typeof loadAdminNews==='function') await loadAdminNews();
          }catch(err){ alert('Silinemedi: ' + (err?.message||String(err))); }
        });
      }
      form.appendChild(actions);

      cancelBtn.addEventListener('click',(e)=>{ e.preventDefault(); closeModal(); });
      saveBtn.addEventListener('click', async (e)=>{
        e.preventDefault();
        try{
          const payload = {
            id: row.id || undefined,
            title: String(tIn.value|| '').trim(),
            summary: String(sIn.value|| '').trim(),
            body: String(ed.innerHTML|| '').trim(),
            image_url: String(iIn.value|| '').trim() || null,
            cover_image_url: null,
            status: String(stSel.value||'draft'),
            published_at: pIn.value ? new Date(pIn.value).toISOString() : null,
            unpublish_at: uIn.value ? new Date(uIn.value).toISOString() : null,
          };
          // Validate status/date relations
          if (payload.status === 'scheduled' && !payload.published_at){ return alert('Planlı için Yayın Tarihi gerekli'); }
          if (payload.unpublish_at && payload.published_at && new Date(payload.unpublish_at) < new Date(payload.published_at)){
            return alert('Yayından kaldırma tarihi, yayın tarihinden önce olamaz');
          }
          if (payload.status === 'published' && !payload.published_at){ payload.published_at = new Date().toISOString(); }

          // Upload image file if provided
          if (iFile.files && iFile.files[0]){
            const f = iFile.files[0];
            const ext = (f.name.split('.').pop()||'jpg').toLowerCase();
            const path = `news/${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
            try{ payload.image_url = await uploadToBucketGeneric(f, 'images', path); }
            catch(err){ return alert(String(err?.message||err)); }
          }
          // Determine cover crop meta
          const coverMeta = { z:Number(cZoom.value||1), ox:Number(cOx.value||0), oy:Number(cOy.value||0) }; const coverMetaM = { zm:Number((typeof mZoom!=='undefined'&&mZoom.value)||coverMeta.z||1), oxm:Number((typeof mOx!=='undefined'&&mOx.value)||0), oym:Number((typeof mOy!=='undefined'&&mOy.value)||0) };
          // Compute percent-based offsets relative to the preview box for responsive parity
          try{
  const boxW = coverPrev && coverPrev.clientWidth ? coverPrev.clientWidth : 0;
  const boxH = coverPrev && coverPrev.clientHeight ? coverPrev.clientHeight : 0;
  if (boxW && boxH){
    coverMeta.oxp = Math.round((coverMeta.ox / boxW) * 1000) / 10; // 0.1% precision
    coverMeta.oyp = Math.round((coverMeta.oy / boxH) * 1000) / 10;
  }
  if (typeof coverPrevM!=='undefined'){
    const boxWM = coverPrevM && coverPrevM.clientWidth ? coverPrevM.clientWidth : 0;
    const boxHM = coverPrevM && coverPrevM.clientHeight ? coverPrevM.clientHeight : 0;
    if (boxWM && boxHM){
      coverMetaM.oxpm = Math.round((coverMetaM.oxm / boxWM) * 1000) / 10;
      coverMetaM.oypm = Math.round((coverMetaM.oym / boxHM) * 1000) / 10;
    }
  }
}catch{}
          // If no new file, still save crop meta referencing existing image_url (or keep previous cover if none)
          if (!payload.image_url){
            const base = String(payload.image_url|| '').trim() || stripPhotoKey(row.cover_image_url||row.image_url|| '');
            if (base){ payload.cover_image_url = buildPhotoValue(base, Object.assign({}, coverMeta, coverMetaM)); }
          } else {
            // Also set cover from uploaded
            if (payload.image_url){ payload.cover_image_url = buildPhotoValue(payload.image_url, Object.assign({}, coverMeta, coverMetaM)); }
          }
          try{
            const parsed = JSON.parse(String(gIn.value||'[]'));
            if (Array.isArray(parsed)) payload.gallery_urls = JSON.stringify(parsed);
          }catch{ /* ignore bad JSON */ }
          const q = row.id ? sb().from('news').update(payload).eq('id', row.id) : sb().from('news').insert(payload).select('id').single();
          const { error } = await q; if (error) throw error;
          closeModal(); if (typeof loadAdminNews==='function') await loadAdminNews();
        }catch(err){ alert('Kaydedilemedi: ' + (err?.message||String(err))); }
      });

      (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
    }catch(e){ alert('Haber formu açılamadı: ' + (e?.message||String(e))); }
  }

  function openAnnModal(row){
    try{
      row = row || { id:null, title:'', body:'', image_url:'', status:'draft', published_at:null, unpublish_at:null };
      $modalTitle().textContent = row.id ? 'Duyuruyu Düzenle' : 'Yeni Duyuru';
      const form = $modalForm(); form.innerHTML= '';

      // Title
      const tLbl=document.createElement('label'); tLbl.style.display='grid'; tLbl.style.gap='6px'; tLbl.innerHTML='<span>Başlık</span>'; const tIn=document.createElement('input'); tIn.value=row.title|| ''; tLbl.appendChild(tIn); form.appendChild(tLbl);
      // Body
      const bLbl=document.createElement('label'); bLbl.style.display='grid'; bLbl.style.gap='6px'; bLbl.innerHTML='<span>İçerik</span>'; const bIn=document.createElement('textarea'); bIn.rows=6; bIn.value=row.body|| ''; bIn.style.fontFamily='ui-monospace, SFMono-Regular, Menlo, Consolas, monospace';
      const tb=document.createElement('div'); tb.style.display='flex'; tb.style.gap='6px'; tb.style.flexWrap='wrap'; tb.style.margin='6px 0';
      function mkBtn(txt, title, fn){ const btn=document.createElement('button'); btn.type='button'; btn.className='btn btn-outline'; btn.textContent=txt; btn.title=title; btn.style.padding='6px 10px'; btn.addEventListener('click', fn); return btn; }
      function selWrap(before, after=''){ const s=bIn.selectionStart||0, e=bIn.selectionEnd||0; const v=bIn.value; const picked=v.slice(s,e); const rep=before+picked+(after||before); bIn.setRangeText(rep, s, e, 'end'); bIn.focus(); renderPrev(); }
      function selLinePrefix(prefix){
  const s = bIn.selectionStart||0, e = bIn.selectionEnd||0;
  const v = bIn.value;
  const start = v.lastIndexOf('\n', s-1) + 1;
  const end = e;
  const lines = v.slice(start, end).split('\n').map(l => prefix + l);
  const rep = lines.join('\n');
  bIn.setSelectionRange(start, end);
  bIn.setRangeText(rep, start, end, 'end');
  bIn.focus();
  renderPrev();
}tb.appendChild(mkBtn('B','Kalın', ()=> selWrap('**','**')));
      tb.appendChild(mkBtn('I','İtalik', ()=> selWrap('*','*')));
      tb.appendChild(mkBtn('H2','Başlık', ()=> selLinePrefix('## ')));
      tb.appendChild(mkBtn('•','Liste', ()=> selLinePrefix('- ')));
      tb.appendChild(mkBtn('🔗','Bağlantı', ()=>{ const s=bIn.selectionStart||0,e=bIn.selectionEnd||0; const picked=bIn.value.slice(s,e)||'metin'; const url=prompt('Bağlantı URL','https://'); if(!url) return; const rep=`[${picked}](${url})`; bIn.setRangeText(rep, s, e, 'end'); bIn.focus(); renderPrev(); }));
      const prev=document.createElement('div'); prev.className='card'; prev.style.padding='10px'; prev.style.maxHeight='220px'; prev.style.overflow='auto'; prev.style.marginTop='6px';
      function mdToHtml(md){ let html='\n'+String(md|| '')+'\n'; html=html.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); html=html.replace(/^######\s+(.*)$/gm,'<h6>$1</h6>').replace(/^#####\s+(.*)$/gm,'<h5>$1</h5>').replace(/^####\s+(.*)$/gm,'<h4>$1</h4>').replace(/^###\s+(.*)$/gm,'<h3>$1</h3>').replace(/^##\s+(.*)$/gm,'<h2>$1</h2>').replace(/^#\s+(.*)$/gm,'<h1>$1</h1>'); html=html.replace(/\*\*(.+?)\*\*/g,'<strong>$1</strong>').replace(/\*(.+?)\*/g,'<em>$1</em>').replace(/`([^`]+)`/g,'<code>$1</code>'); html=html.replace(/\[([^\]]+)\]\((https?:[^\s)]+)\)/g,'<a href="$2" target="_blank" rel="noopener noreferrer">$1<\/a>'); html=html.replace(/^(?:\s*[-\*]\s+.+\n)+/gm,(block)=>{ const items=block.trim().split(/\n/).map(l=> l.replace(/^\s*[-\*]\s+/,'')).map(t=>`<li>${t}<\/li>`).join(''); return `<ul>${items}<\/ul>`; }); html=html.replace(/^(?!<h\d|<ul|<li|<p|<code|<blockquote|<img|<a)(.+)$/gm,'<p>$1<\/p>'); return html; }
      function renderPrev(){ prev.innerHTML = mdToHtml(bIn.value); }
      renderPrev(); bIn.addEventListener('input', renderPrev);
      bLbl.appendChild(tb); bLbl.appendChild(bIn); bLbl.appendChild(prev); form.appendChild(bLbl);

      // Image
      const iLbl=document.createElement('label'); iLbl.style.display='grid'; iLbl.style.gap='6px'; iLbl.innerHTML='<span>Kapak Görseli</span>';
      const iIn=document.createElement('input'); iIn.placeholder='https://...'; iIn.value=row.image_url|| '';
      const iFile=document.createElement('input'); iFile.type='file'; iFile.accept='image/*';
      const iWrap=document.createElement('div'); iWrap.style.display='grid'; iWrap.style.gap='6px'; iWrap.appendChild(iIn); iWrap.appendChild(iFile); iLbl.appendChild(iWrap); form.appendChild(iLbl);
      // Status
      const stLbl=document.createElement('label'); stLbl.style.display='grid'; stLbl.style.gap='6px'; stLbl.innerHTML='<span>Durum</span>'; const stSel=document.createElement('select');
      const statusOptsAnn = [
        { v:'draft', t:'Taslak' },
        { v:'scheduled', t:'Planlı' },
        { v:'published', t:'Yayımlandı' },
        { v:'archived', t:'Arşivli' },
        { v:'unpublished', t:'Yayından Kaldırıldı' },
      ];
      statusOptsAnn.forEach(({v,t})=>{ const o=document.createElement('option'); o.value=v; o.textContent=t; if (v===(row.status||'draft')) o.selected=true; stSel.appendChild(o); });
      stLbl.appendChild(stSel); form.appendChild(stLbl);
      // Dates
      const pLbl=document.createElement('label'); pLbl.style.display='grid'; pLbl.style.gap='6px'; pLbl.innerHTML='<span>Yayın Tarihi</span>'; const pIn=document.createElement('input'); pIn.type='datetime-local'; pIn.value=row.published_at? new Date(row.published_at).toISOString().slice(0,16): ''; pLbl.appendChild(pIn); form.appendChild(pLbl);
      const uLbl=document.createElement('label'); uLbl.style.display='grid'; uLbl.style.gap='6px'; uLbl.innerHTML='<span>Yayından Kaldırma</span>'; const uIn=document.createElement('input'); uIn.type='datetime-local'; uIn.value=row.unpublish_at? new Date(row.unpublish_at).toISOString().slice(0,16): ''; uLbl.appendChild(uIn); form.appendChild(uLbl);
      // Actions
      const actions=document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
      const saveBtn=document.createElement('button'); saveBtn.className='btn btn-success'; saveBtn.textContent='Kaydet';
      const cancelBtn=document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
      actions.appendChild(saveBtn); actions.appendChild(cancelBtn);
      if (row.id){
        const delBtn=document.createElement('button'); delBtn.className='btn btn-danger'; delBtn.textContent='Sil'; actions.appendChild(delBtn);
        delBtn.addEventListener('click', async (e)=>{
          e.preventDefault();
          try{
            if (!confirm('Bu kaydı silmek istediğinize emin misiniz?')) return;
            const { error } = await sb().from('announcements').delete().eq('id', row.id);
            if (error) throw error;
            closeModal(); if (typeof loadAdminAnnouncements==='function') await loadAdminAnnouncements();
          }catch(err){ alert('Silinemedi: ' + (err?.message||String(err))); }
        });
      }
      form.appendChild(actions);

      cancelBtn.addEventListener('click',(e)=>{ e.preventDefault(); closeModal(); });
      saveBtn.addEventListener('click', async (e)=>{
        e.preventDefault();
        try{
          const payload = {
            id: row.id || undefined,
            title: String(tIn.value|| '').trim(),
            body: String(bIn.value|| '').trim(),
            image_url: String(iIn.value|| '').trim() || null,
            status: String(stSel.value||'draft'),
            published_at: pIn.value ? new Date(pIn.value).toISOString() : null,
            unpublish_at: uIn.value ? new Date(uIn.value).toISOString() : null,
          };
          // Validate
          if (payload.status === 'scheduled' && !payload.published_at){ return alert('Planlı için Yayın Tarihi gerekli'); }
          if (payload.unpublish_at && payload.published_at && new Date(payload.unpublish_at) < new Date(payload.published_at)){
            return alert('Yayından kaldırma tarihi, yayın tarihinden önce olamaz');
          }
          if (payload.status === 'published' && !payload.published_at){ payload.published_at = new Date().toISOString(); }

          // Upload image file if provided
          if (iFile.files && iFile.files[0]){
            const f = iFile.files[0];
            const ext = (f.name.split('.').pop()||'jpg').toLowerCase();
            const path = `announcements/${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
            try{ payload.image_url = await uploadToBucketGeneric(f, 'images', path); }
            catch(err){ return alert(String(err?.message||err)); }
          }
          const q = row.id ? sb().from('announcements').update(payload).eq('id', row.id) : sb().from('announcements').insert(payload).select('id').single();
          const { error } = await q; if (error) throw error;
          closeModal(); if (typeof loadAdminAnnouncements==='function') await loadAdminAnnouncements();
        }catch(err){ alert('Kaydedilemedi: ' + (err?.message||String(err))); }
      });

      (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
    }catch(e){ alert('Duyuru formu açılamadı: ' + (e?.message||String(e))); }
  }

  // Generic uploader for various buckets (returns public URL)
  async function uploadToBucketGeneric(file, bucket, filePath){
    try{
      const cleanPath = filePath.replace(/^\/+/, '');
      const { error } = await sb()
        .storage
        .from(bucket)
        .upload(cleanPath, file, {
          cacheControl:'3600',
          upsert:true
        });
      
      if (error) throw error;
      
      // Get the public URL for this object
      const { data: { publicUrl } } = sb()
        .storage
        .from(bucket)
        .getPublicUrl(cleanPath);
        
      return publicUrl;
    }catch(e){
      console.error('Upload generic error:', e);
      throw new Error('Görsel yüklenemedi: ' + (e?.message||'Bilinmeyen hata'));
    }
  }
// ========== AFİŞLER (POSTERS) ==========
async function loadAdminPosters(){
  const tbody = document.querySelector('#postersTableBody'); if (!tbody) return; tbody.innerHTML= '';
  try{
    const table = tbody.closest('table');
    const headRow = table && table.tHead && table.tHead.rows && table.tHead.rows[0];
    if (headRow){
      headRow.innerHTML = '<th>Başlık</th><th>Durum</th><th>Yayınlanma Tarihi</th><th>İşlem</th>';
    }
  }catch{}
  try{
    const { data, error } = await sb().from('posters')
      .select('id, title, status, published_at')
      .order('published_at', { ascending:false, nullsFirst:true });
    if (error) throw error;
    (data||[]).forEach(row => {
      const tr=document.createElement('tr');
      const statusMap = { published:'Yayımlandı', draft:'Taslak', scheduled:'Planlı', archived:'Arşivli', unpublished:'Yayından Kaldırıldı' };
      const statusTr = statusMap[String(row.status|| '').toLowerCase()] || (row.status|| '');
      tr.innerHTML = `
        <td>${escapeHtml(row.title|| '')}</td>
        <td>${escapeHtml(statusTr)}</td>
        <td>${row.published_at ? new Date(row.published_at).toLocaleString('tr-TR') : '-'}</td>
        <td class="actions">
          <button class="btn btn-warning" data-edit-poster="${row.id}">Düzenle</button>
          ${String(row.status|| '').toLowerCase()==='published'
            ? `<button class="btn btn-danger" data-unpub-poster="${row.id}">Yayından Kaldır</button>` 
            : `<button class="btn btn-success" data-pub-poster="${row.id}">Yayınla</button>` 
          }
        </td>`;
      tbody.appendChild(tr);
    });
    const addBtn = document.querySelector('#newPosterBtn');
    if (addBtn && !addBtn.dataset.wired){
      addBtn.dataset.wired='1';
      addBtn.onclick = ()=> openPosterModal({ id:null, title:'', body:'', image_url:'', status:'draft', published_at:null });
    }
    wirePostersRowActions();
  }catch(e){ alert('Afişler yüklenemedi: ' + (e?.message || String(e))); }
}

function wirePostersRowActions(){
  try{
    const tbody = document.querySelector('#postersTableBody'); if (!tbody) return;
    tbody.querySelectorAll('button[data-edit-poster]').forEach(btn=>{
      if (btn.dataset.wired) return; btn.dataset.wired='1';
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-edit-poster');
        try{
          const { data } = await sb().from('posters').select('*').eq('id', id).maybeSingle();
          openPosterModal(data||{ id, title:'', body:'', image_url:'', status:'draft', published_at:null });
        }catch(err){ alert('Afiş yüklenemedi: ' + (err?.message||String(err))); }
      });
    });
    tbody.querySelectorAll('button[data-pub-poster]').forEach(btn=>{
      if (btn.dataset.wired) return; btn.dataset.wired='1';
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-pub-poster');
        try{ const { error } = await sb().from('posters').update({ status:'published', published_at: new Date().toISOString() }).eq('id', id); if (error) throw error; await loadAdminPosters(); }catch(e){ alert('Yayınlanamadı: '+(e?.message||String(e))); }
      });
    });
    tbody.querySelectorAll('button[data-unpub-poster]').forEach(btn=>{
      if (btn.dataset.wired) return; btn.dataset.wired='1';
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-unpub-poster');
        try{ const { error } = await sb().from('posters').update({ status:'unpublished' }).eq('id', id); if (error) throw error; await loadAdminPosters(); }catch(e){ alert('Yayından kaldırılamadı: '+(e?.message||String(e))); }
      });
    });
  }catch{}
}

function openPosterModal(row){
  try{
    row = row || { id:null, title:'', body:'', image_url:'', status:'draft', published_at:null };
    $modalTitle().textContent = row.id ? 'Afişi Düzenle' : 'Yeni Afiş';
    const form = $modalForm(); form.innerHTML= '';

    // Başlık
    const tLbl=document.createElement('label'); tLbl.style.display='grid'; tLbl.style.gap='6px'; tLbl.innerHTML='<span>Başlık</span>';
    const tIn=document.createElement('input'); tIn.value=row.title|| ''; tLbl.appendChild(tIn); form.appendChild(tLbl);

    // İçerik (Word benzeri editör)
    const bLbl=document.createElement('label'); bLbl.style.display='grid'; bLbl.style.gap='6px'; bLbl.innerHTML='<span>İçerik</span>';
    const tb=document.createElement('div'); tb.style.display='flex'; tb.style.flexWrap='wrap'; tb.style.gap='6px'; tb.style.margin='6px 0';
    const ed=document.createElement('div'); ed.contentEditable='true'; ed.className='card'; ed.style.minHeight='160px'; ed.style.padding='10px'; ed.style.overflow='auto'; ed.innerHTML = (row.body|| '');
    function btn(label, title, on){ const b=document.createElement('button'); b.type='button'; b.className='btn btn-outline'; b.textContent=label; b.title=title; b.style.padding='6px 10px'; b.addEventListener('click', (e)=>{ e.preventDefault(); on(); ed.focus(); }); return b; }
    function applyInlineStyle(prop, val){
      const sel = window.getSelection(); if (!sel || !sel.rangeCount) return;
      const range = sel.getRangeAt(0);
      if (range.collapsed){
        const span=document.createElement('span'); span.style[prop]=val; span.appendChild(document.createTextNode('\u200b'));
        range.insertNode(span); sel.removeAllRanges(); const r=document.createRange(); r.selectNodeContents(span); r.collapse(false); sel.addRange(r);
      } else {
        const span=document.createElement('span'); span.style[prop]=val; try{ range.surroundContents(span); }catch{ document.execCommand('foreColor'); }
      }
    }
    function applyBlock(cmd){ document.execCommand(cmd,false,null); }
    const ff=document.createElement('select'); ff.className='btn btn-outline'; ['Default','Arial','Georgia','Tahoma','Times New Roman','Verdana','Courier New'].forEach(f=>{ const o=document.createElement('option'); o.value=(f==='Default')?'':f; o.textContent=f; ff.appendChild(o); }); ff.addEventListener('change',()=>{ if(ff.value) applyInlineStyle('fontFamily', ff.value); });
    const fs=document.createElement('input'); fs.type='number'; fs.min='10'; fs.max='72'; fs.value='16'; fs.title='Yazı Boyutu (px)'; fs.className='btn btn-outline'; fs.style.width='84px'; fs.addEventListener('change',()=> applyInlineStyle('fontSize', fs.value+'px'));
    const lh=document.createElement('select'); lh.className='btn btn-outline'; ['1.0','1.2','1.4','1.6','1.8','2.0'].forEach(v=>{ const o=document.createElement('option'); o.value=v; o.textContent='Satır: '+v; lh.appendChild(o); }); lh.addEventListener('change',()=> applyInlineStyle('lineHeight', lh.value));
    const col=document.createElement('input'); col.type='color'; col.className='btn btn-outline'; col.title='Yazı Rengi'; col.addEventListener('input',()=> applyInlineStyle('color', col.value));
    tb.appendChild(btn('B','Kalın', ()=> applyBlock('bold')));
    tb.appendChild(btn('I','İtalik', ()=> applyBlock('italic')));
    tb.appendChild(btn('U','Altı Çizili', ()=> applyBlock('underline')));
    tb.appendChild(btn('Sol','Sola Hizala', ()=> applyBlock('justifyLeft')));
    tb.appendChild(btn('Ortala','Ortala', ()=> applyBlock('justifyCenter')));
    tb.appendChild(btn('Sağ','Sağa Hizala', ()=> applyBlock('justifyRight')));
    tb.appendChild(btn('İkiye','İki Yana Yasla', ()=> applyBlock('justifyFull')));
    tb.appendChild(btn('• Liste','Madde İşaretli', ()=> applyBlock('insertUnorderedList')));
    tb.appendChild(btn('1. Liste','Numaralı', ()=> applyBlock('insertOrderedList')));
    tb.appendChild(ff); tb.appendChild(fs); tb.appendChild(lh); tb.appendChild(col);
    bLbl.appendChild(tb); bLbl.appendChild(ed); form.appendChild(bLbl);

    // Görsel
    const iLbl=document.createElement('label'); iLbl.style.display='grid'; iLbl.style.gap='6px'; iLbl.innerHTML='<span>Görsel</span>';
    const iIn=document.createElement('input'); iIn.placeholder='https://...'; iIn.value=row.image_url|| '';
    const iFile=document.createElement('input'); iFile.type='file'; iFile.accept='image/*';
    const iWrap=document.createElement('div'); iWrap.style.display='grid'; iWrap.style.gap='6px'; iWrap.appendChild(iIn); iWrap.appendChild(iFile); iLbl.appendChild(iWrap); form.appendChild(iLbl);

    // Durum / Tarih
    const stLbl=document.createElement('label'); stLbl.style.display='grid'; stLbl.style.gap='6px'; stLbl.innerHTML='<span>Durum</span>';
    const stSel=document.createElement('select');
    const statusOptsPoster = [
      { v:'draft', t:'Taslak' },
      { v:'published', t:'Yayımlandı' },
      { v:'unpublished', t:'Yayından Kaldırıldı' },
    ];
    statusOptsPoster.forEach(({v,t})=>{ const o=document.createElement('option'); o.value=v; o.textContent=t; if ((row.status||'draft')===v) o.selected=true; stSel.appendChild(o); });
    stLbl.appendChild(stSel); form.appendChild(stLbl);
    const pLbl=document.createElement('label'); pLbl.style.display='grid'; pLbl.style.gap='6px'; pLbl.innerHTML='<span>Yayın Tarihi</span>';
    const pIn=document.createElement('input'); pIn.type='datetime-local'; pIn.value=row.published_at? new Date(row.published_at).toISOString().slice(0,16): ''; pLbl.appendChild(pIn); form.appendChild(pLbl);

    // Actions
    const actions=document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
    const saveBtn=document.createElement('button'); saveBtn.className='btn btn-success'; saveBtn.textContent='Kaydet';
    const cancelBtn=document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
    actions.appendChild(saveBtn); actions.appendChild(cancelBtn); form.appendChild(actions);
    cancelBtn.addEventListener('click',(e)=>{ e.preventDefault(); closeModal(); });

    saveBtn.addEventListener('click', async (e)=>{
      e.preventDefault();
      try{
        const payload = {
          id: row.id || undefined,
          title: String(tIn.value|| '').trim(),
          body: String(ed.innerHTML|| '').trim(),
          image_url: String(iIn.value|| '').trim() || null,
          status: String(stSel.value||'draft'),
          published_at: pIn.value ? new Date(pIn.value).toISOString() : null,
        };
        // Upload image file if provided
        if (iFile.files && iFile.files[0]){
          const f = iFile.files[0];
          const ext = (f.name.split('.').pop()||'jpg').toLowerCase();
          const path = `posters/${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
          try{ payload.image_url = await uploadToBucketGeneric(f, 'images', path); }
          catch(err){ return alert(String(err?.message||err)); }
        }
        const q = row.id
          ? sb().from('posters').update(payload).eq('id', row.id)
          : sb().from('posters').insert(payload).select('id').single();
        const { error } = await q; if (error) throw error;
        closeModal(); await loadAdminPosters();
      }catch(err){ alert('Kaydedilemedi: ' + (err?.message||String(err))); }
    });

    (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
  }catch(e){ alert('Afiş formu açılamadı: ' + (e?.message||String(e))); }
}
// ========== RAPORLAR (REPORTS) ==========
async function loadAdminReports(){
  const tbody = document.querySelector('#reportsTableBody'); if (!tbody) return; tbody.innerHTML= '';
  try{
    const table = tbody.closest('table');
    const headRow = table && table.tHead && table.tHead.rows && table.tHead.rows[0];
    if (headRow){
      headRow.innerHTML = '<th>Başlık</th><th>Durum</th><th>Yayınlanma Tarihi</th><th>İşlem</th>';
    }
  }catch{}
  try{
    const { data, error } = await sb().from('reports')
      .select('id, title, status, published_at')
      .order('published_at', { ascending:false, nullsFirst:true });
    if (error) throw error;
    (data||[]).forEach(row => {
      const tr=document.createElement('tr');
      const statusMap = { published:'Yayımlandı', draft:'Taslak', scheduled:'Planlı', archived:'Arşivli', unpublished:'Yayından Kaldırıldı' };
      const statusTr = statusMap[String(row.status|| '').toLowerCase()] || (row.status|| '');
      tr.innerHTML = `
        <td>${escapeHtml(row.title|| '')}</td>
        <td>${escapeHtml(statusTr)}</td>
        <td>${row.published_at ? new Date(row.published_at).toLocaleString('tr-TR') : '-'}</td>
        <td class="actions">
          <button class="btn btn-warning" data-edit-report="${row.id}">Düzenle</button>
          ${String(row.status|| '').toLowerCase()==='published'
            ? `<button class="btn btn-danger" data-unpub-report="${row.id}">Yayından Kaldır</button>` 
            : `<button class="btn btn-success" data-pub-report="${row.id}">Yayınla</button>` 
          }
        </td>`;
      tbody.appendChild(tr);
    });
    const addBtn = document.querySelector('#newReportBtn');
    if (addBtn && !addBtn.dataset.wired){
      addBtn.dataset.wired='1';
      addBtn.onclick = ()=> openReportModal({ id:null, title:', file_url:', status:'draft', published_at:null });
    }
    wireReportsRowActions();
  }catch(e){ alert('Raporlar yüklenemedi: ' + (e?.message || String(e))); }
}

function wireReportsRowActions(){
  try{
    const tbody = document.querySelector('#reportsTableBody'); if (!tbody) return;
    tbody.querySelectorAll('button[data-edit-report]').forEach(btn=>{
      if (btn.dataset.wired) return; btn.dataset.wired='1';
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-edit-report');
        try{
          const { data } = await sb().from('reports').select('*').eq('id', id).maybeSingle();
          openReportModal(data||{ id, title:', file_url:', status:'draft', published_at:null });
        }catch(err){ alert('Rapor yüklenemedi: ' + (err?.message||String(err))); }
      });
    });
    tbody.querySelectorAll('button[data-pub-report]').forEach(btn=>{
      if (btn.dataset.wired) return; btn.dataset.wired='1';
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-pub-report');
        try{ const { error } = await sb().from('reports').update({ status:'published', published_at: new Date().toISOString() }).eq('id', id); if (error) throw error; await loadAdminReports(); }catch(e){ alert('Yayınlanamadı: '+(e?.message||String(e))); }
      });
    });
    tbody.querySelectorAll('button[data-unpub-report]').forEach(btn=>{
      if (btn.dataset.wired) return; btn.dataset.wired='1';
      btn.addEventListener('click', async ()=>{
        const id = btn.getAttribute('data-unpub-report');
        try{ const { error } = await sb().from('reports').update({ status:'unpublished' }).eq('id', id); if (error) throw error; await loadAdminReports(); }catch(e){ alert('Yayından kaldırılamadı: '+(e?.message||String(e))); }
      });
    });
  }catch{}
}

function openReportModal(row){
  try{
    row = row || { id:null, title:'', file_url:'', status:'draft', published_at:null };
    $modalTitle().textContent = row.id ? 'Raporu Düzenle' : 'Yeni Rapor';
    const form = $modalForm(); form.innerHTML= '';

    // Başlık
    const tLbl=document.createElement('label'); tLbl.style.display='grid'; tLbl.style.gap='6px'; tLbl.innerHTML='<span>Başlık</span>';
    const tIn=document.createElement('input'); tIn.value=row.title|| ''; tLbl.appendChild(tIn); form.appendChild(tLbl);

    // Dosya
    const fLbl=document.createElement('label'); fLbl.style.display='grid'; fLbl.style.gap='6px'; fLbl.innerHTML='<span>Dosya (PDF/Word)</span>';
    const fIn=document.createElement('input'); fIn.placeholder='https://...'; fIn.value=row.file_url|| '';
    const fFile=document.createElement('input'); fFile.type='file'; fFile.accept='.pdf,.doc,.docx,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    const fWrap=document.createElement('div'); fWrap.style.display='grid'; fWrap.style.gap='6px'; fWrap.appendChild(fIn); fWrap.appendChild(fFile); fLbl.appendChild(fWrap); form.appendChild(fLbl);

    // Durum / Tarih
    const stLbl=document.createElement('label'); stLbl.style.display='grid'; stLbl.style.gap='6px'; stLbl.innerHTML='<span>Durum</span>';
    const stSel=document.createElement('select');
    const statusOptsReport = [
      { v:'draft', t:'Taslak' },
      { v:'published', t:'Yayımlandı' },
      { v:'unpublished', t:'Yayından Kaldırıldı' },
    ];
    statusOptsReport.forEach(({v,t})=>{ const o=document.createElement('option'); o.value=v; o.textContent=t; if ((row.status||'draft')===v) o.selected=true; stSel.appendChild(o); });
    stLbl.appendChild(stSel); form.appendChild(stLbl);
    const pLbl=document.createElement('label'); pLbl.style.display='grid'; pLbl.style.gap='6px'; pLbl.innerHTML='<span>Yayın Tarihi</span>';
    const pIn=document.createElement('input'); pIn.type='datetime-local'; pIn.value=row.published_at? new Date(row.published_at).toISOString().slice(0,16): ''; pLbl.appendChild(pIn); form.appendChild(pLbl);

    // Actions
    const actions = document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
    const saveBtn = document.createElement('button');
    saveBtn.type = 'button';
    saveBtn.className = 'btn btn-success';
    saveBtn.textContent = 'Kaydet';
    actions.appendChild(saveBtn);
    form.appendChild(actions);
    const cancelBtn = document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
    actions.appendChild(cancelBtn);
    cancelBtn.addEventListener('click', (e)=>{ e.preventDefault(); closeModal(); });

    saveBtn.addEventListener('click', async (e)=>{
      e.preventDefault();
      const title = String(tIn.value|| '').trim(); if (!title) return alert('Başlık gerekli');
      let file_url = String(fIn.value|| '').trim();
      const status = stSel.value || 'draft';
      const published_at = pIn.value ? new Date(pIn.value).toISOString() : null;
      try {
        // Dosya seçilmişse yükle
        const f = fFile.files && fFile.files[0];
        if (f){
          const ext = (f.name.split('.').pop()||'pdf').toLowerCase();
          const key = `reports/${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
          try{
            file_url = await uploadToBucketGeneric(f, 'docs', key);
          }catch(err){
            return alert('Dosya yüklenemedi: ' + (err?.message || String(err)));
          }
        }

        // UPDATE vs INSERT
        const q = row.id
          ? sb().from('reports').update({ title, file_url, status, published_at }).eq('id', row.id)
          : sb().from('reports').insert({ title, file_url, status, published_at }).select('id').single();
        const { error } = await q;
        if (error) throw error;

        closeModal();
        await loadAdminReports();
      } catch (err) {
        alert('Kaydedilemedi: ' + (err?.message || String(err)));
      }
    });

    (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
  }catch(e){ alert('Rapor formu açılamadı: ' + (e?.message||String(e))); }
}
  // Cache-bust helper (used by multiple modules)
  const bust = (u)=>{ try{ const x=new URL(u, location.origin); x.searchParams.set('v', Date.now()); return x.toString(); }catch{ return u; } };

  // ======= BOOT HELPERS (reintroduced minimal versions) =======
  function setupTabs(){
    try{
      qsa('.tabs button').forEach(btn => {
        if (btn.dataset.tabWired) return; btn.dataset.tabWired='1';
        btn.addEventListener('click', () => {
          qsa('.tabs button').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          const tab = btn.dataset.tab || 'news';
          enforceTabsVisibility(tab);
          refreshTab();
        });
      });
    }catch{}
  }

  function enforceTabsVisibility(active){
    try{
      const panels = qsa('.tab-panel');
      panels.forEach(el => el.classList.remove('active'));
      const target = qs(`#tab-${active}`);
      if (target){ target.classList.add('active'); target.style.display= ''; }
      // set tab button active class
      qsa('.tabs button').forEach(b => b.classList.remove('active'));
      const btn = qs(`.tabs button[data-tab="${active}"]`);
      if (btn) btn.classList.add('active');
      currentTab = active;
    }catch{}
  }

  async function refreshTab(){
    try{
      if (currentTab === 'news' && typeof loadAdminNews === 'function') await loadAdminNews();
      else if (currentTab === 'ann' && typeof loadAdminAnnouncements === 'function') await loadAdminAnnouncements();
      else if (currentTab === 'msgs' && typeof loadAdminMessages === 'function') await loadAdminMessages();
      else if (currentTab === 'pages' && typeof loadAdminPages === 'function') await loadAdminPages();
      else if (currentTab === 'posters' && typeof loadAdminPosters === 'function') await loadAdminPosters();
      else if (currentTab === 'reports' && typeof loadAdminReports === 'function') await loadAdminReports();
      else if (currentTab === 'docs' && typeof loadAdminDocs === 'function') await loadAdminDocs();
      else if (currentTab === 'benefits' && typeof loadAdminBenefits === 'function') await loadAdminBenefits();
      else if (currentTab === 'founders' && typeof loadAdminFounders === 'function') await loadAdminFounders();
      else if (currentTab === 'chairman' && typeof loadAdminChairman === 'function') await loadAdminChairman();
      else if (currentTab === 'members' && typeof loadAdminMembers === 'function') await loadAdminMembers();
      else if (currentTab === 'users' && typeof loadAdminUsers === 'function') await loadAdminUsers();
      else if (currentTab === 'settings' && typeof loadSettingsForm === 'function') await loadSettingsForm();
      else if (typeof loadAdminMembers === 'function') await loadAdminMembers(); // fallback
    }catch{}
  }

  async function loadAdminPermissions(email){
    try{
      if (!email) return { roles:[], allowed_tabs:[] };
      const { data } = await sb().from('admin_users').select('email, roles, allowed_tabs').eq('email', email).maybeSingle();
      return { email: data?.email||email, roles: (data?.roles)||[], allowed_tabs: (data?.allowed_tabs)||[] };
    }catch(e){ return { email, roles:[], allowed_tabs:[] }; }
  }

  function applyTabPermissions(){
    try{
      const defaultTabs = ['news','ann','msgs','pages','posters','reports','docs','benefits','founders','chairman', 'members','users','settings'];
      const allowed = new Set((currentAdmin.allowed_tabs||[]).length ? currentAdmin.allowed_tabs : defaultTabs);
      defaultTabs.forEach(tab => {
        const btn = qs(`.tabs button[data-tab="${tab}"]`);
        const panel = qs(`#tab-${tab}`);
        const ok = allowed.has(tab) || (currentAdmin.roles||[]).includes('superadmin');
        if (btn) btn.style.display = ok ? '' : 'none';
        if (panel){
          if (!ok){ panel.classList.remove('active'); panel.style.display = 'none'; }
          else { panel.style.display = ''; }
        }
      });
      if (!allowed.has(currentTab) && !(currentAdmin.roles||[]).includes('superadmin')){
        const first = (currentAdmin.allowed_tabs||[]).find(t => allowed.has(t)) || defaultTabs.find(t => allowed.has(t)) || 'news';
        enforceTabsVisibility(first);
      }
    }catch{}
  }

  async function checkSession(){
    try{
      const { data } = await sb().auth.getSession();
      if (data?.session?.user){
        $loginSection().hidden = true;
        $adminSection().hidden = false;
        $authStatus().textContent = data.session.user.email || 'Oturum açık';
        currentAdmin = await loadAdminPermissions(data.session.user.email);
        try{
          currentAdmin = currentAdmin || {};
          const at = Array.isArray(currentAdmin.allowed_tabs) ? currentAdmin.allowed_tabs : [];
          if (!at.includes('benefits')) at.push('benefits');
          currentAdmin.allowed_tabs = at;
        }catch{}
        applyTabPermissions();
        // ensure initial tab
        try{
          const defaultTabs = ['news','ann','msgs','pages','posters','reports','docs','benefits','founders','chairman', 'members','users','settings'];
          const isSuper = (currentAdmin.roles||[]).includes('superadmin');
          const allowed = new Set((currentAdmin.allowed_tabs||[]).length && !isSuper ? currentAdmin.allowed_tabs : defaultTabs);
          const order = ((currentAdmin.allowed_tabs||[]).length && !isSuper) ? currentAdmin.allowed_tabs : defaultTabs;
          const firstTab = order.find(t => allowed.has(t)) || 'news';
          enforceTabsVisibility(firstTab);
        }catch{ enforceTabsVisibility('news'); }
        await refreshTab();
      } else {
        showLogin();
      }
    }catch{ showLogin(); }
  }

  // ========== EVRAK MODÜLÜ (DOCS) ==========
  async function loadAdminDocs(){
    try{
      // Sub-tab wiring
      const incBtn = document.getElementById('incomingSubTabBtn');
      const outBtn = document.getElementById('outgoingSubTabBtn');
      const newBtn = document.getElementById('newIncomingDocBtn');
      const newOutBtn = document.getElementById('newOutgoingDocBtn');
      if (incBtn && !incBtn.dataset.wired){
        incBtn.dataset.wired='1';
        incBtn.addEventListener('click', (e)=>{
          e.preventDefault();
          currentDocsSubTab = 'incoming';
          try{ incBtn.className = 'btn btn-success'; }catch{}
          try{ if (outBtn) outBtn.className = 'btn btn-outline'; }catch{}
          if (newBtn) newBtn.style.display = '';
          if (newOutBtn) newOutBtn.style.display = 'none';
          loadAdminIncomingDocs();
        });
      }
      if (outBtn && !outBtn.dataset.wired){
        outBtn.dataset.wired='1';
        outBtn.addEventListener('click', (e)=>{
          e.preventDefault();
          currentDocsSubTab = 'outgoing';
          try{ outBtn.className = 'btn btn-success'; }catch{}
          try{ if (incBtn) incBtn.className = 'btn btn-outline'; }catch{}
          if (newBtn) newBtn.style.display = 'none';
          if (newOutBtn) newOutBtn.style.display = '';
          loadAdminOutgoingDocs();
        });
      }
      if (newBtn && !newBtn.dataset.wired){
        newBtn.dataset.wired='1';
        newBtn.addEventListener('click', async (e)=>{ e.preventDefault(); await openIncomingDocModal(null); });
      }
      if (newOutBtn && !newOutBtn.dataset.wired){
        newOutBtn.dataset.wired='1';
        newOutBtn.addEventListener('click', async (e)=>{ e.preventDefault(); await openOutgoingDocModal(null); });
      }

      // Default load and style
      if (currentDocsSubTab !== 'outgoing'){
        if (incBtn) incBtn.className = 'btn btn-success';
        if (outBtn) outBtn.className = 'btn btn-outline';
        if (newBtn) newBtn.style.display = '';
        if (newOutBtn) newOutBtn.style.display = 'none';
        await loadAdminIncomingDocs();
      } else {
        if (incBtn) incBtn.className = 'btn btn-outline';
        if (outBtn) outBtn.className = 'btn btn-success';
        if (newBtn) newBtn.style.display = 'none';
        if (newOutBtn) newOutBtn.style.display = '';
        await loadAdminOutgoingDocs();
      }
    }catch(e){ alert('Evrak Modülü yüklenemedi: ' + (e?.message||String(e))); }
  }

  async function loadAdminIncomingDocs(){
    const tbody = document.getElementById('incomingDocsTableBody'); if (!tbody) return; tbody.innerHTML = '';
    try{
      const table = tbody.closest('table');
      const headRow = table && table.tHead && table.tHead.rows && table.tHead.rows[0];
      if (headRow){
        headRow.innerHTML = '<th>Kayıt Sıra No</th><th>Geldiği Kişi/Kurum</th><th>Tarihi</th><th>Sayısı</th><th>Eki</th><th>Konusu</th><th>Muhafaza Dosya No</th><th>Dosya</th>';
      }
    }catch{}
    try{
      const { data, error } = await sb().from('incoming_docs')
        .select('id, record_no, from_org, date, number, attachment, subject, file_no, file_url, created_at')
        .order('date', { ascending:false, nullsFirst:true })
        .order('created_at', { ascending:false, nullsFirst:true });
      if (error) throw error;
      incomingDocsById.clear();
      (data||[]).forEach(row => {
        incomingDocsById.set(String(row.id), row);
        const tr = document.createElement('tr');
        const fileLink = row.file_url
          ? `<a href="#" data-preview-url="${escapeHtml(row.file_url)}" aria-label="Önizleme" title="Önizleme" style="display:inline-flex;align-items:center;gap:4px;">
               <svg xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                 <path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7-11-7-11-7Z"></path>
                 <circle cx="12" cy="12" r="3"></circle>
               </svg>
             </a>`
          : '-';
        tr.innerHTML = `
          <td>${escapeHtml(row.record_no||'')}</td>
          <td>${escapeHtml(row.from_org||'')}</td>
          <td>${row.date ? new Date(row.date).toLocaleDateString('tr-TR') : '-'}</td>
          <td>${escapeHtml(row.number||'')}</td>
          <td>${escapeHtml(row.attachment||'')}</td>
          <td>${escapeHtml(row.subject||'')}</td>
          <td>${escapeHtml(row.file_no||'')}</td>
          <td>${fileLink}</td>`;
        tbody.appendChild(tr);
      });
      // Wire preview clicks
      try{
        tbody.querySelectorAll('a[data-preview-url]').forEach(a => {
          if (a.dataset.wired) return; a.dataset.wired='1';
          a.addEventListener('click', (e)=>{ e.preventDefault(); const u=a.getAttribute('data-preview-url'); if (u) openDocPreviewModal(u); });
        });
      }catch{}
    }catch(e){ alert('Gelen evraklar yüklenemedi: ' + (e?.message||String(e))); }
  }

  async function loadAdminOutgoingDocs(){
    const tbody = document.getElementById('incomingDocsTableBody'); if (!tbody) return; tbody.innerHTML = '';
    try{
      const table = tbody.closest('table');
      const headRow = table && table.tHead && table.tHead.rows && table.tHead.rows[0];
      if (headRow){
        headRow.innerHTML = '<th>Kayıt Sıra No</th><th>Gönderildiği Kişi/Kurum</th><th>Tarihi</th><th>Eki</th><th>Konusu</th><th>Muhafaza Dosya No</th><th>Dosya</th>';
      }
    }catch{}
    try{
      const { data, error } = await sb().from('outgoing_docs')
        .select('id, record_no, to_org, date, attachment, subject, file_no, file_url, created_at')
        .order('date', { ascending:false, nullsFirst:true })
        .order('created_at', { ascending:false, nullsFirst:true });
      if (error) throw error;
      outgoingDocsById.clear();
      (data||[]).forEach(row => {
        outgoingDocsById.set(String(row.id), row);
        const tr = document.createElement('tr');
        const fileLink = row.file_url
          ? `<a href="#" data-preview-url="${escapeHtml(row.file_url)}" aria-label="Önizleme" title="Önizleme" style="display:inline-flex;align-items:center;gap:4px;">
               <svg xmlns="http://www.w3.org/2000/svg" width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                 <path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7-11-7-11-7Z"></path>
                 <circle cx="12" cy="12" r="3"></circle>
               </svg>
             </a>`
          : '-';
        tr.innerHTML = `
          <td>${escapeHtml(row.record_no||'')}</td>
          <td>${escapeHtml(row.to_org||'')}</td>
          <td>${row.date ? new Date(row.date).toLocaleDateString('tr-TR') : '-'}</td>
          <td>${escapeHtml(row.attachment||'')}</td>
          <td>${escapeHtml(row.subject||'')}</td>
          <td>${escapeHtml(row.file_no||'')}</td>
          <td>${fileLink}</td>`;
        tbody.appendChild(tr);
      });
      try{
        tbody.querySelectorAll('a[data-preview-url]').forEach(a => {
          if (a.dataset.wired) return; a.dataset.wired='1';
          a.addEventListener('click', (e)=>{ e.preventDefault(); const u=a.getAttribute('data-preview-url'); if (u) openDocPreviewModal(u); });
        });
      }catch{}
    }catch(e){ alert('Giden evraklar yüklenemedi: ' + (e?.message||String(e))); }
  }

  function wireIncomingDocsRowActions(){
    const tbody = document.getElementById('incomingDocsTableBody'); if (!tbody) return;
    try{
      tbody.querySelectorAll('button[data-edit-incoming]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-edit-incoming');
          const row = incomingDocsById.get(String(id));
          await openIncomingDocModal(row||null);
        });
      });
      tbody.querySelectorAll('button[data-del-incoming]').forEach(btn=>{
        if (btn.dataset.wired) return; btn.dataset.wired='1';
        btn.addEventListener('click', async ()=>{
          const id = btn.getAttribute('data-del-incoming');
          if (!id) return;
          if (!confirm('Bu evrağı silmek istediğinize emin misiniz?')) return;
          try{ await sb().from('incoming_docs').delete().eq('id', id); }
          catch(e){ return alert('Silinemedi: ' + (e?.message||String(e))); }
          try{ await loadAdminIncomingDocs(); }catch{}
        });
      });
    }catch{}
  }

  async function openDocPreviewModal(url){
    try{
      $modalTitle().textContent = 'Dosya Önizleme';
      const form = $modalForm(); form.innerHTML = '';
      const wrap = document.createElement('div'); wrap.style.display='grid'; wrap.style.gap='8px';
      const loading = document.createElement('div'); loading.textContent='Yükleniyor...'; loading.className='muted'; wrap.appendChild(loading);
      form.appendChild(wrap);
      (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));

      const ext = (url.split('?')[0].split('#')[0].split('.').pop()||'').toLowerCase();
      const imgExt = new Set(['jpg','jpeg','png','gif','bmp','webp','svg']);
      const officeExt = new Set(['doc','docx','docm','dot','dotx','xls','xlsx','xlsm','xlsb','ppt','pptx','pptm','pps','ppsx','odt','rtf']);
      let addedActions = false;

      const buildOfficeFallback = () => {
        try{ wrap.removeChild(loading); }catch{}
        const info = document.createElement('div'); info.className='muted'; info.textContent='Bu dosya türü modalde önizlenmiyor. Aşağıdan yeni sekmede açabilir veya indirebilirsiniz.'; wrap.appendChild(info);
        const actions = document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
        const viewerUrl = 'https://view.officeapps.live.com/op/view.aspx?src=' + encodeURIComponent(url);
        const openNew = document.createElement('a'); openNew.href=viewerUrl; openNew.target='_blank'; openNew.rel='noopener'; openNew.textContent='Yeni sekmede aç'; openNew.className='btn btn-outline'; actions.appendChild(openNew);
        const dlBtn = document.createElement('button'); dlBtn.type='button'; dlBtn.className='btn btn-success'; dlBtn.textContent='İndir'; actions.appendChild(dlBtn);
        const filename = () => {
          try{ const p = new URL(url).pathname; const seg = p.split('/').filter(Boolean).pop()||'dosya'; return decodeURIComponent(seg); }catch{ return 'dosya'; }
        };
        dlBtn.addEventListener('click', async ()=>{
          try{
            dlBtn.disabled = true; dlBtn.textContent='İndiriliyor...';
            const resp = await fetch(url, { credentials:'omit' }); if (!resp.ok) throw new Error('Dosya indirilemedi');
            const blob = await resp.blob(); const oUrl = URL.createObjectURL(blob);
            const a = document.createElement('a'); a.href=oUrl; a.download=filename(); document.body.appendChild(a); a.click(); a.remove();
            setTimeout(()=>{ try{ URL.revokeObjectURL(oUrl); }catch{} }, 60_000);
            dlBtn.textContent='İndir'; dlBtn.disabled=false;
          }catch(e){ alert('İndirme başarısız: ' + (e?.message||String(e))); dlBtn.textContent='İndir'; dlBtn.disabled=false; }
        });
        wrap.appendChild(actions);
        addedActions = true;
      };

      // Direct load for images (no need to blob)
      if (imgExt.has(ext)){
        wrap.removeChild(loading);
        const img = document.createElement('img'); img.src = url; img.alt = 'Önizleme'; img.style.maxWidth='100%'; img.style.maxHeight='75vh'; img.style.objectFit='contain'; wrap.appendChild(img);
      } else if (officeExt.has(ext)){
        // Standardize office files: only open in new tab or download
        buildOfficeFallback();
      } else {
        let blobUrl = null;
        try{
          const resp = await fetch(url, { credentials: 'omit' });
          if (!resp.ok) throw new Error('Önizleme için dosya alınamadı');
          const ct = String(resp.headers.get('content-type')||'').toLowerCase();
          const blob = await resp.blob();
          blobUrl = URL.createObjectURL(blob);
          wrap.removeChild(loading);
          if (ct.includes('pdf') || ext === 'pdf'){
            const embed = document.createElement('embed');
            embed.type='application/pdf';
            embed.src = blobUrl;
            embed.style.width='100%'; embed.style.height='75vh';
            wrap.appendChild(embed);
          } else if (ct.includes('word') || ct.includes('officedocument') || ct.includes('msword') || ct.includes('powerpoint') || ct.includes('excel') || ct.includes('rtf') || ct.includes('opendocument')){
            // Detected office content type after fetch -> show standardized actions, no inline preview
            try{ URL.revokeObjectURL(blobUrl); }catch{} blobUrl=null;
            buildOfficeFallback();
          } else if (ct.startsWith('image/')){
            const img = document.createElement('img'); img.src = blobUrl; img.alt = 'Önizleme'; img.style.maxWidth='100%'; img.style.maxHeight='75vh'; img.style.objectFit='contain'; wrap.appendChild(img);
          } else if (ct.startsWith('text/') || ct.includes('html')){
            const frame = document.createElement('iframe'); frame.src = blobUrl; frame.style.width='100%'; frame.style.height='75vh'; frame.setAttribute('title','Önizleme'); frame.setAttribute('loading','eager'); wrap.appendChild(frame);
          } else {
            try{ URL.revokeObjectURL(blobUrl); }catch{} blobUrl=null;
            buildOfficeFallback();
          }
          // Revoke later to free memory
          setTimeout(()=>{ try{ blobUrl && URL.revokeObjectURL(blobUrl); }catch{} }, 60_000);
        }catch(err){
          try{ wrap.removeChild(loading); }catch{}
          const info = document.createElement('div'); info.className='muted'; info.textContent='Önizleme yerleşik pencerede açılamadı.'; wrap.appendChild(info);
        }
      }

      // Fallback link
      if (!addedActions){
        const openNew = document.createElement('a'); openNew.href=url; openNew.target='_blank'; openNew.rel='noopener'; openNew.textContent='Yeni sekmede aç'; openNew.className='btn btn-outline';
        wrap.appendChild(openNew);
      }
    }catch(e){ alert('Önizleme açılamadı: ' + (e?.message||String(e))); }
  }

  async function generateNextIncomingRecordNo(){
    try{
      const year = new Date().getFullYear();
      const prefix = `${year}/`;
      // Fetch last few for this year and compute max suffix
      const { data } = await sb().from('incoming_docs')
        .select('record_no')
        .ilike('record_no', `${prefix}%`)
        .order('record_no', { ascending:false })
        .limit(1000);
      let maxSeq = 0;
      (data||[]).forEach(r => {
        const s = String(r.record_no||'');
        if (s.startsWith(prefix)){
          const tail = s.slice(prefix.length).replace(/[^0-9]/g,'');
          const n = parseInt(tail||'0',10);
          if (!isNaN(n) && n > maxSeq) maxSeq = n;
        }
      });
      const next = maxSeq + 1;
      const pad = String(next).padStart(3,'0');
      return `${prefix}${pad}`;
    }catch{ return `${new Date().getFullYear()}/001`; }
  }

  async function openIncomingDocModal(row){
    try{
      row = row || null;
      const isEdit = !!(row && row.id);
      $modalTitle().textContent = isEdit ? 'Gelen Evrak Düzenle' : 'Yeni Gelen Evrak';
      const form = $modalForm(); form.innerHTML='';

      // Fields
      const recLbl = document.createElement('label'); recLbl.style.display='grid'; recLbl.style.gap='6px'; recLbl.innerHTML = '<span>Kayıt Sıra No</span>';
      const recIn = document.createElement('input'); recIn.readOnly = true; recIn.disabled = true; recIn.value = row?.record_no || '';
      recLbl.appendChild(recIn); form.appendChild(recLbl);

      const fromLbl = document.createElement('label'); fromLbl.style.display='grid'; fromLbl.style.gap='6px'; fromLbl.innerHTML = '<span>Geldiği Kişi veya Kurum</span>';
      const fromIn = document.createElement('input'); fromIn.value = row?.from_org || '';
      fromLbl.appendChild(fromIn); form.appendChild(fromLbl);

      const dateLbl = document.createElement('label'); dateLbl.style.display='grid'; dateLbl.style.gap='6px'; dateLbl.innerHTML = '<span>Tarihi</span>';
      const dateIn = document.createElement('input'); dateIn.type='date'; dateIn.value = row?.date ? String(row.date).slice(0,10) : '';
      dateLbl.appendChild(dateIn); form.appendChild(dateLbl);

      const numLbl = document.createElement('label'); numLbl.style.display='grid'; numLbl.style.gap='6px'; numLbl.innerHTML = '<span>Sayısı</span>';
      const numIn = document.createElement('input'); numIn.value = row?.number || '';
      numLbl.appendChild(numIn); form.appendChild(numLbl);

      const attLbl = document.createElement('label'); attLbl.style.display='grid'; attLbl.style.gap='6px'; attLbl.innerHTML = '<span>Eki</span>';
      const attIn = document.createElement('input'); attIn.value = row?.attachment || '';
      attLbl.appendChild(attIn); form.appendChild(attLbl);

      const subLbl = document.createElement('label'); subLbl.style.display='grid'; subLbl.style.gap='6px'; subLbl.innerHTML = '<span>Konusu</span>';
      const subIn = document.createElement('input'); subIn.value = row?.subject || '';
      subLbl.appendChild(subIn); form.appendChild(subLbl);

      const fileNoLbl = document.createElement('label'); fileNoLbl.style.display='grid'; fileNoLbl.style.gap='6px'; fileNoLbl.innerHTML = '<span>Muhafaz Edildiği Dosya No</span>';
      const fileNoIn = document.createElement('input'); fileNoIn.value = row?.file_no || '';
      fileNoLbl.appendChild(fileNoIn); form.appendChild(fileNoLbl);

      const fileLbl = document.createElement('label'); fileLbl.style.display='grid'; fileLbl.style.gap='6px'; fileLbl.innerHTML = '<span>Dosya Ekle</span>';
      const fileUrlIn = document.createElement('input'); fileUrlIn.placeholder='https://...'; fileUrlIn.value = row?.file_url || '';
      const fileFile = document.createElement('input'); fileFile.type='file'; fileFile.accept='application/pdf,image/*,.doc,.docx';
      const fileWrap = document.createElement('div'); fileWrap.style.display='grid'; fileWrap.style.gap='6px'; fileWrap.appendChild(fileUrlIn); fileWrap.appendChild(fileFile);
      fileLbl.appendChild(fileWrap); form.appendChild(fileLbl);

      const actions = document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
      const saveBtn = document.createElement('button'); saveBtn.className='btn btn-success'; saveBtn.textContent='Kaydet';
      const cancelBtn = document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
      actions.appendChild(saveBtn); actions.appendChild(cancelBtn); form.appendChild(actions);

      cancelBtn.addEventListener('click', (e)=>{ e.preventDefault(); closeModal(); });

      // Initialize record number for new entries
      if (!isEdit){
        try{ recIn.value = await generateNextIncomingRecordNo(); }catch{ recIn.value = `${new Date().getFullYear()}/001`; }
      }

      saveBtn.addEventListener('click', async (e)=>{
        e.preventDefault();
        try{
          const record_no = String(recIn.value||'').trim();
          if (!record_no) return alert('Kayıt Sıra No üretilemedi');
          const from_org = String(fromIn.value||'').trim();
          const date = dateIn.value ? new Date(dateIn.value).toISOString().slice(0,10) : null;
          const number = String(numIn.value||'').trim();
          const attachment = String(attIn.value||'').trim();
          const subject = String(subIn.value||'').trim();
          const file_no = String(fileNoIn.value||'').trim();
          let file_url = String(fileUrlIn.value||'').trim();

          // Upload file if chosen
          const f = fileFile.files && fileFile.files[0];
          if (f){
            const ext = (f.name.split('.').pop()||'pdf').toLowerCase();
            const safeKey = record_no.replace(/\//g,'_');
            const key = `incoming/${new Date().getFullYear()}/${safeKey}_${Date.now()}.${ext}`;
            try{ file_url = await uploadToBucketGeneric(f, 'docs', key); }
            catch(err){ return alert('Dosya yüklenemedi: ' + (err?.message||String(err))); }
          }

          const payload = { record_no, from_org, date, number, attachment, subject, file_no, file_url };
          let q;
          if (isEdit){ q = sb().from('incoming_docs').update(payload).eq('id', row.id); }
          else { q = sb().from('incoming_docs').insert(payload).select('id, record_no').single(); }
          const { data, error } = await q; if (error) throw error;
          closeModal(); await loadAdminIncomingDocs();
        }catch(err){ alert('Kaydedilemedi: ' + (err?.message||String(err))); }
      });

      (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
    }catch(e){ alert('Evrak formu açılamadı: ' + (e?.message||String(e))); }
  }

  // ============ Login reCAPTCHA (Admin) ============
  async function initLoginRecaptcha(){
    try{
      const container = document.getElementById('recaptcha-login-container');
      const btn = document.getElementById('loginSubmitBtn');
      const fbWrap = document.getElementById('captcha-login-fallback');
      const fbCb = document.getElementById('captchaLoginAgree');
      if (!container || !btn) return; // no login form

      const setEnabled = (v)=>{ try{ btn.disabled = !v; btn.setAttribute('aria-disabled', (!v).toString()); }catch{} };
      // default: require verification
      setEnabled(false);

      // Fetch site key from settings
      let siteKey = '';
      try{ siteKey = await getSettingValue('recaptcha_site_key'); }catch{}

      // Expose a small state
      try{ window.__loginRecaptcha = window.__loginRecaptcha || { token:null, widgetId:null, siteKey:null }; }catch{}
      try{ window.__loginRecaptcha.siteKey = siteKey || null; window.__loginRecaptcha.token = null; }catch{}

      if (!siteKey){
        // No site key: present fallback checkbox
        if (fbWrap) fbWrap.hidden = false;
        if (fbCb){ const onChange = ()=> setEnabled(!!fbCb.checked); fbCb.addEventListener('change', onChange); onChange(); }
        return;
      }

      // We have siteKey: ensure fallback hidden
      if (fbWrap) fbWrap.hidden = true;

      // Load API once
      const RENDER_CB = 'onAdminRecaptchaReady';
      const scriptId = 'recaptcha-api';
      function render(){
        try{
          if (!window.grecaptcha) return;
          if (fbWrap) fbWrap.hidden = true;
          const widgetId = window.grecaptcha.render('recaptcha-login-container', {
            sitekey: siteKey,
            callback: (token)=>{ try{ window.__loginRecaptcha.token = token; }catch{} setEnabled(!!token); },
            'expired-callback': ()=>{ try{ window.__loginRecaptcha.token = null; }catch{} setEnabled(false); }
          });
          try{ window.__loginRecaptcha.widgetId = widgetId; }catch{}
        }catch(e){ /* no-op */ }
      }
      if (!document.getElementById(scriptId)){
        const s = document.createElement('script'); s.id = scriptId; s.async = true; s.defer = true;
        s.src = `https://www.google.com/recaptcha/api.js?onload=${RENDER_CB}&render=explicit`;
        document.head.appendChild(s);
      }
      window[RENDER_CB] = function(){ try{ render(); }catch{} };
      // Safety timeout: if script fails to load, show fallback checkbox
      setTimeout(()=>{
        try{
          if (!window.grecaptcha || !window.grecaptcha.render){
            if (fbWrap) fbWrap.hidden = false;
            if (fbCb){ const onChange = ()=> setEnabled(!!fbCb.checked); fbCb.addEventListener('change', onChange); onChange(); }
          }
        }catch{}
      }, 6000);
    }catch{}
  }

  document.addEventListener('DOMContentLoaded', () => {
    if (!sb()) {
      alert('Supabase yapılandırması bulunamadı.');
      return;
    }
    try{ if (location.search) history.replaceState(null, '', location.pathname + location.hash); }catch{}
    setupAuth();
    setupTabs();
    try{ const btn = document.getElementById('logoutBtn'); if (btn && !btn.dataset.wired){ btn.dataset.wired='1'; btn.type='button'; btn.addEventListener('click', async (e)=>{ e.preventDefault(); try{ await sb().auth.signOut(); }catch{} try{ sessionStorage.clear(); }catch{} location.reload(); }); } }catch{}
    checkSession();
  });

  function setupAuth(){
    const form = qs('#loginForm');
    if (!form) return;
    try{ form.setAttribute('autocomplete','off'); }catch{}
    try{
      const emailIn = form.querySelector('input[name="email"]');
      const passIn = form.querySelector('input[name="password"]');
      if (emailIn) { emailIn.autocomplete='off'; emailIn.setAttribute('autocapitalize','none'); emailIn.setAttribute('spellcheck','false'); }
      if (passIn) {
        // Ensure normal keyboard shows on mobile for password input
        try{ passIn.autocomplete = 'current-password'; }catch{}
        try{ passIn.removeAttribute('inputmode'); }catch{}
        try{ passIn.setAttribute('autocapitalize','none'); passIn.setAttribute('spellcheck','false'); }catch{}
      }
    }catch{}
    // Initialize reCAPTCHA for login form (async, with fallback)
    try{ initLoginRecaptcha(); }catch{}
    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(form);
      const email = String(fd.get('email')|| '').trim();
      const password = String(fd.get('password')|| '').trim();
      // Enforce captcha: either token present or fallback checkbox checked
      try{
        const needEnforce = !!document.getElementById('recaptcha-login-container');
        if (needEnforce){
          const tokenOk = !!(window.__loginRecaptcha && window.__loginRecaptcha.token);
          const cb = document.getElementById('captchaLoginAgree');
          const cbOk = !!(cb && cb.checked);
          if (!tokenOk && !cbOk){ alert('Lütfen doğrulamayı tamamlayınız.'); return; }
        }
      }catch{}
      if (!email || !password) return;
      const submitBtn = form.querySelector('button[type="submit"]');
      if (submitBtn) submitBtn.disabled = true;
      try{
        // wait for supabase client to be ready
        let tries = 0;
        while ((!sb() || !sb().auth) && tries < 10){ await new Promise(r=>setTimeout(r,100)); tries++; }

        const attempt = async () => {
          const { error } = await sb().auth.signInWithPassword({ email, password });
          if (error) throw error;
        };
        try{
          await attempt();
        }catch(err1){
          // transient NetworkError retry once
          const msg = (err1?.message|| '').toLowerCase();
          if (msg.includes('network') || msg.includes('failed to fetch')){
            await new Promise(r=>setTimeout(r,300));
            await attempt();
          } else {
            throw err1;
          }
        }
        form.reset();
        await checkSession();
      } catch (err) {
        alert('Giriş başarısız: ' + (err?.message||String(err)));
      } finally {
        if (submitBtn) submitBtn.disabled = false;
      }
    });
  }
// ========== CHAIRMAN (GENEL BAŞKAN) ==========
async function loadAdminChairman(){
  try{
    const panel = document.getElementById('chairmanPanel'); if (!panel) return; panel.innerHTML = '';
    const reloadBtn = document.getElementById('reloadChairmanBtn');
    if (reloadBtn && !reloadBtn.dataset.wired){
      reloadBtn.dataset.wired='1';
      reloadBtn.addEventListener('click', (e)=>{ e.preventDefault(); loadAdminChairman(); });
    }

    // Fetch latest single record
    let row = { id:null, photo_url:', message_html:', status:'draft' };
    try{
      const { data } = await sb()
        .from('chairman')
        .select('id, photo_url, message_html, status, updated_at')
        .order('updated_at', { ascending:false, nullsFirst:true })
        .limit(1);
      if (data && data.length) row = data[0] || row;
    }catch{}

    // Layout: left photo+crop, right message
    const wrap = document.createElement('div');
    wrap.style.display='grid';
    wrap.style.gridTemplateColumns='minmax(280px, 380px) 1fr';
    wrap.style.gap='16px'; wrap.style.alignItems='start';

    // Left: Photo + crop
    const left = document.createElement('div');
    left.className='card'; left.style.padding='10px'; left.style.display='grid'; left.style.gap='8px'; left.style.marginTop='8px';

    const lTitle = document.createElement('div'); lTitle.textContent='Fotoğraf ve Kadraj'; lTitle.style.fontWeight='600'; left.appendChild(lTitle);

    const prev = document.createElement('div');
    prev.style.width='100%'; prev.style.aspectRatio='3/4';
    prev.style.border='1px solid #e5e7eb'; prev.style.borderRadius='10px'; prev.style.overflow='hidden'; prev.style.backgroundRepeat='no-repeat'; prev.style.backgroundPosition='center center'; prev.style.backgroundSize='cover';
    left.appendChild(prev);

    const iUrl = document.createElement('input'); iUrl.type='url'; iUrl.placeholder='https://... (opsiyonel)'; iUrl.value = stripPhotoKey(row.photo_url|| ''); left.appendChild(iUrl);
    const iFile = document.createElement('input'); iFile.type='file'; iFile.accept='image/*'; left.appendChild(iFile);

    const meta0 = parsePhotoMeta(row.photo_url|| '');
    const zLbl = document.createElement('label'); zLbl.style.display='grid'; zLbl.style.gap='4px'; zLbl.innerHTML='<span>Yakınlık</span>';
    const zRange = document.createElement('input'); zRange.type='range'; zRange.min='0.30'; zRange.max='2.00'; zRange.step='0.01'; zRange.value=String(meta0.z||1); zLbl.appendChild(zRange); left.appendChild(zLbl);

    const oyLbl = document.createElement('label'); oyLbl.style.display='grid'; oyLbl.style.gap='4px'; oyLbl.innerHTML='<span>Dikey Ofset</span>';
    const oyRange = document.createElement('input'); oyRange.type='range'; oyRange.min='-200'; oyRange.max='200'; oyRange.step='1'; oyRange.value=String(meta0.oy||0); oyLbl.appendChild(oyRange); left.appendChild(oyLbl);

    const oxLbl = document.createElement('label'); oxLbl.style.display='grid'; oxLbl.style.gap='4px'; oxLbl.innerHTML='<span>Yatay Ofset</span>';
    const oxRange = document.createElement('input'); oxRange.type='range'; oxRange.min='-200'; oxRange.max='200'; oxRange.step='1'; oxRange.value=String(meta0.ox||0); oxLbl.appendChild(oxRange); left.appendChild(oxLbl);

   // Önizleme
   const prevCanvas = document.createElement('canvas');
   prevCanvas.width = 240;
   prevCanvas.height = 320;
   prevCanvas.style.border = '1px solid #e5e7eb';
   prevCanvas.style.borderRadius = '10px';
   left.appendChild(prevCanvas);

   let prevImg = null;
   async function ensurePrevImg(){
     if (prevImg) return prevImg;
     try{
       if (iFile.files && iFile.files[0]){
         prevImg = await new Promise((res)=>{ const i=new Image(); i.onload=()=>res(i); try{i.src=URL.createObjectURL(iFile.files[0]);}catch{res(null);} });
         return prevImg;
       } else if (iUrl.value || row.photo_url){
         const u = stripPhotoKey(iUrl.value || row.photo_url || '');
         if (!u) return null;
         prevImg = await new Promise((res)=>{ const i=new Image(); i.crossOrigin='anonymous'; i.onload=()=>res(i); i.onerror=()=>res(null); i.src=u; });
         return prevImg;
       }
     }catch{}
     return null;
   }

   async function drawPrev(){
  const ctx = prevCanvas.getContext('2d');
  ctx.clearRect(0,0,prevCanvas.width,prevCanvas.height);
  const img = await ensurePrevImg();
  if (!img){ ctx.fillStyle='#e5e7eb'; ctx.fillRect(0,0,prevCanvas.width,prevCanvas.height); return; }
  // Scale to cover the canvas (rectangular), then apply user zoom
  const baseScale = Math.max(prevCanvas.width / img.width, prevCanvas.height / img.height);
  let scale = Math.max(baseScale, baseScale * Number(zRange.value||1));
  const w = img.width * scale, h = img.height * scale;
  const cx = prevCanvas.width/2, cy = prevCanvas.height/2;
  const x = cx - w/2 + Number(oxRange.value||0);
  const y = cy - h/2 + Number(oyRange.value||0);
  try{ ctx.drawImage(img, x, y, w, h); }catch{}
}
  drawPrev();

   zRange.addEventListener('input', drawPrev);
   oxRange.addEventListener('input', drawPrev);
   oyRange.addEventListener('input', drawPrev);
   iFile.addEventListener('change', ()=>{ prevImg = null; drawPrev(); });
   iUrl.addEventListener('change', ()=>{ prevImg = null; drawPrev(); });

   // Right: Message editor + status + actions
   const right = document.createElement('div'); right.className='card'; right.style.padding='10px'; right.style.display='grid'; right.style.gap='8px'; right.style.marginTop='8px';
   const rTitle = document.createElement('div'); rTitle.textContent='Mesaj Metni'; rTitle.style.fontWeight='600'; right.appendChild(rTitle);

   const tb=document.createElement('div'); tb.style.display='flex'; tb.style.flexWrap='wrap'; tb.style.gap='6px'; tb.style.margin='6px 0';
   const ed=document.createElement('div'); ed.contentEditable='true'; ed.className='card'; ed.style.minHeight='260px'; ed.style.padding='10px'; ed.style.overflow='auto'; ed.innerHTML = (row.message_html|| '');
   function btn(label, title, on){ const b=document.createElement('button'); b.type='button'; b.className='btn btn-outline'; b.textContent=label; b.title=title; b.style.padding='6px 10px'; b.addEventListener('click', (e)=>{ e.preventDefault(); on(); ed.focus(); }); return b; }
   function applyBlock(cmd){ document.execCommand(cmd,false,null); }
   tb.appendChild(btn('B','Kalın', ()=> applyBlock('bold')));
   tb.appendChild(btn('I','İtalik', ()=> applyBlock('italic')));
   tb.appendChild(btn('U','Altı Çizili', ()=> applyBlock('underline')));
   tb.appendChild(btn('Sol','Sola Hizala', ()=> applyBlock('justifyLeft')));
   tb.appendChild(btn('Ortala','Ortala', ()=> applyBlock('justifyCenter')));
   tb.appendChild(btn('Sağ','Sağa Hizala', ()=> applyBlock('justifyRight')));
   tb.appendChild(btn('İkiye','İki Yana Yasla', ()=> applyBlock('justifyFull')));
   tb.appendChild(btn('• Liste','Madde İşaretli', ()=> applyBlock('insertUnorderedList')));
   tb.appendChild(btn('1. Liste','Numaralı', ()=> applyBlock('insertOrderedList')));
   right.appendChild(tb);
   right.appendChild(ed);

   const stLbl=document.createElement('label'); stLbl.style.display='grid'; stLbl.style.gap='6px'; stLbl.innerHTML='<span>Durum</span>';
   const stSel=document.createElement('select'); ['draft','published','unpublished'].forEach(v=>{ const o=document.createElement('option'); o.value=v; o.textContent=({draft:'Taslak', published:'Yayımlandı', unpublished:'Yayından Kaldırıldı'}[v]||v); if ((row.status||'draft')===v) o.selected=true; stSel.appendChild(o); }); stLbl.appendChild(stSel); right.appendChild(stLbl);

   const actions=document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
   const saveBtn=document.createElement('button'); saveBtn.className='btn btn-success'; saveBtn.textContent='Kaydet';
   const cancelBtn=document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
   actions.appendChild(saveBtn); actions.appendChild(cancelBtn); right.appendChild(actions);

   cancelBtn.addEventListener('click', (e)=>{ e.preventDefault(); /* stay in tab */ });

   saveBtn.addEventListener('click', async (e)=>{
     e.preventDefault();
     try{
       let photoUrl = String(iUrl.value|| '').trim() || stripPhotoKey(row.photo_url|| '');
       const f = iFile.files && iFile.files[0];
       if (f){
         const ext = (f.name.split('.').pop()||'jpg').toLowerCase();
         const key = `chairman/${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
         photoUrl = await uploadToBucketGeneric(f, 'images', key);
       }
       const meta = { z: Number(zRange.value||1), ox: Number(oxRange.value||0), oy: Number(oyRange.value||0) };
       const finalPhoto = photoUrl ? buildPhotoValue(photoUrl, meta) : '';
       const payload = { photo_url: finalPhoto, message_html: String(ed.innerHTML|| '').trim(), status: stSel.value||'draft', updated_at: new Date().toISOString() };
       const q = row.id
         ? sb().from('chairman').update(payload).eq('id', row.id)
         : sb().from('chairman').insert(payload).select('id').single();
       const { error } = await q; if (error) throw error;
       alert('Kaydedildi');
       await loadAdminChairman();
     }catch(err){ alert('Kaydedilemedi: ' + (err?.message||String(err))); }
   });

   wrap.appendChild(left); wrap.appendChild(right);
   panel.appendChild(wrap);
  }catch(e){ alert('Genel Başkan formu açılamadı: ' + (e?.message||String(e))); }
}
// ========== KURUCULAR (FOUNDERS) ==========
async function loadAdminFounders(){
  const tbody = $foundersTBody(); if (!tbody) return;
  tbody.innerHTML = '';

  try{
    const table = tbody.closest('table');
    const headRow = table && table.tHead && table.tHead.rows && table.tHead.rows[0];
    if (headRow){
      headRow.innerHTML = '<th>Görsel</th><th>Ad Soyad</th><th>Sıra</th><th>Durum</th><th class="col-actions">İşlem</th>';
    }
  }catch{}

  try{
    const { data, error } = await sb()
      .from('founders')
      .select('id, name, image_url, sort, status')
      .order('sort', { ascending: true, nullsFirst: true })
      .order('name');
    if (error) throw error;
    (data||[]).forEach(row => {
      const tr=document.createElement('tr');
      const img = row.image_url ? `<img src="${escapeHtml(stripPhotoKey(row.image_url))}" alt="${escapeHtml(row.name||'Kurucu')}" />`  : '';
      tr.innerHTML = `
        <td>${img}</td>
        <td>${escapeHtml(row.name || '')}</td>
        <td>${Number(row.sort || '') || ''}</td>
        <td>${escapeHtml(({published:'Yayımlandı',draft:'Taslak',scheduled:'Planlı',archived:'Arşivli',unpublished:'Yayından Kaldırıldı'}[String(row.status|| '').toLowerCase()]||row.status||'draft'))}</td>
        <td class="actions">
          <button class="btn btn-warning" data-edit-founder="${row.id}">Düzenle</button>
          <button class="btn btn-danger" data-del-founder="${row.id}">Sil</button>
          ${String(row.status|| '').toLowerCase()==='published'
            ? `<button class="btn btn-danger" data-unpub-founder="${row.id}">Yayından Kaldır</button>` 
            : `<button class="btn btn-success" data-pub-founder="${row.id}">Yayınla</button>` 
          }
        </td>`;
      tbody.appendChild(tr);
    });

    // “Yeni Kurucu” butonunu bağla
    const addBtn = qs('#newFounderBtn');
    if (addBtn && !addBtn.dataset.wired){
      addBtn.dataset.wired = '1';
      addBtn.onclick = () =>
        openFounderModal({ id: null, name: ', image_url: ', sort: (Date.now() % 1000), status: 'draft' });
    }

    wireFoundersRowActions();
  }catch(e){
    alert('Kurucular yüklenemedi: ' + (e?.message || String(e)));
  }
}

function wireFoundersRowActions(){
  try{
    const tbody = $foundersTBody(); if (!tbody) return;

    tbody.querySelectorAll('button[data-edit-founder]').forEach(btn => {
      if (btn.dataset.wired) return; btn.dataset.wired = '1';
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-edit-founder');
        try{
          const { data } = await sb().from('founders').select('*').eq('id', id).maybeSingle();
          openFounderModal(data || { id, name:', image_url:', sort:1, status:'draft' });
        }catch(e){
          alert('Açılamadı: ' + (e?.message || String(e)));
        }
      });
    });

    tbody.querySelectorAll('button[data-del-founder]').forEach(btn => {
      if (btn.dataset.wired) return; btn.dataset.wired = '1';
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-del-founder');
        if (!confirm('Silinsin mi?')) return;
        try{
          const { error } = await sb().from('founders').delete().eq('id', id);
          if (error) throw error;
          await loadAdminFounders();
        }catch(e){
          alert('Silinemedi: ' + (e?.message || String(e)));
        }
      });
    });

    tbody.querySelectorAll('button[data-pub-founder]').forEach(btn => {
      if (btn.dataset.wired) return; btn.dataset.wired = '1';
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-pub-founder');
        try{
          const { error } = await sb().from('founders').update({ status:'published' }).eq('id', id);
          if (error) throw error;
          await loadAdminFounders();
        }catch(e){
          alert('Yayınlanamadı: ' + (e?.message || String(e)));
        }
      });
    });

    tbody.querySelectorAll('button[data-unpub-founder]').forEach(btn => {
      if (btn.dataset.wired) return; btn.dataset.wired = '1';
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-unpub-founder');
        try{
          const { error } = await sb().from('founders').update({ status:'unpublished' }).eq('id', id);
          if (error) throw error;
          await loadAdminFounders();
        }catch(e){
          alert('Yayından kaldırılamadı: ' + (e?.message || String(e)));
        }
      });
    });
  }catch{}
}

function openFounderModal(row){
  try{
    row = row || { id: null, name: '', image_url: '', sort: 1, status: 'draft' };
    $modalTitle().textContent = row.id ? 'Kurucuyu Düzenle' : 'Yeni Kurucu';
    const form = $modalForm(); form.innerHTML= '';

    // Ad Soyad
    const nLbl = document.createElement('label'); nLbl.style.display='grid'; nLbl.style.gap='6px'; nLbl.innerHTML='<span>Ad Soyad</span>';
    const nIn = document.createElement('input'); nIn.value=row.name|| ''; nLbl.appendChild(nIn); form.appendChild(nLbl);

    // Görsel (URL + Dosya Yükleme)
    const iLbl = document.createElement('label'); iLbl.style.display='grid'; iLbl.style.gap='6px'; iLbl.innerHTML='<span>Görsel URL</span>';
    const iIn = document.createElement('input'); iIn.placeholder='https://...'; iIn.value=row.image_url|| '';
    const iFile = document.createElement('input'); iFile.type='file'; iFile.accept='image/*';
    const iWrap = document.createElement('div'); iWrap.style.display='grid'; iWrap.style.gap='6px'; iWrap.appendChild(iIn); iWrap.appendChild(iFile); iLbl.appendChild(iWrap); form.appendChild(iLbl);

    // Fotoğraf Kadrajı (kurucular için)
    const cropWrap = document.createElement('div');
    cropWrap.className = 'card';
    cropWrap.style.padding = '10px';
    cropWrap.style.display = 'grid';
    cropWrap.style.gap = '8px';
    cropWrap.style.marginTop = '8px';

    const cropTitle = document.createElement('div');
    cropTitle.textContent = 'Fotoğraf Kadrajı';
    cropTitle.style.fontWeight = '600';
    cropWrap.appendChild(cropTitle);

    const initMeta = parsePhotoMeta(row.image_url || '');

    // Zoom
    const zoomLbl = document.createElement('label');
    zoomLbl.style.display = 'grid';
    zoomLbl.style.gap = '4px';
    zoomLbl.innerHTML = '<span>Yakınlık</span>';
    const zRange = document.createElement('input');
    zRange.type = 'range';
    zRange.min = '0.30';
    zRange.max = '2.00';
    zRange.step = '0.01';
    zRange.value = String(initMeta.z || 1);
    zoomLbl.appendChild(zRange);
    cropWrap.appendChild(zoomLbl);

    // Dikey Ofset
    const oyLbl = document.createElement('label');
    oyLbl.style.display = 'grid';
    oyLbl.style.gap = '4px';
    oyLbl.innerHTML = '<span>Dikey Ofset</span>';
    const oyRange = document.createElement('input');
    oyRange.type = 'range';
    oyRange.min = '-160';
    oyRange.max = '160';
    oyRange.step = '1';
    oyRange.value = String(initMeta.oy || 0);
    oyLbl.appendChild(oyRange);
    cropWrap.appendChild(oyLbl);

    // Yatay Ofset
    const oxLbl = document.createElement('label');
    oxLbl.style.display = 'grid';
    oxLbl.style.gap = '4px';
    oxLbl.innerHTML = '<span>Yatay Ofset</span>';
    const oxRange = document.createElement('input');
    oxRange.type = 'range';
    oxRange.min = '-160';
    oxRange.max = '160';
    oxRange.step = '1';
    oxRange.value = String(initMeta.ox || 0);
    oxLbl.appendChild(oxRange);
    cropWrap.appendChild(oxLbl);

   // Önizleme
   const prevCanvas = document.createElement('canvas');
   prevCanvas.width = 240;
   prevCanvas.height = 320;
   prevCanvas.style.border = '1px solid #e5e7eb';
   prevCanvas.style.borderRadius = '10px';
   cropWrap.appendChild(prevCanvas);

   let prevImg = null;
   async function ensurePrevImg(){
     if (prevImg) return prevImg;
     try{
       if (iFile.files && iFile.files[0]){
         prevImg = await new Promise((res)=>{ const i=new Image(); i.onload=()=>res(i); try{i.src=URL.createObjectURL(iFile.files[0]);}catch{res(null);} });
         return prevImg;
       } else if (iIn.value || row.image_url){
         const u = stripPhotoKey(iIn.value || row.image_url || '');
         if (!u) return null;
         prevImg = await new Promise((res)=>{ const i=new Image(); i.crossOrigin='anonymous'; i.onload=()=>res(i); i.onerror=()=>res(null); i.src=u; });
         return prevImg;
       }
     }catch{}
     return null;
   }

   async function drawPrev(){
  const ctx = prevCanvas.getContext('2d');
  ctx.clearRect(0,0,prevCanvas.width,prevCanvas.height);
  const img = await ensurePrevImg();
  if (!img){ ctx.fillStyle='#e5e7eb'; ctx.fillRect(0,0,prevCanvas.width,prevCanvas.height); return; }
  // Scale to cover the canvas (rectangular), then apply user zoom
  const baseScale = Math.max(prevCanvas.width / img.width, prevCanvas.height / img.height);
  let scale = Math.max(baseScale, baseScale * Number(zRange.value||1));
  const w = img.width * scale, h = img.height * scale;
  const cx = prevCanvas.width/2, cy = prevCanvas.height/2;
  const x = cx - w/2 + Number(oxRange.value||0);
  const y = cy - h/2 + Number(oyRange.value||0);
  try{ ctx.drawImage(img, x, y, w, h); }catch{}
}
  drawPrev();

   zRange.addEventListener('input', drawPrev);
   oxRange.addEventListener('input', drawPrev);
   oyRange.addEventListener('input', drawPrev);
   iFile.addEventListener('change', ()=>{ prevImg = null; drawPrev(); });
   iIn.addEventListener('change', ()=>{ prevImg = null; drawPrev(); });

   drawPrev();

   // Sıra
   const sLbl = document.createElement('label'); sLbl.style.display='grid'; sLbl.style.gap='6px'; sLbl.innerHTML='<span>Sıra</span>';
   const sIn = document.createElement('input'); sIn.type='number'; sIn.step='1'; sIn.value = String(row.sort || 1); sLbl.appendChild(sIn); form.appendChild(sLbl);

   // Durum
   const stLbl = document.createElement('label'); stLbl.style.display='grid'; stLbl.style.gap='6px'; stLbl.innerHTML='<span>Durum</span>';
   const stSel = document.createElement('select');
   ['draft', 'published', 'unpublished'].forEach(v => {
     const o = document.createElement('option'); o.value = v; o.textContent = ({draft:'Taslak', published:'Yayımlandı', unpublished:'Yayından Kaldırıldı'}[v]||v);
     if ((row.status || 'draft') === v) o.selected = true;
     stSel.appendChild(o);
   });
   stLbl.appendChild(stSel); form.appendChild(stLbl);

   // Actions
   const actions = document.createElement('div'); actions.style.display='flex'; actions.style.gap='8px';
   const saveBtn = document.createElement('button');
   saveBtn.type = 'button';
   saveBtn.className = 'btn btn-success';
   saveBtn.textContent = 'Kaydet';
   actions.appendChild(saveBtn);
   form.appendChild(actions);
   const cancelBtn = document.createElement('button'); cancelBtn.className='btn btn-danger'; cancelBtn.textContent='İptal';
   actions.appendChild(cancelBtn);
   cancelBtn.addEventListener('click', (e)=>{ e.preventDefault(); closeModal(); });

   saveBtn.addEventListener('click', async (e)=>{
     e.preventDefault();
     const name = String(nIn.value || '').trim(); if (!name) return alert('Ad Soyad gerekli');
     let image_url = String(iIn.value || '').trim();
     const sort = Number(sIn.value || 1) || 1;
     const status = stSel.value || 'draft';
     try{
       // Dosya seçilmişse yükle
       const f = iFile.files && iFile.files[0];
       if (f){
         const ext = (f.name.split('.').pop()||'jpg').toLowerCase();
         const key = `founders/${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
         try{
           image_url = await uploadToBucketGeneric(f, 'founders', key);
         }catch(err){
           return alert('Görsel yüklenemedi: ' + (err?.message || String(err)));
         }
       }

       // Kadraj bilgisini URL'e ekle
       const cropMeta = {
        z: Number((zRange && zRange.value) || 1),
        ox: Number((oxRange && oxRange.value) || 0),
        oy: Number((oyRange && oyRange.value) || 0),
      };
      if (image_url) image_url = buildPhotoValue(image_url, cropMeta);

       const payload = { name, image_url, sort, status };
       const q = row.id
         ? sb().from('founders').update(payload).eq('id', row.id)
         : sb().from('founders').insert(payload).select('id').single();
       const { error } = await q;
       if (error) throw error;

       closeModal();
       await loadAdminFounders();
     }catch(err){
       alert('Kaydedilemedi: ' + (err?.message || String(err)));
     }
   });

   (typeof openModal === 'function' ? openModal() : (window.openModal && window.openModal()));
 }catch(e){
   alert('Kurucu formu açılamadı: ' + (e?.message || String(e)));
 }
}
})();






























