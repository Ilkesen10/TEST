// Supabase Edge Function (Deno)
// Name: idcard
// Purpose: Receive { type: 'digital_id', member } and send a digital ID email.
// Requirements:
// - Set RESEND_API_KEY as an environment variable in Supabase project.
// - Optionally set IDCARD_BG_URL to a public PNG for the ID background (export from kimlik.psd).
// - This function renders a simple HTML-based digital ID (no server-side image composition).
//   You can replace the renderer with a canvas/pdf solution later.

// @ts-ignore Remote import is resolved by Deno at runtime
import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
function corsHeaders(origin?: string) {
  return {
    "Access-Control-Allow-Origin": origin || "*",
    "Access-Control-Allow-Headers": "authorization, content-type, x-client-info, apikey",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Cache-Control": "no-store",
  } as HeadersInit;
}
const ENV = (globalThis as any)?.Deno?.env;
const RESEND_API_KEY = ENV?.get("RESEND_API_KEY") ?? "";
const RESEND_FROM = ENV?.get("RESEND_FROM") ?? "onboarding@resend.dev"; // Resend demo from (test için)
const IDCARD_BG_URL = ENV?.get("IDCARD_BG_URL") ?? "";

function maskTc(tc?: string) {
  if (!tc) return "";
  if (!/^\d{11}$/.test(tc)) return tc;
  return tc.slice(0, 6) + "*****"; // 6 visible + 5 masked, adjust to your policy
}

function formatPhone(e164?: string) {
  if (!e164) return "";
  const d = e164.replace(/\D/g, "");
  let p = d;
  if (p.startsWith("90")) p = p.slice(2);
  if (p.startsWith("0")) p = p.slice(1);
  if (p.length !== 10) return e164;
  return `+90 ${p.slice(0,3)} ${p.slice(3,6)} ${p.slice(6,8)} ${p.slice(8,10)}`;
}

function slugifyNameForFile(input: string): string {
  try{
    let base = String(input || '').trim();
    if (!base) return 'uye';
    base = base.normalize('NFD').replace(/[\u0300-\u036f]/g,'');
    base = base
      .replace(/ı/g,'i').replace(/İ/g,'I')
      .replace(/ş/g,'s').replace(/Ş/g,'S')
      .replace(/ğ/g,'g').replace(/Ğ/g,'G')
      .replace(/ç/g,'c').replace(/Ç/g,'C')
      .replace(/ö/g,'o').replace(/Ö/g,'O')
      .replace(/ü/g,'u').replace(/Ü/g,'U');
    base = base.toLowerCase().replace(/[^a-z0-9]+/g,'_').replace(/^_+|_+$/g,'');
    return base || 'uye';
  }catch{ return 'uye'; }
}

function renderHtmlCard(member: any) {
  const fullName = `${member.first_name ?? ''} ${member.last_name ?? ''}`.trim();
  const maskedTc = maskTc(member.national_id);
  const phone = formatPhone(member.phone);
  const email = member.email ?? '';
  const memberNo = member.member_no ?? '';
  const title = member.title ?? '';
  const institution = member.institution_name ?? '';
  const blood = member.blood_type ?? '';
  const joinDate = member.join_date ?? '';
  const photo = member.photo_url ?? '';

  // Simple email-safe HTML with an optional background image
  return `
  <div style="font-family:Arial,Helvetica,sans-serif;color:#0f172a;max-width:520px;margin:0 auto">
    <div style="text-align:center;margin-bottom:16px;background:#0B3A60;padding:12px 8px;border-radius:10px">
      <h1 style="font-size:36px;font-weight:bold;margin:0 0 4px;color:#ffffff;letter-spacing:1px">İLKE-SEN</h1>
      <div style="font-size:12px;color:#E03B3B;margin-bottom:8px">İLKELİ YEREL YÖNETİM HİZMETLERİ KOLU KAMU GÖREVLİLERİ SENDİKASI</div>
      <img src="https://www.ilke-sen.org.tr/ilke-sen-logo.png" alt="İLKE-SEN Logo" style="max-width:250px;height:auto">
    </div>
    <div style="border:1px solid #e5e7eb;border-radius:12px;overflow:hidden;max-width:100%">
      ${IDCARD_BG_URL ? `<div style=\"background:url('${IDCARD_BG_URL}') center/cover no-repeat;\">` : '<div>'}
        <div style="display:flex;gap:16px;padding:16px;background:rgba(255,255,255,0.94);align-items:center">
          <div style="flex:0 0 120px;height:120px;border-radius:9999px;overflow:hidden;background:#f1f5f9;border:4px solid #0B3A60;box-shadow:0 0 0 4px #E03B3B">
            ${photo ? `<img src="${photo}" alt="" style="width:100%;height:100%;object-fit:cover"/>` : ''}
          </div>
          <div style="flex:1;display:flex;flex-direction:column;gap:6px">
            <div style="font-size:20px;font-weight:800;color:#0B3A60;letter-spacing:0.5px;text-transform:uppercase">${fullName}</div>
            ${title ? `<div style=\"font-size:16px;font-weight:700;color:#E03B3B;text-transform:uppercase\">${title}</div>` : ''}
            <div style="display:flex;flex-direction:column;gap:6px;margin-top:8px">
              <div><span style="color:#E03B3B;font-weight:700;display:inline-block;min-width:140px">Üye Numarası</span><span style="color:#0B3A60;font-weight:700">: ${memberNo}</span></div>
              <div><span style="color:#E03B3B;font-weight:700;display:inline-block;min-width:140px">Üyelik Tarihi</span><span style="color:#0B3A60;font-weight:700">: ${joinDate}</span></div>
              <div><span style="color:#E03B3B;font-weight:700;display:inline-block;min-width:140px">TC. No</span><span style="color:#0B3A60;font-weight:700">: ${maskedTc}</span></div>
            </div>
          </div>
        </div>
      </div>
    </div>
    <p style="color:#64748b;font-size:12px">Bu e-posta, üyeliğiniz kapsamında oluşturulan dijital kimlik kartını içerir.</p>
  </div>`;
}

async function sendEmail(to: string, subject: string, html: string, attachment?: { filename: string; content: string }) {
  if (!RESEND_API_KEY) throw new Error("RESEND_API_KEY tanımlı değil");
  const payload: any = {
    from: RESEND_FROM,
    to: [to],
    subject,
    html,
  };
  if (attachment?.content && attachment?.filename) {
    payload.attachments = [{ filename: attachment.filename, content: attachment.content }];
  }
  const resp = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${RESEND_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  if (!resp.ok) {
    const txt = await resp.text();
    throw new Error(`Resend error: ${resp.status} ${txt}`);
  }
}

serve(async (req: Request) => {
  const origin = req.headers.get("Origin") || "*";
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders(origin) });
  }
  try {
    if (req.method !== "POST") return new Response("Method Not Allowed", { status: 405, headers: corsHeaders(origin) });
    const payload = await req.json();
    if (payload?.type !== "digital_id" || !payload?.member) {
      return new Response("Bad Request", { status: 400, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    }
    const m = payload.member;
    if (!m.email) return new Response("Member email missing", { status: 400, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    let attachment: { filename: string; content: string } | undefined = undefined;
    try {
      const dataUrl: string | null = payload?.image_base64 || null;
      if (typeof dataUrl === 'string' && dataUrl.startsWith('data:image/')) {
        const comma = dataUrl.indexOf(',');
        if (comma > -1) {
          const metaEnd = dataUrl.indexOf(';', 5);
          const meta = dataUrl.substring(5, metaEnd !== -1 ? metaEnd : comma).toLowerCase(); // e.g. image/jpeg
          let ext = 'jpg';
          if (meta.includes('png')) ext = 'png';
          const base64 = dataUrl.substring(comma + 1);
          const fullName = `${m.first_name ?? ''} ${m.last_name ?? ''}`.trim();
          const base = slugifyNameForFile(fullName || String(m.member_no || m.id || 'uye'));
          const fname = `kimlik_${base}.${ext}`;
          attachment = { filename: fname, content: base64 };
        }
      }
    } catch {}
    const html = attachment
      ? '<p>Merhaba, İLKE-SEN ailemize hoşgeldiniz. Sizi aramızda görmekten mutluluk duyuyoruz. Dijital kimlik kartınız ekte sunulmuştur.</p>'
      : renderHtmlCard(m);
    await sendEmail(m.email, "Dijital Üyelik Kimliği", html, attachment);
    return new Response(JSON.stringify({ ok: true }), { status: 200, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return new Response(JSON.stringify({ ok: false, error: msg }), { status: 500, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
  }
});
