// Supabase Edge Function: send_contact
// Env (set in Supabase project's function config):
// - RESEND_API_KEY
// - RESEND_FROM (e.g., 'noreply@ilkesen.org.tr' - must be verified in Resend)
// - CONTACT_TO (e.g., 'info@ilkesen.org.tr')
// - RECAPTCHA_SECRET (optional; if set, token will be verified)
/// <reference lib="dom" />
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

function jsonHeaders(origin?: string): Headers {
  const h = new Headers(corsHeaders(origin));
  h.set("Content-Type", "application/json");
  return h;
}

async function verifyRecaptcha(token: string | null, remoteIp?: string) {
  const secret = (globalThis as any)?.Deno?.env?.get("RECAPTCHA_SECRET") || "";
  if (!secret || !token) return true; // soft-allow if not configured
  try {
    const r = await fetch("https://www.google.com/recaptcha/api/siteverify", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ secret, response: token, remoteip: remoteIp || "" }),
    });
    const j = await r.json().catch(() => ({}));
    return !!j.success;
  } catch {
    return false;
  }
}

serve(async (req) => {
  const origin = req.headers.get("Origin") || "*";
  if (req.method === "OPTIONS") return new Response(null, { headers: corsHeaders(origin) });

  try {
    if (req.method !== "POST") {
      return new Response(JSON.stringify({ error: "method_not_allowed" }), { status: 405, headers: jsonHeaders(origin) });
    }

    const body = await req.json().catch(() => ({}));
    const name = String(body?.name || "").trim().slice(0, 200);
    const email = String(body?.email || "").trim().slice(0, 320);
    const subject = String(body?.subject || "").trim().slice(0, 200);
    const message = String(body?.message || "").trim().slice(0, 5000);
    const captcha = String(body?.captcha_token || "").trim();

    if (!name || !email || !subject || !message) {
      return new Response(JSON.stringify({ ok: false, error: "invalid_input" }), { status: 400, headers: jsonHeaders(origin) });
    }

    const ip = req.headers.get("X-Forwarded-For") || undefined;
    const captchaOk = await verifyRecaptcha(captcha || null, ip);
    if (!captchaOk) {
      return new Response(JSON.stringify({ ok: false, error: "captcha_failed" }), { status: 400, headers: jsonHeaders(origin) });
    }

    const ENV = (globalThis as any)?.Deno?.env;
    const RESEND_API_KEY = ENV?.get("RESEND_API_KEY") || "";
    const FROM = ENV?.get("RESEND_FROM") || "onboarding@resend.dev"; // dev fallback
    const TO = ENV?.get("CONTACT_TO") || "";
    if (!RESEND_API_KEY || !TO) {
      return new Response(JSON.stringify({ ok: false, error: "server_not_configured" }), { status: 500, headers: jsonHeaders(origin) });
    }

    const html = `
      <div>
        <h2>Yeni İletişim Mesajı</h2>
        <p><strong>İsim:</strong> ${name}</p>
        <p><strong>E-posta:</strong> ${email}</p>
        <p><strong>Konu:</strong> ${subject}</p>
        <p><strong>Mesaj:</strong></p>
        <div>${message.replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br/>")}</div>
      </div>
    `;

    const resp = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${RESEND_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from: FROM,
        to: [TO],
        reply_to: email,
        subject: `[İletişim] ${subject}`,
        html,
      }),
    });

    if (!resp.ok) {
      const errTxt = await resp.text().catch(() => "");
      return new Response(JSON.stringify({ ok: false, error: "resend_failed", details: errTxt }), { status: 502, headers: jsonHeaders(origin) });
    }

    return new Response(JSON.stringify({ ok: true }), { status: 200, headers: jsonHeaders(origin) });
  } catch (e) {
    return new Response(JSON.stringify({ ok: false, error: e?.message || String(e) }), { status: 500, headers: jsonHeaders(origin) });
  }
});