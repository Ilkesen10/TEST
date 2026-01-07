// Supabase Edge Function: admin_send_mail (JWT + RBAC check)
// Env vars to set in Supabase Function config:
// - RESEND_API_KEY
// - RESEND_FROM (e.g., 'noreply@ilkesen.org.tr' - must be verified in Resend)
// - REPLY_TO (optional)
// - SUPABASE_URL, SUPABASE_ANON_KEY, SUPABASE_SERVICE_ROLE_KEY
// deno-lint-ignore-file no-explicit-any
/// <reference lib="dom" />
// @ts-ignore Remote import is resolved by Deno at runtime
import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
// @ts-ignore Supabase client ESM for Deno
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

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

const ENV = (globalThis as any)?.Deno?.env;
const SUPABASE_URL = ENV?.get("SUPABASE_URL") || "";
const SUPABASE_ANON_KEY = ENV?.get("SUPABASE_ANON_KEY") || "";
const SUPABASE_SERVICE_ROLE_KEY = ENV?.get("SUPABASE_SERVICE_ROLE_KEY") || "";
const RESEND_API_KEY = ENV?.get("RESEND_API_KEY") || "";
const RESEND_FROM = ENV?.get("RESEND_FROM") || "onboarding@resend.dev";
const REPLY_TO = ENV?.get("REPLY_TO") || "";

const sbAnon = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, { auth: { persistSession: false } });
const sbAdmin = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, { auth: { persistSession: false } });

serve(async (req) => {
  const origin = req.headers.get("Origin") || "*";
  if (req.method === "OPTIONS") return new Response(null, { headers: corsHeaders(origin) });

  try {
    if (req.method !== "POST") {
      return new Response(JSON.stringify({ ok: false, error: "method_not_allowed" }), { status: 405, headers: jsonHeaders(origin) });
    }

    // 1) AuthN: require Bearer token
    const authH = req.headers.get("Authorization") || "";
    const token = authH.startsWith("Bearer ") ? authH.slice(7) : "";
    if (!token) {
      return new Response(JSON.stringify({ ok: false, error: "unauthorized" }), { status: 401, headers: jsonHeaders(origin) });
    }
    const { data: userData, error: userErr } = await sbAnon.auth.getUser(token);
    if (userErr || !userData?.user?.email) {
      return new Response(JSON.stringify({ ok: false, error: "unauthorized" }), { status: 401, headers: jsonHeaders(origin) });
    }
    const email = userData.user.email as string;

    // 2) AuthZ: check admin role
    const { data: adminRow, error: adminErr } = await sbAdmin
      .from("admin_users")
      .select("roles")
      .eq("email", email)
      .maybeSingle();
    if (adminErr || !adminRow) {
      return new Response(JSON.stringify({ ok: false, error: "forbidden" }), { status: 403, headers: jsonHeaders(origin) });
    }
    const roles = Array.isArray(adminRow.roles) ? adminRow.roles.map((r: any) => String(r).toLowerCase()) : [];
    const allowed = roles.includes("superadmin") || roles.includes("admin");
    if (!allowed) {
      return new Response(JSON.stringify({ ok: false, error: "forbidden" }), { status: 403, headers: jsonHeaders(origin) });
    }

    // 3) Parse input
    const body = await req.json().catch(() => ({}));
    let to = String(body?.to || "").trim().replace(/^mailto:/i, "");
    const subject = String(body?.subject || "").trim().slice(0, 200);
    const message = String(body?.message || "").trim().slice(0, 10000);
   
  // ALAN BAZLI DOĞRULAMA
if (!to || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) {
  return new Response(JSON.stringify({ ok: false, error: "invalid_input", field: "to", value: to }), { status: 400, headers: jsonHeaders(origin) });
}
if (!subject) {
  return new Response(JSON.stringify({ ok: false, error: "invalid_input", field: "subject" }), { status: 400, headers: jsonHeaders(origin) });
}
if (!message) {
  return new Response(JSON.stringify({ ok: false, error: "invalid_input", field: "message" }), { status: 400, headers: jsonHeaders(origin) });
}

    if (!RESEND_API_KEY) {
      return new Response(JSON.stringify({ ok: false, error: "server_not_configured" }), { status: 500, headers: jsonHeaders(origin) });
    }

    const html = `
      <div>
        <div>${message.replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\\n/g, "<br/>")}</div>
      </div>
    `;

    const resp = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${RESEND_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from: RESEND_FROM,
        to: [to],
        reply_to: REPLY_TO || email, // allow replies to the admin sender by default
        subject,
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