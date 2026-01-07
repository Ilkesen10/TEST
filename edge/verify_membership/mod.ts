// Supabase Edge Function: verify_membership
// Purpose: Verify a member by QR token and return limited public info
// Env required (set in Supabase project settings):
// - SUPABASE_URL
// - SUPABASE_SERVICE_ROLE_KEY (never expose client-side)
//
// Request: GET /verify_membership?t=TOKEN or POST { t: TOKEN }
// Response: 200 { ok:true, verified:true, member:{ full_name, membership_no, status, join_date } } or { ok:true, verified:false }
// @ts-ignore Remote import is resolved by Deno at runtime
import { serve } from "https://deno.land/std@0.224.0/http/server.ts?target=deno";
// @ts-ignore Supabase client ESM for Deno
import { createClient } from "https://esm.sh/@supabase/supabase-js@2?target=deno";

const supabaseUrl = (globalThis as any).Deno?.env?.get("SUPABASE_URL") ?? "";
const serviceKey = (globalThis as any).Deno?.env?.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
const db = createClient(supabaseUrl, serviceKey, { auth: { persistSession: false } });

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

serve(async (req) => {
  const origin = req.headers.get("Origin") || "*";
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders(origin) });
  }

  try {
    if (req.method !== "GET" && req.method !== "POST") {
      return new Response(JSON.stringify({ error: "method_not_allowed" }), { status: 405, headers: jsonHeaders(origin) });
    }

    let token = "";
    try{
      if (req.method === "GET"){
        const u = new URL(req.url);
        token = String(u.searchParams.get("t") || u.searchParams.get("token") || "").trim();
      } else if (req.method === "POST"){
        const body = await req.json().catch(()=>({} as any));
        token = String(body?.t || body?.token || "").trim();
      }
    }catch{}

    if (!token || token.length > 256) {
      return new Response(JSON.stringify({ ok: true, verified: false, reason: "invalid" }), { status: 200, headers: jsonHeaders(origin) });
    }

    // Query members table by token
    const { data, error } = await db
      .from("members")
      .select("id, first_name, last_name, member_no, status, join_date")
      .eq("verify_token", token)
      .maybeSingle();

    if (error) {
      console.error("verify_membership query error", error);
      return new Response(JSON.stringify({ ok: false, error: error.message }), { status: 500, headers: jsonHeaders(origin) });
    }

    if (!data) {
      return new Response(JSON.stringify({ ok: true, verified: false, reason: "not_found" }), { status: 200, headers: jsonHeaders(origin) });
    }

    const isActive = String(data.status || "").toLowerCase() === "active" || String(data.status || "").toLowerCase() === "aktif";
    if (!isActive) {
      return new Response(JSON.stringify({ ok: true, verified: false, reason: "inactive" }), { status: 200, headers: jsonHeaders(origin) });
    }

    const member = {
      full_name: `${data.first_name || ''} ${data.last_name || ''}`.trim(),
      membership_no: data.member_no || "",
      status: data.status || "",
      join_date: data.join_date || null,
    };

    return new Response(JSON.stringify({ ok: true, verified: true, member }), { status: 200, headers: jsonHeaders(origin) });
  } catch (e) {
    console.error("verify_membership error", e);
    return new Response(JSON.stringify({ ok: false, error: e?.message || String(e) }), { status: 500, headers: jsonHeaders(origin) });
  }
});
