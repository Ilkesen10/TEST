// Supabase Edge Function: auth_admin
// Purpose: Create or update a Supabase Auth user with a given password (admin-only)
// Env required (set in Supabase project settings):
// - SUPABASE_URL
// - SUPABASE_SERVICE_ROLE_KEY (never expose client-side)
//
// Request (JSON):
// { action: 'upsert_user', email: string, password: string }
//
// Response: 200 { ok: true, user: { id, email } } or { error }

import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const supabaseUrl = Deno.env.get("SUPABASE_URL") ?? "";
const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
const anonKey = Deno.env.get("SUPABASE_ANON_KEY") ?? "";
if (!supabaseUrl || !serviceKey || !anonKey) {
  console.error("Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY or SUPABASE_ANON_KEY environment variables");
}
// Admin client (service role) for privileged operations
const sbAdmin = createClient(supabaseUrl, serviceKey, {
  auth: { persistSession: false }
});
// Anon client for validating Bearer tokens from callers
const sbAnon = createClient(supabaseUrl, anonKey, {
  auth: { persistSession: false }
});

function corsHeaders(origin?: string) {
  return {
    "Access-Control-Allow-Origin": origin || "*",
    "Access-Control-Allow-Headers": "authorization, content-type, x-client-info, apikey",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
  } as HeadersInit;
}

async function findUserByEmail(email: string) {
  // Try listUsers pages to find by email (limited pages for safety)
  let page = 1;
  const perPage = 200;
  for (let i = 0; i < 10; i++) { // up to 2000 users scan
    const { data, error } = await sbAdmin.auth.admin.listUsers({ page, perPage });
    if (error) throw error;
    const hit = data?.users?.find(u => (u.email || "").toLowerCase() === email.toLowerCase());
    if (hit) return hit;
    if (!data || data.users.length < perPage) break; // no more pages
    page++;
  }
  return null;
}

serve(async (req) => {
  const origin = req.headers.get("Origin") || "*";
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders(origin) });
  }

  try {
    // AuthN: require Bearer token and validate user
    const authH = req.headers.get("Authorization") || "";
    const token = authH.startsWith("Bearer ") ? authH.slice(7) : "";
    if (!token) {
      return new Response(JSON.stringify({ error: "unauthorized" }), { status: 401, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    }
    const { data: userData, error: userErr } = await sbAnon.auth.getUser(token);
    if (userErr || !userData?.user?.email) {
      return new Response(JSON.stringify({ error: "unauthorized" }), { status: 401, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    }

    // AuthZ: only allow admin/superadmin emails found in admin_users table
    const email = String(userData.user.email || "");
    const { data: adminRow, error: adminErr } = await sbAdmin
      .from("admin_users")
      .select("roles")
      .eq("email", email)
      .maybeSingle();
    if (adminErr || !adminRow) {
      return new Response(JSON.stringify({ error: "forbidden" }), { status: 403, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    }
    const roles = Array.isArray((adminRow as any).roles) ? (adminRow as any).roles.map((r: any) => String(r).toLowerCase()) : [];
    const allowed = roles.includes("superadmin") || roles.includes("admin");
    if (!allowed) {
      return new Response(JSON.stringify({ error: "forbidden" }), { status: 403, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    }

    const body = await req.json().catch(() => ({}));
    const action = body?.action;

    if (action !== "upsert_user") {
      return new Response(JSON.stringify({ error: "invalid_action" }), { status: 400, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    }

    const email = String(body?.email || "").trim();
    const password = String(body?.password || "").trim();
    if (!email || !password) {
      return new Response(JSON.stringify({ error: "email_and_password_required" }), { status: 400, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    }

    // Try create user first
    const { data: created, error: createErr } = await sbAdmin.auth.admin.createUser({
      email,
      password,
      email_confirm: true,
    });
    if (!createErr && created?.user) {
      return new Response(JSON.stringify({ ok: true, user: { id: created.user.id, email: created.user.email } }), { status: 200, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    }

    // If exists, update password by user id
    const existing = await findUserByEmail(email);
    if (!existing) {
      return new Response(JSON.stringify({ error: createErr?.message || "user_not_found" }), { status: 400, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    }
    const { data: updated, error: updErr } = await sbAdmin.auth.admin.updateUserById(existing.id, { password });
    if (updErr) {
      return new Response(JSON.stringify({ error: updErr.message }), { status: 400, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
    }
    return new Response(JSON.stringify({ ok: true, user: { id: existing.id, email: existing.email } }), { status: 200, headers: { "Content-Type": "application/json", ...corsHeaders(origin) } });
  } catch (e) {
    return new Response(JSON.stringify({ error: e?.message || String(e) }), { status: 500, headers: { "Content-Type": "application/json", ...corsHeaders() } });
  }
});
