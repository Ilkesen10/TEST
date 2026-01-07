// deno-lint-ignore-file no-explicit-any
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

function corsHeaders(origin?: string) {
  return {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": origin || "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  } as Record<string, string>;
}

serve(async (req: Request) => {
  const origin = req.headers.get("origin") || undefined;
  if (req.method === "OPTIONS") {
    return new Response("", { headers: corsHeaders(origin) });
  }

  const headers = corsHeaders(origin);

  try {
    const { payload, captcha_token } = await req.json();

    if (!payload || typeof payload !== "object") {
      return new Response(JSON.stringify({ error: "invalid_payload" }), { status: 400, headers });
    }

    if (!captcha_token || typeof captcha_token !== "string") {
      return new Response(JSON.stringify({ error: "missing_captcha_token" }), { status: 400, headers });
    }

    const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
    const SERVICE_ROLE = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    const RECAPTCHA_SECRET = Deno.env.get("RECAPTCHA_SECRET_KEY") || Deno.env.get("RECAPTCHA_SECRET");

    if (!SUPABASE_URL || !SERVICE_ROLE) {
      return new Response(
        JSON.stringify({ error: "server_misconfigured", hint: "Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY" }),
        { status: 500, headers }
      );
    }
    if (!RECAPTCHA_SECRET) {
      return new Response(
        JSON.stringify({ error: "server_misconfigured", hint: "Missing RECAPTCHA_SECRET_KEY/RECAPTCHA_SECRET" }),
        { status: 500, headers }
      );
    }

    // Verify reCAPTCHA with Google
    const form = new URLSearchParams();
    form.set("secret", RECAPTCHA_SECRET);
    form.set("response", captcha_token);
    const remoteIp = req.headers.get("x-forwarded-for") || "";
    if (remoteIp) form.set("remoteip", remoteIp);

    const verifyResp = await fetch("https://www.google.com/recaptcha/api/siteverify", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: form.toString(),
    });
    const verifyJson = await verifyResp.json();
    if (!verifyJson?.success) {
      return new Response(
        JSON.stringify({ error: "captcha_failed", details: verifyJson }),
        { status: 400, headers }
      );
    }

    const supabase = createClient(SUPABASE_URL, SERVICE_ROLE);

    // Normalize documents_urls to string JSON as used on the site
    let documents_urls: string = "[]";
    try {
      if (typeof payload.documents_urls === "string") {
        // ensure it's valid JSON
        JSON.parse(payload.documents_urls);
        documents_urls = payload.documents_urls;
      } else if (Array.isArray(payload.documents_urls)) {
        documents_urls = JSON.stringify(payload.documents_urls);
      }
    } catch {
      documents_urls = "[]";
    }

    const row: Record<string, any> = {
      first_name: payload.first_name,
      last_name: payload.last_name,
      national_id: payload.national_id,
      father_name: payload.father_name,
      mother_name: payload.mother_name,
      birth_place: payload.birth_place,
      birth_date: payload.birth_date,
      gender: payload.gender,
      education: payload.education,
      work_province: payload.work_province,
      work_district: payload.work_district,
      institution_name: payload.institution_name,
      work_unit: payload.work_unit,
      work_unit_address: payload.work_unit_address ?? null,
      corp_reg_no: payload.corp_reg_no,
      title: payload.title,
      blood_type: payload.blood_type,
      retirement_no: payload.retirement_no ?? null,
      ssk_no: payload.ssk_no ?? null,
      email: payload.email,
      phone: payload.phone,
      // Create application as pending; member_no will be assigned upon approval
      status: "pending",
      member_no: null,
      join_date: null,
      leave_date: null,
      photo_url: payload.photo_url ?? null,
      documents_urls,
    };

    let { error } = await supabase.from("members").insert(row);
    if (error) {
      const msg = (error.message || "") + " " + (error.details || "");
      const code = String((error as any).code || "");
      if ((code === "42703" || /undefined column|does not exist/i.test(msg)) && /work_unit_address/.test(msg)){
        try{
          delete row.work_unit_address;
          const r2 = await supabase.from("members").insert(row);
          error = r2.error || null;
        }catch{}
      }
      if (error) {
        return new Response(
          JSON.stringify({ error: "insert_failed", details: error }),
          { status: 400, headers }
        );
      }
    }

    return new Response(JSON.stringify({ ok: true }), { status: 200, headers });
  } catch (e) {
    return new Response(
      JSON.stringify({ error: "server_error", details: String(e?.message || e) }),
      { status: 500, headers }
    );
  }
});
