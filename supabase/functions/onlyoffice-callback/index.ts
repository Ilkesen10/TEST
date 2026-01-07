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
    const url = new URL(req.url);
    const bucket = url.searchParams.get("bucket") || "docs";
    const key = url.searchParams.get("key") || "";
    if (!key) {
      return new Response(JSON.stringify({ error: "missing_key" }), { status: 400, headers });
    }

    // Read ONLYOFFICE callback payload
    const payload = await req.json().catch(() => ({}));
    const status = Number(payload?.status ?? 0);
    const fileUrl: string | undefined = payload?.url || payload?.fileUrl || payload?.downloadUrl;

    // ONLYOFFICE spec: status >= 2 means something changed and a URL is provided
    if (!fileUrl || !/^https?:\/\//i.test(fileUrl)) {
      return new Response(JSON.stringify({ error: 0, note: "no_file_url", status }), { status: 200, headers });
    }

    const SUPABASE_URL = Deno.env.get("SUPABASE_URL");
    const SERVICE_ROLE = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    const DS_URL = Deno.env.get("DOCUMENT_SERVER_URL") || Deno.env.get("ONLYOFFICE_SERVER_URL") || "";
    if (!SUPABASE_URL || !SERVICE_ROLE) {
      return new Response(
        JSON.stringify({ error: "server_misconfigured", hint: "Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY" }),
        { status: 500, headers },
      );
    }

    const supabase = createClient(SUPABASE_URL, SERVICE_ROLE);

    // Download updated document from ONLYOFFICE DS
    let downloadUrl = String(fileUrl);
    let fileResp: Response | null = null;
    try { fileResp = await fetch(downloadUrl); } catch { fileResp = null; }
    if (!fileResp || !fileResp.ok) {
      let alt = "";
      try {
        if (DS_URL) {
          if (/^https?:\/\//i.test(downloadUrl)) {
            const fu = new URL(downloadUrl);
            const du = new URL(DS_URL);
            fu.protocol = du.protocol;
            fu.host = du.host;
            alt = fu.href;
          } else {
            alt = new URL(downloadUrl, DS_URL).href;
          }
        }
      } catch {}
      if (alt) {
        try { fileResp = await fetch(alt); } catch { fileResp = null; }
      }
    }
    if (!fileResp || !fileResp.ok) {
      return new Response(
        JSON.stringify({ error: "download_failed", status: fileResp ? fileResp.status : 0 }),
        { status: 502, headers },
      );
    }
    const arr = new Uint8Array(await fileResp.arrayBuffer());

    // Upload to Supabase Storage (upsert into same key)
    const { error: upErr } = await supabase.storage
      .from(bucket)
      .upload(key, arr, {
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        upsert: true,
        cacheControl: "3600",
      });
    if (upErr) {
      return new Response(
        JSON.stringify({ error: "upload_failed", details: upErr }),
        { status: 500, headers },
      );
    }

    return new Response(JSON.stringify({ error: 0 }), { status: 200, headers });
  } catch (e) {
    return new Response(
      JSON.stringify({ error: "server_error", details: String(e?.message || e) }),
      { status: 500, headers },
    );
  }
});
